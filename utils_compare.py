"""
utils.py — PDF extraction + Gemini AI contract comparison
Uses the NEW google.genai SDK (google-genai package)
"""
import io
import re
import json
import streamlit as st
from google import genai
from google.genai import types

from gemini_client import _get_client, detect_available_model, validate_api_key, _FALLBACK_MODELS


# ─── JSON Schema Definition ───────────────────────────────────────────────────
_policy_schema = types.Schema(
    type=types.Type.OBJECT,
    properties={
        "topic": types.Schema(type=types.Type.STRING),
        "contract_1": types.Schema(type=types.Type.ARRAY, items=types.Schema(type=types.Type.STRING)),
        "contract_2": types.Schema(type=types.Type.ARRAY, items=types.Schema(type=types.Type.STRING)),
        "diff_summary": types.Schema(type=types.Type.STRING),
    },
    required=["topic", "contract_1", "contract_2", "diff_summary"]
)

_response_schema = types.Schema(
    type=types.Type.OBJECT,
    properties={
        "step_by_step_analysis": types.Schema(type=types.Type.ARRAY, items=types.Schema(type=types.Type.STRING)),
        "hotel_name": types.Schema(type=types.Type.STRING),
        "year_1": types.Schema(type=types.Type.STRING),
        "year_2": types.Schema(type=types.Type.STRING),
        "seasons": types.Schema(
            type=types.Type.ARRAY,
            items=types.Schema(
                type=types.Type.OBJECT,
                properties={
                    "season_name": types.Schema(type=types.Type.STRING),
                    "period_1": types.Schema(type=types.Type.STRING),
                    "period_2": types.Schema(type=types.Type.STRING),
                    "conditions": types.Schema(type=types.Type.ARRAY, items=types.Schema(type=types.Type.STRING)),
                    "rooms": types.Schema(
                        type=types.Type.ARRAY,
                        items=types.Schema(
                            type=types.Type.OBJECT,
                            properties={
                                "room_name": types.Schema(type=types.Type.STRING),
                                "price_1": types.Schema(type=types.Type.STRING),
                                "price_2": types.Schema(type=types.Type.STRING),
                            },
                            required=["room_name", "price_1", "price_2"]
                        )
                    )
                },
                required=["season_name", "period_1", "period_2", "conditions", "rooms"]
            )
        ),
        "extra_bed": types.Schema(type=types.Type.ARRAY, items=_policy_schema),
        "early_bird": types.Schema(type=types.Type.ARRAY, items=_policy_schema),
        "bonus_night": types.Schema(type=types.Type.ARRAY, items=_policy_schema),
        "wellbeing": types.Schema(type=types.Type.ARRAY, items=_policy_schema),
        "cancellation": types.Schema(type=types.Type.ARRAY, items=_policy_schema),
    },
    required=[
        "step_by_step_analysis", "hotel_name", "year_1", "year_2", 
        "seasons", "extra_bed", "early_bird", "bonus_night", "wellbeing", "cancellation"
    ]
)


# ─── Build prompt ─────────────────────────────────────────────────────────────
def _build_prompt() -> str:
    
    import os
    # Try different fallback paths depending on deployment structure
    base_dir = os.path.dirname(os.path.dirname(__file__))
    current_dir = os.path.dirname(__file__)
    prompt_path_1 = os.path.join(base_dir, "Recheck excel data", "prompts", "contract_compare.txt")
    prompt_path_2 = os.path.join(base_dir, "prompts", "contract_compare.txt")
    prompt_path_3 = os.path.join(current_dir, "prompts", "contract_compare.txt")
    
    if os.path.exists(prompt_path_1):
        with open(prompt_path_1, "r", encoding="utf-8") as f:
            return f.read()
    elif os.path.exists(prompt_path_2):
        with open(prompt_path_2, "r", encoding="utf-8") as f:
            return f.read()
    elif os.path.exists(prompt_path_3):
        with open(prompt_path_3, "r", encoding="utf-8") as f:
            return f.read()
    else:
        return "ERROR: contract_compare.txt not found"



# ─── Streaming comparison ─────────────────────────────────────────────────────
def stream_contract_comparison(pdf1_bytes: bytes, pdf2_bytes: bytes, api_key: str):
    """
    Generator yielding text chunks from Gemini as they stream in.
    Auto-detects the best available model for this API key.
    """
    client = _get_client(api_key)
    prompt = _build_prompt()

    config = types.GenerateContentConfig(
        temperature=0.1,
        max_output_tokens=65536,
        response_mime_type="application/json",
        response_schema=_response_schema,
    )

    # Convert PDFs to native Gemini Parts
    pdf1_part = types.Part.from_bytes(data=pdf1_bytes, mime_type="application/pdf")
    pdf2_part = types.Part.from_bytes(data=pdf2_bytes, mime_type="application/pdf")
    
    # Construct multi-modal payload
    contents = [
        "Please analyze the following two hotel contracts.\n"
        "Contract 1 is the FIRST document. Contract 2 is the SECOND document. Never mix them.\n",
        "\n\n--- CONTRACT 1 (Previous Year) ---\n",
        pdf1_part,
        "\n\n--- CONTRACT 2 (New Year) ---\n",
        pdf2_part,
        "\n\n--- INSTRUCTIONS ---\n",
        prompt
    ]

    # Auto-detect best model
    best_model, all_models = detect_available_model(api_key)
    if not best_model:
        # Last resort: try every fallback
        models_to_try = _FALLBACK_MODELS
    else:
        # Try best first, then the rest of the detected list as fallback
        others = [m for m in all_models if m != best_model]
        models_to_try = [best_model] + others + _FALLBACK_MODELS

    # Deduplicate while preserving order
    seen = set()
    unique_models = []
    for m in models_to_try:
        if m not in seen:
            seen.add(m)
            unique_models.append(m)

    all_errors: list[str] = []
    for model_name in unique_models:
        try:
            for chunk in client.models.generate_content_stream(
                model=model_name,
                contents=contents,
                config=config,
            ):
                if chunk.text:
                    yield chunk.text
            return  # success — stop trying other models
        except Exception as e:
            err = str(e)
            all_errors.append(f"{model_name}: {err}")
            # If we failed mid-stream, tell the UI to reset its buffer before the next model starts
            yield "[RESET_STREAM]"
            continue

    # If we get here, ALL models failed
    summary = " | ".join(all_errors)
    safe = summary.replace('"', "'")
    yield f'{{"error":"All models failed. Details: {safe}"}}'





# ─── Non-streaming alias (backward compatible) ────────────────────────────────
def run_contract_comparison(pdf1_bytes: bytes, pdf2_bytes: bytes, api_key: str) -> str:
    """Collects all streaming chunks and returns the full JSON string."""
    return "".join(stream_contract_comparison(pdf1_bytes, pdf2_bytes, api_key))
