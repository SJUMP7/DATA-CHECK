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
        "other_promotions": types.Schema(type=types.Type.ARRAY, items=_policy_schema),
    },
    required=[
        "step_by_step_analysis", "hotel_name", "year_1", "year_2", 
        "seasons", "extra_bed", "early_bird", "bonus_night", "wellbeing", "cancellation", "other_promotions"
    ]
)

_revise_policy_schema = types.Schema(
    type=types.Type.OBJECT,
    properties={
        "topic": types.Schema(type=types.Type.STRING),
        "contract_1": types.Schema(type=types.Type.ARRAY, items=types.Schema(type=types.Type.STRING)),
        "contract_2": types.Schema(type=types.Type.ARRAY, items=types.Schema(type=types.Type.STRING)),
        "contract_3": types.Schema(type=types.Type.ARRAY, items=types.Schema(type=types.Type.STRING)),
        "diff_summary": types.Schema(type=types.Type.STRING),
        "diff_summary_2": types.Schema(type=types.Type.STRING),
    },
    required=["topic", "contract_1", "contract_2", "contract_3", "diff_summary", "diff_summary_2"]
)

_revise_response_schema = types.Schema(
    type=types.Type.OBJECT,
    properties={
        "step_by_step_analysis": types.Schema(type=types.Type.ARRAY, items=types.Schema(type=types.Type.STRING)),
        "hotel_name": types.Schema(type=types.Type.STRING),
        "year_1": types.Schema(type=types.Type.STRING),
        "year_2": types.Schema(type=types.Type.STRING),
        "year_3": types.Schema(type=types.Type.STRING),
        "seasons": types.Schema(
            type=types.Type.ARRAY,
            items=types.Schema(
                type=types.Type.OBJECT,
                properties={
                    "season_name": types.Schema(type=types.Type.STRING),
                    "period_1": types.Schema(type=types.Type.STRING),
                    "period_2": types.Schema(type=types.Type.STRING),
                    "period_3": types.Schema(type=types.Type.STRING),
                    "conditions": types.Schema(type=types.Type.ARRAY, items=types.Schema(type=types.Type.STRING)),
                    "rooms": types.Schema(
                        type=types.Type.ARRAY,
                        items=types.Schema(
                            type=types.Type.OBJECT,
                            properties={
                                "room_name": types.Schema(type=types.Type.STRING),
                                "price_1": types.Schema(type=types.Type.STRING),
                                "price_2": types.Schema(type=types.Type.STRING),
                                "price_3": types.Schema(type=types.Type.STRING),
                            },
                            required=["room_name", "price_1", "price_2", "price_3"]
                        )
                    )
                },
                required=["season_name", "period_1", "period_2", "period_3", "conditions", "rooms"]
            )
        ),
        "extra_bed": types.Schema(type=types.Type.ARRAY, items=_revise_policy_schema),
        "early_bird": types.Schema(type=types.Type.ARRAY, items=_revise_policy_schema),
        "bonus_night": types.Schema(type=types.Type.ARRAY, items=_revise_policy_schema),
        "wellbeing": types.Schema(type=types.Type.ARRAY, items=_revise_policy_schema),
        "cancellation": types.Schema(type=types.Type.ARRAY, items=_revise_policy_schema),
        "other_promotions": types.Schema(type=types.Type.ARRAY, items=_revise_policy_schema),
    },
    required=[
        "step_by_step_analysis", "hotel_name", "year_1", "year_2", "year_3",
        "seasons", "extra_bed", "early_bird", "bonus_night", "wellbeing", "cancellation", "other_promotions"
    ]
)


def clean_date_period_format(text: str) -> str:
    """
    Standardizes date strings (like seasons periods) to use 3-letter uppercase English months
    and 2-digit years. E.g.:
    "1 November 2024 - 23 December 2024" -> "1 NOV 24 - 23 DEC 24"
    "1 JAN 26 - 31 DEC 26" -> "1 JAN 26 - 31 DEC 26"
    "N/A" -> "N/A"
    """
    if not text:
        return ""
    text_str = str(text).upper().strip()
    if text_str in ("N/A", "NA", "-", ""):
        return "N/A"
        
    months = {
        'JANUARY': 'JAN', 'FEBRUARY': 'FEB', 'MARCH': 'MAR', 'APRIL': 'APR', 'MAY': 'MAY',
        'JUNE': 'JUN', 'JULY': 'JUL', 'AUGUST': 'AUG', 'SEPTEMBER': 'SEP', 'OCTOBER': 'OCT',
        'NOVEMBER': 'NOV', 'DECEMBER': 'DEC',
        'JAN': 'JAN', 'FEB': 'FEB', 'MAR': 'MAR', 'APR': 'APR', 'JUN': 'JUN', 'JUL': 'JUL',
        'AUG': 'AUG', 'SEP': 'SEP', 'OCT': 'OCT', 'NOV': 'NOV', 'DEC': 'DEC'
    }
    
    # Replace full month names and abbreviations with standard uppercase abbreviations
    for full_m, short_m in months.items():
        text_str = re.sub(rf'\b{full_m}\b', short_m, text_str)
        
    # Standardize 4-digit years to 2-digit years (e.g. 2024 -> 24, 2025 -> 25)
    text_str = re.sub(r'\b20(\d{2})\b', r'\1', text_str)
    
    # Standardize spaces around dashes
    text_str = re.sub(r'\s*-\s*', ' - ', text_str)
    
    # Remove excessive whitespace
    text_str = re.sub(r'\s+', ' ', text_str)
    
    return text_str.strip()


# ─── Build prompt ─────────────────────────────────────────────────────────────
def _build_prompt(use_revise: bool = False) -> str:
    import os
    current_dir = os.path.dirname(os.path.abspath(__file__))
    filename = "contract_compare_revise.txt" if use_revise else "contract_compare.txt"
    possible_paths = [
        os.path.join(current_dir, "prompts", filename),
        os.path.join(current_dir, "Recheck excel data", "prompts", filename),
        os.path.join(os.path.dirname(current_dir), "prompts", filename),
        os.path.join(os.getcwd(), "prompts", filename),
        os.path.join(os.getcwd(), "Recheck excel data", "prompts", filename)
    ]
    
    for p in possible_paths:
        if os.path.exists(p):
            with open(p, "r", encoding="utf-8") as f:
                return f.read()
                
    return f"ERROR: {filename} not found in any expected paths."



# ─── Streaming comparison ─────────────────────────────────────────────────────
def stream_contract_comparison(pdf1_bytes: bytes, pdf2_bytes: bytes, api_key: str, pdf3_bytes: bytes = None):
    """
    Generator yielding text chunks from Gemini as they stream in.
    Auto-detects the best available model for this API key.
    If pdf3_bytes is provided, performs a 3-contract comparison (Previous, New, Revise).
    """
    client = _get_client(api_key)
    use_revise = pdf3_bytes is not None
    prompt = _build_prompt(use_revise=use_revise)

    config = types.GenerateContentConfig(
        temperature=0.1,
        max_output_tokens=65536,
        response_mime_type="application/json",
        response_schema=_revise_response_schema if use_revise else _response_schema,
    )

    # Convert PDFs to native Gemini Parts
    pdf1_part = types.Part.from_bytes(data=pdf1_bytes, mime_type="application/pdf")
    pdf2_part = types.Part.from_bytes(data=pdf2_bytes, mime_type="application/pdf")
    
    # Construct multi-modal payload
    if use_revise:
        pdf3_part = types.Part.from_bytes(data=pdf3_bytes, mime_type="application/pdf")
        contents = [
            "Please analyze the following three hotel contracts.\n"
            "Contract 1 is the FIRST document. Contract 2 is the SECOND document. Contract 3 is the THIRD document. Never mix them.\n",
            "\n\n--- CONTRACT 1 (Previous Year) ---\n",
            pdf1_part,
            "\n\n--- CONTRACT 2 (New Year) ---\n",
            pdf2_part,
            "\n\n--- CONTRACT 3 (New Year Revise Contract) ---\n",
            pdf3_part,
            "\n\n--- INSTRUCTIONS ---\n",
            prompt
        ]
    else:
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
            # Bug #7 Fix: yield RESET_STREAM ก่อน เพื่อให้ caller ล้าง partial JSON buffer
            # ก่อนที่ model ถัดไปจะเริ่ม stream ใหม่
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
