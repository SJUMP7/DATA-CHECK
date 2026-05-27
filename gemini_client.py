"""
gemini_client.py - Manages the Gemini API client and model detection.
"""
import streamlit as st
from google import genai

# ─── Fallback model list ───────────────────────────────────────
_FALLBACK_MODELS = [
    "gemini-2.0-flash",
    "gemini-2.0-flash-lite",
    "gemini-1.5-flash",
    "gemini-1.5-flash-latest",
    "gemini-2.0-flash-001",
    "gemini-2.0-flash-lite-001",
    "gemini-1.5-pro",
    "gemini-1.5-pro-latest",
]

@st.cache_resource(show_spinner=False, ttl=3600)
def _get_client(api_key: str) -> genai.Client:
    """Create Gemini client. TTL=1hr prevents stale connections on repeated audits."""
    return genai.Client(api_key=api_key)

@st.cache_data(show_spinner=False, ttl=600)  # Cache model list for 10 min to save quota
def _cached_model_list(api_key: str) -> list[str]:
    """Returns list of available model short-names. Cached 10 min to avoid wasting RPM."""
    client = _get_client(api_key)
    available: list[str] = []
    try:
        for m in client.models.list():
            name = getattr(m, "name", "") or ""
            short = name.replace("models/", "")
            available.append(short)
    except Exception as e:
        print(f"[DEBUG] _cached_model_list error: {e}")
    return available

def detect_available_model(api_key: str) -> tuple[str, list[str]]:
    """Uses cached model list to avoid burning quota on every call."""
    all_names = _cached_model_list(api_key)
    if not all_names:
        return "", []  # API Key is invalid or network error

    # Filter out non-text / preview / low-quota / specialized models
    skip_keywords = [
        "tts", "audio", "vision", "embedding", "tuning",
        "research", "lyria", "live",
        "image",    # image-gen models have tiny quota
        "preview",  # preview = unstable / low quota
        "think",    # thinking models need special config
        "exp",      # experimental = unreliable quota
        "3.1",      # Gemini 3.x has very low free-tier quota
        "3.0",      # same
    ]
    available = [
        short for name in all_names
        for short in [name.replace("models/", "")]
        if any(k in short for k in ["flash", "pro"])
        and not any(skip in short.lower() for skip in skip_keywords)
    ]

    # Prefer stable, high-quota models for large JSON generation
    preferred_order = [
        "gemini-2.5-flash",
        "gemini-3.5-flash",
        "gemini-flash-latest",
        "gemini-1.5-flash",
        "gemini-1.5-flash-latest",
        "gemini-2.0-flash",
        "gemini-2.0-flash-lite",
        "gemini-2.0-flash-001",
    ]
    for preferred in preferred_order:
        if preferred in available:
            return preferred, available

    for m in _FALLBACK_MODELS:
        if m in available:
            return m, available

    if available:
        return available[0], available
    return "", []

def validate_api_key(api_key: str) -> tuple[bool, str]:
    if not api_key: return False, "No API key provided."
    model, available = detect_available_model(api_key)
    if model: return True, f"CONNECTED - MODEL: {model.upper()}"
    if available: return True, f"CONNECTED (MODELS FOUND: {len(available)})"
    return False, "API KEY INVALID OR GEMINI API NOT ENABLED."
