import os
"""
utils.py — Business logic สำหรับ Hotel Audit Desk

Functions หลัก:
  stream_recheck_analysis()   : Audit Excel vs PDF (streaming)
  extract_pdf_to_excel_json() : แปลง PDF → JSON rows [IN TEST]
  create_upload_excel()       : แปลง JSON rows → .xlsx [IN TEST]
  convert_excel_to_csv_with_letters() : helper สำหรับ token optimization

ต้องการ: google-genai, pandas, openpyxl
"""
import io
import time
import pandas as pd
import streamlit as st
import json
import re
from datetime import datetime
from google import genai
from google.genai import types


from gemini_client import _get_client, detect_available_model, validate_api_key, _FALLBACK_MODELS

def _excel_col_name(n):
    """Convert 0-indexed column number to Excel letter (0 -> A, 1 -> B, 26 -> AA)."""
    name = ""
    while n >= 0:
        name = chr(n % 26 + 65) + name
        n = n // 26 - 1
    return name

def convert_excel_to_csv_with_letters(excel_bytes: bytes, focus_list: list = None) -> str:
    """Read Excel, drop empty/unused columns, and return CSV string where headers include Excel column letters."""
    # sheet_name=0 reads the first sheet explicitly
    df = pd.read_excel(io.BytesIO(excel_bytes), header=None, sheet_name=0)
    if df.empty: return ""
    
    letters = [_excel_col_name(i) for i in range(len(df.columns))]
    headers = [f"{letters[i]}: {str(val)}" if pd.notna(val) else letters[i] for i, val in enumerate(df.iloc[0])]
    df.columns = headers
    df = df.iloc[1:].reset_index(drop=True)
    
    # --- Token Optimizations ---
    df = df.dropna(axis=0, how='all')
    df = df.dropna(axis=1, how='all')
    
    # If focus mode is active, we can drop EVERYTHING not related to the focus
    if focus_list and "All-in-One Full Scan" not in focus_list:
        needed_prefixes = {'AL', 'M'} # Always keep Room Name and Hotel Supplier
        mapping = {
            "Net Price & Extra Beds": ['Q', 'AE', 'AF', 'AG', 'AI', 'AJ'],
            "Cancellation Policy": ['AA', 'G', 'H'],
            "Child Policy": ['AD', 'AE', 'AF', 'AG'],
            "Period & Seasons": ['G', 'H', 'K', 'O', 'P', 'U', 'V', 'W'],
            "Meals & Info": ['AO', 'AI', 'AJ']
        }
        for f in focus_list:
            if f in mapping:
                needed_prefixes.update(mapping[f])
        
        cols_to_keep = []
        for col in df.columns:
            prefix = col.split(':')[0].strip()
            if prefix in needed_prefixes:
                cols_to_keep.append(col)
        df = df[cols_to_keep]
    else:
        # Default: Drop Col J (ABF) since the prompt says "ไม่ต้องตรวจสอบ"
        cols_to_keep = [c for c in df.columns if not c.startswith('J:')]
        df = df[cols_to_keep]
    
    return df.to_csv(index=False)

def get_recheck_prompt(focus_list: list = None) -> str:
    """
    สร้าง prompt สำหรับ Gemini ในการ audit Excel vs PDF
    Input:  focus_list = รายการ scope เช่น ["Net Price", "Cancellation Policy"]
                         None หรือมี "All-in-One Full Scan" = ตรวจทุก column
    Output: prompt string พร้อม column mapping และ output format
    """
    focus_instr = ""
    if focus_list and "All-in-One Full Scan" not in focus_list:
        focus_instr = f"\n\n[SYSTEM ALERT: STRICT FOCUS MODE ACTIVE]\nYour primary objective is EXCLUSIVELY to audit: **{', '.join(focus_list)}**.\nIgnore other columns unless they are critical for mapping the selected areas. Do not report issues outside these focus areas.\n\n"
    prompt_path = os.path.join(os.path.dirname(__file__), "prompts", "recheck_audit.txt")
    with open(prompt_path, "r", encoding="utf-8") as f:
        prompt_template = f.read()
    
    # We must format the focus_instr into the prompt manually since we aren't using f-strings natively
    return prompt_template.replace("{focus_instr}", focus_instr)


def stream_recheck_analysis(pdf_bytes: bytes, excel_bytes: bytes, api_key: str, focus_list: list = None):
    """Streams the analysis from Gemini."""
    client = _get_client(api_key)
    prompt = get_recheck_prompt(focus_list)
    
    # Convert Excel to CSV
    csv_data = convert_excel_to_csv_with_letters(excel_bytes, focus_list)
    
    config = types.GenerateContentConfig(
        temperature=0.2,
        max_output_tokens=32768,  # 32K is enough for audit reports
    )
    
    pdf_part = types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf")
    
    contents = [
        "Please act as the Data Recheck Auditor. I am providing you with two documents:\n",
        "1. The original Hotel Contract (PDF)\n",
        pdf_part,
        "\n2. The Data Team's Excel File (converted to CSV with Excel Column Letters in headers)\n",
        csv_data,
        "\n\n--- INSTRUCTIONS & RULES ---\n",
        prompt
    ]
    
    best_model, all_models = detect_available_model(api_key)
    if not best_model:
        models_to_try = _FALLBACK_MODELS
    else:
        others = [m for m in all_models if m != best_model]
        models_to_try = [best_model] + others + _FALLBACK_MODELS

    # Deduplicate
    seen = set()
    unique_models = []
    for m in models_to_try:
        if m not in seen:
            seen.add(m)
            unique_models.append(m)

    all_errors = []
    max_retries_per_run = 2  # Try the whole list twice if needed
    
    for attempt in range(max_retries_per_run):
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
                
                if "429" in err or "RESOURCE_EXHAUSTED" in err:
                    yield f"\n\n[SYSTEM] Model {model_name} quota full (429). Trying next...\n"
                    time.sleep(2)
                elif "400" in err and "token count" in err.lower():
                    yield f"\n\n[SYSTEM] Data too large for {model_name} (400). Trying next...\n"
                elif "404" in err:
                    pass  # Just move to next model silently
                
                yield "[RESET_STREAM]"
                continue
        
        if attempt < max_retries_per_run - 1:
            yield "\n\n[SYSTEM] All models busy. Waiting 10s before retrying...\n"
            time.sleep(10)

    # If all fail after all attempts — show clean user-friendly message
    # Classify what went wrong
    quota_errors = [e for e in all_errors if "429" in e or "RESOURCE_EXHAUSTED" in e]
    invalid_errors = [e for e in all_errors if "INVALID_ARGUMENT" in e or "API_KEY_INVALID" in e]
    not_found_errors = [e for e in all_errors if "404" in e or "NOT_FOUND" in e]

    if invalid_errors:
        yield (
            "\n\n**API Key ไม่ถูกต้อง**\n\n"
            "กรุณาตรวจสอบ API Key ในแถบด้านซ้าย:\n"
            "- ต้องขึ้นต้นด้วย `AIza`\n"
            "- ไม่มีช่องว่างหรืออักขระพิเศษ\n"
            "- ได้รับจาก [Google AI Studio](https://aistudio.google.com/app/apikey) เท่านั้น"
        )
    elif quota_errors and not_found_errors:
        yield (
            "\n\n**โควต้า API หมดสำหรับวันนี้**\n\n"
            "โมเดล AI ทุกตัวที่มีสิทธิ์ใช้งานถูกใช้ครบโควต้ารายวันแล้วครับ\n\n"
            "**วิธีแก้ไข:**\n"
            "1. รอถึงพรุ่งนี้ (โควต้าจะ Reset อัตโนมัติ)\n"
            "2. ใช้ API Key จาก Gmail บัญชีอื่น\n"
            "3. เปิดใช้ Pay-as-you-go บน Google AI Studio เพื่อปลดล็อกโควต้า"
        )
    elif quota_errors:
        yield (
            "\n\n**โควต้า API หมดชั่วคราว (429)**\n\n"
            "กรุณารอ 1-2 นาที แล้วลองใหม่อีกครั้งครับ"
        )
    else:
        yield (
            "\n\n**ไม่สามารถเชื่อมต่อ AI ได้**\n\n"
            "API Key นี้อาจไม่รองรับโมเดลที่จำเป็น\n"
            "ลองสร้าง API Key ใหม่จาก Google AI Studio ครับ"
        )


# ─── EXCEL GENERATION ENGINE [IN TEST] ────────────────────────────────────────

