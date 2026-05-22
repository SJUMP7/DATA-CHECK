import os
import io
import time
import json
import re
import pandas as pd
from datetime import datetime
from google.genai import types
from gemini_client import _get_client, detect_available_model, _FALLBACK_MODELS

# ─── EXCEL SCHEMA DEFINITION [IN TEST] ────────────────────────────────────────
# This list defines the exact column order required for the system upload.
EXCEL_UPLOAD_COLUMNS = [
    '_id', 'hotel_id', 'room_id', 'status', 'created_date', 'edited_date',
    'start_date', 'end_date', 'refundable', 'abf', 'contract_type', 'cutoff_date',
    'hotel_supplier', 'important_message', 'min_nights_stay', 'min_advance_days',
    'net_price', 'net_price_emerald', 'net_price_ruby', 'net_price_topaz',
    'promo_book_till', 'promo_code', 'promo_note', 'room_allotment', 'all_inclusive',
    'baby_cot', 'cancellation_policy', 'cancellation_policy_net', 'early_check_in',
    'child_policy', 'child_share_bed_abf', 'child_extra_bed_abf', 'extra_bed_abf',
    'extra_bed_no_abf', 'full_board', 'half_board', 'hotel_extra_fees', 'room_name',
    'hotel_transfer', 'late_check_out', 'meals_and_info', 'tags', 'action'
]

def _build_extraction_prompt(extra_instruction: str = "", reference_csv: str = "") -> str:
    """Builds the main extraction prompt for PDF → JSON conversion."""
    
    reference_section = ""
    if reference_csv:
        # Show only the first ~30 rows of the reference CSV to save tokens
        csv_lines = reference_csv.strip().splitlines()
        preview_lines = csv_lines[:31]  # header + up to 30 data rows
        reference_section = f"""
REFERENCE EXCEL (EXISTING DATA — USE AS FORMAT EXAMPLE):
Below is the existing Excel file for this hotel, converted to CSV.
Use it to:
  1. Understand the exact format expected for each column (dates, HTML structure, etc.)
  2. Infer how many rows a contract of this size should produce.
  3. Match your output style to what is shown below.

{chr(10).join(preview_lines)}
--- END OF REFERENCE ---
"""

    
    prompt_path = os.path.join(os.path.dirname(__file__), "prompts", "excel_generator.txt")
    with open(prompt_path, "r", encoding="utf-8") as f:
        prompt_template = f.read()
    
    return prompt_template.format(
        extra_instruction=extra_instruction,
        reference_section=reference_section,
        json_keys=json.dumps(EXCEL_UPLOAD_COLUMNS[6:])
    )



def extract_pdf_to_excel_json(pdf_bytes: bytes, api_key: str, excel_bytes: bytes = None):
    """
    [IN TEST] — ยังไม่ production-ready
    Extracts every room/season/promotion combination from the PDF into a JSON list.
    If excel_bytes is provided, uses the existing Excel as a format reference and
    derives a dynamic row count threshold from the actual data (no hardcoded minimum).
    """
    client = _get_client(api_key)

    # ── Parse existing Excel for reference context (if available) ─────────────
    reference_csv = ""
    expected_rows = None  # None = unknown; skip row count validation
    if excel_bytes:
        try:
            ref_df = pd.read_excel(io.BytesIO(excel_bytes), header=0, sheet_name=0)
            ref_df = ref_df.dropna(how='all')

            # Detect multi-hotel file: if >1 unique hotel_supplier → don't use row count
            # (the file is a combined export, not a single-hotel contract)
            if 'hotel_supplier' in ref_df.columns:
                unique_hotels = ref_df['hotel_supplier'].dropna().nunique()
                if unique_hotels == 1:
                    # Single hotel → row count is meaningful for validation
                    expected_rows = max(1, len(ref_df))
                # else: multi-hotel → keep expected_rows = None, skip validation

            # Always build format reference CSV (helps AI learn column format)
            # We select a diverse subset of rows (e.g. 1 sample row per contract_type/room_name combination)
            # to show the model how different contract types look without overwhelming it with repetitive rows.
            lean_cols = [c for c in ref_df.columns if c not in [
                'cancellation_policy', 'cancellation_policy_net',
                'child_policy', 'meals_and_info'
            ]]
            
            # Find unique combination examples
            group_keys = []
            if 'contract_type' in ref_df.columns: group_keys.append('contract_type')
            if 'room_name' in ref_df.columns: group_keys.append('room_name')
            
            if group_keys:
                # Group and take the first row of each group
                sample_df = ref_df.groupby(group_keys, as_index=False).first()
                # If still too large, limit to 8 rows
                sample_df = sample_df[lean_cols].head(8)
            else:
                sample_df = ref_df[lean_cols].head(5)
                
            reference_csv = sample_df.to_csv(index=False)
        except Exception as e:
            print(f"[DEBUG] Excel reference error: {e}")  # If Excel can't be read, proceed without reference

    # ── Build model list ───────────────────────────────────────────────────────
    available_models_list = []
    try:
        best_model, all_models = detect_available_model(api_key)
        others = [m for m in all_models if m != best_model]
        models_to_try = [best_model] + others + _FALLBACK_MODELS
        seen = set()
        for m in models_to_try:
            if m not in seen:
                seen.add(m)
                available_models_list.append(m)
    except Exception as e:
        print(f"[DEBUG] _FALLBACK_MODELS error: {e}")
        available_models_list = _FALLBACK_MODELS

    # ── Config: no response_mime_type so the model doesn't suppress repeated rows ──
    config = types.GenerateContentConfig(
        temperature=0.1,
        max_output_tokens=65536,
    )

    def _call_model(model_name: str, extra_instruction: str = "") -> tuple:
        """Returns (data_list, error_string). data_list is [] on failure."""
        prompt_text = _build_extraction_prompt(extra_instruction, reference_csv)
        contents = [
            types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
            prompt_text,
        ]
        full_text = ""
        for chunk in client.models.generate_content_stream(
            model=model_name,
            contents=contents,
            config=config,
        ):
            if chunk.text:
                full_text += chunk.text

        if not full_text:
            return [], "Empty response from model"

        # Robust JSON extraction: find outermost [ ... ] bracket pair
        json_start = full_text.find('[')
        json_end   = full_text.rfind(']')
        if json_start != -1 and json_end != -1 and json_end > json_start:
            clean_json = full_text[json_start:json_end + 1].strip()
        else:
            clean_json = full_text.strip().replace('```json', '').replace('```', '').strip()

        parsed = json.loads(clean_json)   # raises on invalid JSON
        return parsed, None

    last_error = "Unknown Error"

    for model_name in available_models_list:
        try:
            # ── First attempt ─────────────────────────────────────────────────
            data, err = _call_model(model_name)
            if err:
                last_error = f"{model_name}: {err}"
                continue

            # ── Row count validation (only when we have a reference) ─────────
            # If expected_rows is known and AI returned significantly fewer, retry once.
            # A 50% tolerance allows for legitimate differences (e.g., only Main Contract,
            # no promos). A single-row promo contract is always accepted as-is.
            if expected_rows is not None and expected_rows > 1:
                threshold = max(1, int(expected_rows * 0.5))
                if len(data) < threshold:
                    escalation = (
                        f"[WARNING] PREVIOUS ATTEMPT: You only generated {len(data)} rows, "
                        f"but the existing Excel for this hotel has {expected_rows} rows. "
                        f"Go through the PDF again and extract EVERY room type × season × rate type combination. "
                        f"Match the scale of the reference data shown above."
                    )
                    data_retry, err_retry = _call_model(model_name, escalation)
                    if not err_retry and len(data_retry) > len(data):
                        data = data_retry

            return data, None

        except Exception as e:
            err_str = str(e)
            last_error = f"Model {model_name} failed: {err_str}"
            if "429" in err_str:
                time.sleep(1)
            continue

    return [], f"All models exhausted. Last error: {last_error}"

def create_upload_excel(data_list: list):
    """
    Converts extracted JSON list into a formatted Excel file buffer. [IN TEST]
    """
    # Ensure all required columns exist in the DataFrame, even if missing from JSON
    df = pd.DataFrame(data_list)
    for col in EXCEL_UPLOAD_COLUMNS:
        if col not in df.columns:
            df[col] = ""
            
    # Reorder to strict column layout
    df = df[EXCEL_UPLOAD_COLUMNS]
    
    # 1. HARDCODED OVERRIDES (To prevent AI hallucination)
    df['status'] = 'True'
    df['refundable'] = 'False'
    df['abf'] = 'Included'
    df['action'] = 'insert'
    df['tags'] = '[]'
    df['created_date'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Optional logic: if AI left room_allotment empty, set it
    df['room_allotment'] = df['room_allotment'].replace("", "Free Sales")
    
    # ── promo_book_till fallback logic ───────────────────────────────────────
    # If contract_type is "Promotion" or "Early Bird", and promo_book_till is empty/null/missing,
    # set it to the stay's end_date with time set to 23:59:59.
    def fill_promo_book_till(row):
        ctype = str(row.get('contract_type', '')).strip().lower()
        book_till = row.get('promo_book_till')
        if ctype in ['promotion', 'early bird']:
            if pd.isna(book_till) or str(book_till).strip() == "" or str(book_till).strip().lower() in ["none", "null", "nan"]:
                end_dt = str(row.get('end_date', '')).strip()
                if end_dt:
                    date_part = end_dt.split(' ')[0]
                    if re.match(r'^\d{4}-\d{2}-\d{2}$', date_part):
                        return f"{date_part} 23:59:59"
        return book_till

    if 'promo_book_till' in df.columns:
        df['promo_book_till'] = df.apply(fill_promo_book_till, axis=1)

    
    # ── Room name cleanup & dynamic ABF setup ─────────────────────────────────
    # If the extracted room name contains suffix like (RB), (RO), Room with Breakfast, etc.
    # We clean the room_name and map it to the proper 'abf' column.
    def clean_room_and_abf(row):
        rname = str(row['room_name']).strip()
        abf_val = str(row['abf']).strip() if 'abf' in row else 'Included'
        
        # Check for RO / Room Only
        if re.search(r'\b(ro|room\s+only)\b', rname, re.IGNORECASE):
            abf_val = 'Excluded'
        # Check for RB / Breakfast
        elif re.search(r'\b(rb|breakfast|abf)\b', rname, re.IGNORECASE):
            abf_val = 'Included'
            
        # 1. First remove bracketed terms like (RO), (RB), (Room Only), (ABF), etc.
        rname_clean = re.sub(r'[\(\[\{]\s*(ro|rb|room\s+only|breakfast|abf)\s*[\)\]\}]', '', rname, flags=re.IGNORECASE)
        # 2. Then remove standalone words ro, rb, abf, breakfast
        rname_clean = re.sub(r'\b(ro|rb|abf|breakfast)\b', '', rname_clean, flags=re.IGNORECASE)
        # 3. If there is a trailing 'only' (e.g. Deluxe Room Only), remove it
        rname_clean = re.sub(r'\b(only)\b\s*$', '', rname_clean, flags=re.IGNORECASE)
        
        # Clean up trailing/leading dashes, spaces, or empty/unclosed brackets
        rname_clean = re.sub(r'\s*-\s*$', '', rname_clean)
        rname_clean = re.sub(r'\s*[\(\[\{]\s*[\)\]\}]\s*$', '', rname_clean) # remove empty trailing brackets like ()
        rname_clean = re.sub(r'\s*[\(\[\{]\s*$', '', rname_clean) # remove trailing unclosed brackets like (
        rname_clean = re.sub(r'\s+', ' ', rname_clean).strip()
        
        return pd.Series([rname_clean, abf_val])
        
    if 'room_name' in df.columns:
        df[['room_name', 'abf']] = df.apply(clean_room_and_abf, axis=1)
    
    # 2. FORCE EMPTY STRING ON COLUMNS THAT SHOULD NOT HAVE 0
    empty_cols = [
        'net_price_emerald', 'net_price_ruby', 'net_price_topaz',
        'all_inclusive', 'baby_cot', 'cancellation_policy_net',
        'early_check_in', 'extra_bed_no_abf', 'full_board', 'half_board',
        'hotel_extra_fees', 'hotel_transfer', 'late_check_out'
    ]
    for col in empty_cols:
        df[col] = ""
        
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Upload')
    
    return output.getvalue()