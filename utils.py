import io
import time
import pandas as pd
import streamlit as st
import json
import re
from datetime import datetime
from google import genai
from google.genai import types

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

# ─── Fallback model list ───────────────────────────────────────
# Expanded with models from user logs to maximize chances of finding an open quota
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

@st.cache_resource(show_spinner=False)
def _get_client(api_key: str) -> genai.Client:
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
    except Exception:
        pass
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
    if model: return True, f"Connected ✓ Model: {model}"
    if available: return True, f"Connected (models found: {len(available)})"
    return False, "API key invalid or Gemini API not enabled."

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
    focus_instr = ""
    if focus_list and "All-in-One Full Scan" not in focus_list:
        focus_instr = f"\n\n[SYSTEM ALERT: STRICT FOCUS MODE ACTIVE]\nYour primary objective is EXCLUSIVELY to audit: **{', '.join(focus_list)}**.\nIgnore other columns unless they are critical for mapping the selected areas. Do not report issues outside these focus areas.\n\n"

    return f"""
Data Recheck : (100% Full Scan)
{focus_instr}
Role: Senior Auditor. 
Objective: Compare PDF Contract vs Excel Data. 100% Accuracy. 
Strictly use provided documents only. Report discrepancies precisely.

How to check : 
1. อ่านไฟล์ PDF ที่ส่งให้ (100% Full Scan) เพื่อเตรียมไว้ตรวจสอบไฟล์ Excel ที่ถูกทำไว้แล้ว
2. นำข้อมูลในไฟล์ PDF มาตรวจสอบเปรียบเทียบกับไฟล์ Excel (ที่ถูกแปลงเป็น CSV โดยมีชื่อคอลัมน์เช่น G:, H:, Q: กำกับไว้) เพื่อตรวจสอบความถูกต้อง
3. หาก Full Scan และตรวจสอบเสร็จแล้ว ให้รายงานผล ตามรูปแบบที่ตั้งไว้ให้ผู้ใช้งาน อ่านง่าย

Directives:
1. Full Scan: Audit every cell. No sampling.
2. Row Mapping: Map rates (Q, AE-AG) to room type (AL) strictly.
3. Exceptions: Search for "excluding", "not including", "only" in PDF.
4. Child Logic: Under 12 = 11.99, 5-11 = 5-11.99.
5. Cancellation: Match days in AA to seasons in PDF.
6. Discount & Extra Person: For Early Booking/Promotions, ALWAYS check if the contract states "apply with extra person charge" or similar. If YES, Extra Bed (AG) and Child Extra Bed (AF) MUST be discounted. DO NOT flag this as a FAIL.
7. GROUPING (NON-NEGOTIABLE RULE):
   - BEFORE (WRONG - never do this):
     | AO (Room A, Peak) | 3 nights | 7 nights | ... |
     | AO (Room B, Peak) | 3 nights | 7 nights | ... |
     | AO (Room C, Peak) | 3 nights | 7 nights | ... |
   - AFTER (CORRECT - always do this):
     | AO (ทุก Room Type, Peak Season) | 3 nights | 7 nights | ... |
   If 2 or more rows have the SAME column, SAME error, SAME fix — MERGE into ONE row. Write location as "ทุก Room Type" or "Row 5-15" etc.
8. IGNORE FORMATTING & DECIMALS (CRITICAL): Do NOT flag an error if the difference is purely numerical formatting. For example: `1600` is EXACTLY THE SAME as `1600.00`. `7352.5` is EXACTLY THE SAME as `7352.50`. If the factual numeric value is identical, you MUST consider it CORRECT. Do NOT flag it as FAIL. Do NOT ask the user to "ปรับรูปแบบตัวเลขให้เป็นทศนิยม 2 ตำแหน่ง".
9. HTML Code Block Placement (CRITICAL): For policy columns (AA, AD, AO), when you find an error:
   - In the table row, write ONLY a plain-text summary in the "ข้อมูลที่ถูกต้อง" column.
   - AFTER the table ends (not inside a cell!), add the HTML dropdown on its own line like this:

---
HTML สำหรับแก้ไข [Column Name] — [Room/Period]:

<details>
<summary>คลิกเพื่อดูโค้ด HTML (Copy โค้ดด้านล่าง)</summary>

```html
[INSERT FULL HTML CODE HERE]
```

</details>

10. Long Stay Offer Rules: If the file shows Long Stay Offer with MAX 7 NIGHTS (e.g., #MIN. 4 NIGHTS get 10% or #MIN. 7 NIGHTS get 15%), this is the COMPANY'S STANDARD POLICY. Do NOT flag as FAIL. Only flag REVIEW with note "ตรวจสอบ Long Stay Offer กับสัญญาอีกครั้ง" if the number of nights or discount percentage differs from the PDF.
11. MUST COMPLETE: You MUST audit and report ALL columns and ALL rows. Do NOT stop early. Do NOT say "etc" or abbreviate. If data is large, continue until every row is covered.

Process:
1. นำข้อมูลไฟล์ PDF มาตรวจสอบในไฟล์ Excel
2. Column Mapping : ข้อมูลพื้นฐาน Excel (ชื่อคอลัมน์ใน CSV จะมีตัวอักษรกำกับ เช่น G:, H:)
   2.1 Col G (start_date): วันที่เริ่มต้น Period (Format: วัน/เดือน/ปี)
   2.2 Col H (end_date): วันที่สิ้นสุด Period (Format: วัน/เดือน/ปี)
   2.3 Col I (refundable): บังคับเป็น FALSE เท่านั้น (ห้ามตรวจสอบกับสัญญา และห้ามรายงานผลในส่วนนี้เด็ดขาด)
   2.4 Col J (abf): ไม่ต้องตรวจสอบ
   2.5 Col K (contract_type): ระบุค่าเป็นค่าใดค่าหนึ่ง: Main Contract, Early Bird, Promotion, POR
   2.6 Col L (cutoff_date): ตัวเลข ตามสัญญา
   2.7 Col M (hotel_supplier): ชื่อโรงแรม 
   2.8 Col O (min_nights_stay): ตัวเลขจำนวนคืน ตามสัญญา (หากไม่มีเว้นว่าง)
   2.9 Col P (min_advance_days): ตัวเลข ตรวจสอบเฉพาะ Early Bird
   2.10 Col Q (net_price): ตัวเลขตามราคาสัญญา
   2.11 Col U (promo_book_till): วันที่สิ้นสุด Booking (Format: ปี/เดือน/วัน 23:59:59) 
   2.12 Col V (promo_code): Code Promotion / E.B xx DAYS, LONG STAY OFFER
   2.13 Col W (promo_note): ใส่เฉพาะ Condition สำคัญโดยเฉพาะ MIN. xx NIGHTS, COMPULSORY NEW YEAR GALA DINNER, NOT ALLOWED CHECK OUT 
   2.14 Col X (room_allotment): Free Sales, On Request, หรือ ตัวเลข (หากไม่มีเว้นว่าง)
   2.15 Col AA (cancellation_policy): ใช้รูปแบบนี้เท่านั้น (ใช้ • แทน <ul><li> ห้ามใช้ HTML list tags):
         CANCELLATION : LOW/SHOULDER/HIGH/PEAK SEASON OR PERIOD<br>
         • Cancellation made (policy)<br>
         • <span style="color: red; font-weight: bold;">NO-SHOW/Early Check Out</span> : (policy)<br>
         <span style="color: red; font-weight: bold;">*Remark</span> : (ถ้ามี หากไม่มีตัดออก)
         [CONTENT ONLY RULE]: ตรวจสอบเฉพาะ เนื้อหา/ตัวเลข/วัน ว่าถูกต้องตามสัญญา
         ห้าม flag FAIL แม้ format bullet จะต่างกันเล็กน้อย ให้ flag REVIEW แทนหากเนื้อหาถูกแต่ format ต่าง
   2.16 Col AC (cancellation_policy_net): บรรจุข้อมูลการยกเลิกในรูปแบบ HTML (ต้องแกะข้อความภายใน Tags มาตรวจ)
   2.17 Col AD (child_policy): จัดรูปแบบข้อความ HTML ตาม Pattern นี้เป๊ะๆ (รักษาโค้ดสีไว้):
        <span style="color: green; font-weight: bold;">Room rate includes ABF for xx persons</span><br>
        <span style="color: green; font-weight: bold;">Maximum Occupancy : XA / XA+XC</span><br><br>
        (policy)<br><br>
        in case have 2 Children;<br>
        if have <strong>2 children (xx-xx.99 years old)</strong> stay in room, subject to charge as policy below;<br>
        1st child (policy)<br>
        2nd child (policy)<br><br>
        Adult (policy)<br>
        <span style="color: red; font-weight: bold;">*Maximum X Extra bed / *CANNOT ADD EXTRA BED*</span>
   2.18 Col AE (child_share_bed_abf): ระบุเป็นตัวเลขโดยอิงราคา Child Sharing Bed+ABF 
   2.19 Col AF (child_extra_bed_abf): ระบุเป็นตัวเลขโดยอิงราคา Child Extra Bed +ABF 
   2.20 Col AG (extra_bed_abf): ระบุเป็นตัวเลขโดยอิงราคา Adult Extra Bed +ABF 
   2.21 Col AI (full_board): ตัวเลข 
   2.22 Col AJ (half_board): ตัวเลข 
   2.23 Col AL (room_name): ระบุชื่อห้องพักตามสัญญา “ยืดหยุ่นได้”
   2.24 Col AO (meals_and_info): รวมข้อมูลเป็น HTML ตาม Pattern นี้เป๊ะๆ (รักษาโค้ดสีไว้):
        <span style="color: red; font-weight: bold;">#MIN. XX NIGHTS - COMPULSORY CHRISTMAS/NEW YEAR'S GALA DINNER on xx - xx</span><br>
        <span style="color: red; font-weight: bold;">**NOT ALLOWED CHECK OUT on xx**</span> (only peak / period that have conditions)
        <hr>
        <strong>MAIN CONTRACT 25/26 : 1 NOV 25 - 31 OCT 26</strong><br><br>
        <span style="color: orange; font-weight: bold;">[ &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(only for Peak)</span><br>
        <strong>MIN. XX NIGHTS during xx-xx</strong><br>
        <strong>NOT ALLOWED CHECK OUT on xx</strong><br><br>
        ※ <strong>Compulsory</strong> New Year's Gala Dinner on 31 DEC xx<br>
        Adult = xxx THB / Child (age) = xxx THB<br>
        <span style="color: red; font-weight: bold;">*Remark : </span><br>
        <span style="color: orange; font-weight: bold;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;]</span><br><br>
        ※ <strong>MEAL RATES</strong><br>
        <strong>Half Board</strong> (Lunch OR Dinner)<br>
        <strong>Full Board</strong> (Lunch AND Dinner)<br>
        <span style="color: red; font-weight: bold;">*Remark :</span><br><br>
        ※ <strong>Compulsory / Optional Meal Plan</strong><br>
        Adult = / Child =<br>
        <span style="color: red; font-weight: bold;">*Remark :</span><br><br>
        ※ <strong>Benefits</strong><br><br>
        ※ <strong>Transfer</strong><br><br>
        ※ <strong>Early Bird</strong><br>
        <span style="color: blue; font-weight: bold;">Validity : (Period)</span><br>
        <span style="color: red; font-weight: bold;">*Black Out : (Period)</span><br>
        • E.B 120 Days, get 25% discount.<br>
        • E.B 90 Days, get 20% discount.<br>
        • E.B 60 Days, get 10% discount.<br>
        <span style="color: red; font-weight: bold;">*Remark :</span><br><br>
        ※ <strong>Long Stay Offer / Minimum Nights Stay Offer</strong><br>
        <span style="color: blue; font-weight: bold;">Validity : (Period)</span><br>
        <span style="color: red; font-weight: bold;">*Black Out : (Period)</span><br>
        • <span style="color: green; font-weight: bold;">#MIN. 4 NIGHTS</span>, get 10% discount<br>
        • <span style="color: purple; font-weight: bold;">#MIN. 6 NIGHTS</span>, get 15% discount<br><br>
        ※ <strong>Honeymooner / Anniversary</strong>

Output: รายงานผลเป็นภาษาไทย 100% ด้วยภาษาทางการ (ห้ามใช้อิโมจิ ห้ามใช้เครื่องหมาย ### และห้ามมีคำพูดเกริ่นนำ)

[SECTION_FAIL]
**พบข้อผิดพลาด**

### หมวดหมู่: [ชื่อคอลัมน์ เช่น Net Price (Q)]
| ตำแหน่ง | ข้อมูลในไฟล์ | ข้อมูลที่ถูกต้อง | แนวทางแก้ไข |
|:---|:---|:---|:---|
| [ตำแหน่ง เช่น ทุก Room Type ใน Peak Season] | `[ค่าปัจจุบัน]` | [สรุปค่าสัญญาเป็นข้อความปกติ] | [สิ่งที่ต้องทำ] |

(หากคอลัมน์นี้เป็น AA, AD, หรือ AO คุณ **ต้อง** ใส่โค้ด HTML ตามรูปแบบนี้ด้านล่างตารางเสมอ)
<details>
<summary>คลิกเพื่อดูโค้ด HTML (Copy โค้ดด้านล่าง)</summary>

```html
[โค้ด HTML ที่ถูกต้อง]
```

</details>

---

### หมวดหมู่: [ชื่อคอลัมน์ต่อไป]
| ตำแหน่ง | ข้อมูลในไฟล์ | ข้อมูลที่ถูกต้อง | แนวทางแก้ไข |
|:---|:---|:---|:---|
| [ตำแหน่ง] | `[ค่าปัจจุบัน]` | [สรุปค่าสัญญาเป็นข้อความปกติ] | [สิ่งที่ต้องทำ] |

[SECTION_REVIEW]
**จุดที่ควรตรวจสอบเพิ่มเติม [REVIEW]**
| สถานะ | ตำแหน่ง | รายละเอียดข้อสงสัย | ข้อแนะนำ |
|:---|:---|:---|:---|
| [REVIEW] | **[Col]** | [เหตุผลที่สงสัย] | [สิ่งที่ควรเช็ค] |

---

**สรุปผลการตรวจสอบ**
- **คะแนนความถูกต้อง:** [xx]%
- **บทสรุป:** [สรุปสั้นๆ 1 ประโยค]

---

[SECTION_VERIFIED]
*** กรณีถูกต้องทั้งหมด: **STATUS: [VERIFIED] ข้อมูลถูกต้องตามสัญญาทุกประการ** ***
"""

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

    return f"""
ACT AS: Senior Hotel Data Specialist.
TASK: Extract EVERY contract type, room type, rate season, and promotion from the attached PDF into a JSON list.

{extra_instruction}
{reference_section}
━━━ PHASE 1: MANDATORY PRE-SCAN (DO THIS FIRST BEFORE GENERATING ANY ROWS) ━━━
Read the ENTIRE PDF from first page to last page. Identify and list:
  A. All CONTRACT TYPES present: Main Contract? Early Bird? Promotion? POR?
  B. All ROOM TYPES / ROOM NAMES listed in the contract tables.
  C. All SEASONS / PERIODS (e.g., Low Season, High Season, Peak Season, specific date ranges).
Do NOT generate any rows yet. Count A × B × C = total rows you MUST produce.

━━━ PHASE 2: EXTRACTION ━━━
Generate one row per unique combination of [Room Type] + [Season/Period] + [Contract Type]:
- Main Contract rows: one per room type per season period.
- Early Bird rows: one per room type per EB tier (e.g., EB 60 days, EB 90 days).
- Promotion rows: one per room type per promotion period.
- DO NOT merge multiple rooms into one row.
- DO NOT merge multiple seasons into one row.
- A single-room / single-promotion contract may legitimately have only 1–2 rows.

━━━ PHASE 3: SELF-CHECK BEFORE OUTPUTTING ━━━
Before returning JSON, verify:
  - Did you include ALL contract types found in Phase 1?
  - Did you include ALL room types in EACH contract type section?
  - Is your total row count close to A × B × C from Phase 1?
If not, go back and add the missing rows.

COLUMN RULES & FORMATTING:
1. start_date / end_date: MUST use format "YYYY-MM-DD 00:00:00.000" (always include time).
2. contract_type: Exactly one of: 'Main Contract', 'Promotion', 'Early Bird', 'POR'.
3. net_price: The contract rate (Number only). DO NOT add compulsory dinner/gala prices here.
4. HTML FORMATTING REQUIRED (STRICT TEMPLATES):

[cancellation_policy Template]
CANCELLATION : [LOW/SHOULDER/HIGH/PEAK SEASON OR PERIOD]<br>
• Cancellation made [policy details]<br>
• <span style="color: red; font-weight: bold;">NO-SHOW/Early Check Out</span> : [policy details]<br>
<span style="color: red; font-weight: bold;">*Remark</span> : [details if any]

[child_policy Template]
<span style="color: green; font-weight: bold;">Room rate includes ABF for [xx] persons</span><br>
<span style="color: green; font-weight: bold;">Maximum Occupancy : [XA / XA+XC]</span><br><br>
[policy details]<br><br>
in case have 2 Children;<br>
if have <strong>2 children ([xx-xx.99] years old)</strong> stay in room, subject to charge as policy below;<br>
1st child [policy]<br>
2nd child [policy]<br><br>
Adult [policy]<br>
<span style="color: red; font-weight: bold;">*Maximum [X] Extra bed / *CANNOT ADD EXTRA BED*</span>

[meals_and_info Template]
<span style="color: red; font-weight: bold;">#MIN. [XX] NIGHTS - COMPULSORY CHRISTMAS/NEW YEAR'S GALA DINNER on [xx - xx]</span><br>
<span style="color: red; font-weight: bold;">**NOT ALLOWED CHECK OUT on [xx]**</span> (only peak / period that have conditions)
<hr>
<strong>MAIN CONTRACT [Year] : [Period]</strong><br><br>
<span style="color: orange; font-weight: bold;">[ &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(only for Peak)</span><br>
<strong>MIN. [XX] NIGHTS during [xx-xx]</strong><br>
<strong>NOT ALLOWED CHECK OUT on [xx]</strong><br><br>
※ <strong>Compulsory</strong> New Year's Gala Dinner on 31 DEC [xx]<br>
Adult = [xxx] THB / Child ([age]) = [xxx] THB<br>
<span style="color: red; font-weight: bold;">*Remark : </span>[details]<br>
<span style="color: orange; font-weight: bold;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;]</span><br><br>
※ <strong>MEAL RATES</strong><br>
<strong>Half Board</strong> (Lunch OR Dinner) = [xxx] THB<br>
<strong>Full Board</strong> (Lunch AND Dinner) = [xxx] THB<br>
<span style="color: red; font-weight: bold;">*Remark :</span> [details]<br><br>
※ <strong>Compulsory / Optional Meal Plan</strong><br>
Adult = [xxx] THB / Child = [xxx] THB<br>
<span style="color: red; font-weight: bold;">*Remark :</span> [details]<br><br>
※ <strong>Benefits</strong><br>[details]<br><br>
※ <strong>Transfer</strong><br>[details]<br><br>
※ <strong>Early Bird</strong><br>
<span style="color: blue; font-weight: bold;">Validity : [Period]</span><br>
<span style="color: red; font-weight: bold;">*Black Out : [Period]</span><br>
• E.B [xx] Days, get [xx]% discount.<br>
<span style="color: red; font-weight: bold;">*Remark :</span> [details]<br><br>
※ <strong>Long Stay Offer / Minimum Nights Stay Offer</strong><br>
<span style="color: blue; font-weight: bold;">Validity : [Period]</span><br>
<span style="color: red; font-weight: bold;">*Black Out : [Period]</span><br>
• <span style="color: green; font-weight: bold;">#MIN. [X] NIGHTS</span>, get [xx]% discount<br><br>
※ <strong>Honeymooner / Anniversary</strong><br>[details]

5. Numeric columns (net_price, child_share_bed_abf, child_extra_bed_abf, extra_bed_abf): Numbers only (e.g. 1400.0).
6. PERIOD SPLITTING: Split into separate rows if conditions change within the season.
7. MISSING DATA: For any key not found, output "". DO NOT output 0 or FALSE.

OUTPUT FORMAT: A pure JSON LIST of objects with REQUIRED KEYS: {json.dumps(EXCEL_UPLOAD_COLUMNS[6:])}
CRITICAL: Return ONLY valid JSON — no markdown, no backticks, no explanation text.
"""


def extract_pdf_to_excel_json(pdf_bytes: bytes, api_key: str, excel_bytes: bytes = None):
    """
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
        except Exception:
            pass  # If Excel can't be read, proceed without reference

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
    except Exception:
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
                        f"⚠️ PREVIOUS ATTEMPT WARNING: You only generated {len(data)} rows, "
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