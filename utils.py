import pandas as pd
import pdfplumber
import time
import google.generativeai as genai
import json

def extract_pdf_text(pdf_file):
    """
    Extracts text from an uploaded PDF file.
    """
    text = ""
    try:
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
        return text
    except Exception as e:
        return f"Error reading PDF: {str(e)}"

def extract_excel_data(excel_file):
    """
    Reads the Excel file and returns a pandas DataFrame.
    """
    try:
        # Read the excel file
        df = pd.read_excel(excel_file)
        return df
    except Exception as e:
        return None

def mock_audit_process(pdf_text, df):
    """
    A mock rule engine that simulates analyzing the data.
    In the future, this function will call the LLM API.
    """
    # Simulate processing time
    time.sleep(2)
    
    total_rows = len(df) if df is not None else 0
    if total_rows == 0:
        return {
            "accuracy": 0,
            "summary": "ไม่พบข้อมูลในไฟล์ Excel ที่อัปโหลด",
            "correct_items": [],
            "wrong_items": [],
            "confuse_items": []
        }
        
    # Mock Logic: We'll pretend we checked the data and found some correct and some wrong.
    # In reality, this requires LLM.
    
    accuracy = 85
    summary = "พบจุดที่ผิดพลาดเล็กน้อยในส่วนของราคา Early Bird และการตีความ Child Policy กรุณาตรวจสอบรายละเอียดด้านล่าง"
    
    correct_items = [
        "ตรวจสอบ Mapping วันที่ (Col G, H) ถูกต้องทั้งหมด",
        "ชื่อประเภทห้องพัก (Col AL) ตรงกับในสัญญา",
        "ตรวจสอบ Cancellation Policy (Col AA) ส่วนใหญ่ตรงตามเงื่อนไขใน PDF"
    ]
    
    wrong_items = [
        {
            "issue": "ราคา Net Price (Col Q) ของห้อง Deluxe ในช่วง High Season คำนวณส่วนลด Early Bird ผิด",
            "should_be": "ควรเป็น 2,500 THB (ลด 10% จากราคาหลัก 2,777 THB)",
            "action": "แก้ไขตัวเลขใน Excel เป็น 2500"
        },
        {
            "issue": "ข้อมูล Child Policy (Col AD) ขาดเงื่อนไข 'ไม่รวมเตียงเสริม'",
            "should_be": "Child 5-11.99 years old Sharing Bed + ABF = 500 THB (No extra bed)",
            "action": "เพิ่มข้อความใน HTML pattern ตาม PDF"
        }
    ]
    
    confuse_items = [
        "ส่วนลด Long Stay Offer ระบุว่า 'As per main contract' แต่ในระบบไม่มีการเชื่อมโยงกับสัญญาหลักชัดเจน"
    ]
    
    return {
        "accuracy": accuracy,
        "summary": summary,
        "correct_items": correct_items,
        "wrong_items": wrong_items,
        "confuse_items": confuse_items
    }

def run_gemini_audit(pdf_text, df, api_key):
    """
    Runs the actual 100% Full Scan using Gemini API.
    """
    genai.configure(api_key=api_key)
    
    # ระบบเลือกโมเดลอัตโนมัติ เพื่อป้องกันปัญหา 404 Model Not Found
    try:
        available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        if 'models/gemini-1.5-pro' in available_models:
            model_name = 'gemini-1.5-pro'
        elif 'models/gemini-1.5-pro-latest' in available_models:
            model_name = 'gemini-1.5-pro-latest'
        elif 'models/gemini-1.5-flash' in available_models:
            model_name = 'gemini-1.5-flash'
        elif 'models/gemini-pro' in available_models:
            model_name = 'gemini-pro'
        else:
            # ใช้ตัวแรกที่รองรับ
            model_name = available_models[0].replace('models/', '') if available_models else 'gemini-pro'
    except Exception:
        # ถ้าเช็คไม่ได้ให้ตกมาใช้ gemini-1.5-flash ก่อน
        model_name = 'gemini-1.5-flash'

    model = genai.GenerativeModel(model_name)
    
    # Convert DataFrame to a string format that LLM can easily read (CSV format is usually best for tabular data)
    csv_data = df.to_csv(index=False)
    
    prompt = f"""
    Data Recheck : (100% Full Scan)

    Tone & Persona :
    ผู้เชี่ยวชาญที่มีความละเอียดสูงด้านการตรวจสอบความถูกต้อง โดยนำสัญญาโรงแรม (PDF) มาตรวจสอบว่า ไฟล์ Excel ทำถูกต้องตามสัญญาโรงแรมหรือไม่ - โดยยึดหลักการตรวจสอบแบบ 100% Full Scan ไม่สุ่มตรวจ มีความละเอียดรอบคอบไม่ประมาท ไม่ละเลยข้อมูลในช่องเล็กๆ และพร้อมรายงานจุดผิดสังเกตทันทีที่ตรวจพบ 

    How to check : 
    1. อ่านไฟล์ PDF ที่ส่งให้ (100% Full Scan) เพื่อเตรียมไว้ตรวจสอบไฟล์ Excel ที่ถูกทำไว้แล้ว
    2. นำข้อมูลในไฟล์ PDF มาตรวจสอบเปรียบเทียบกับไฟล์ Excel เพื่อตรวจสอบความถูกต้อง
    3. หาก Full Scan และตรวจสอบเสร็จแล้ว ให้รายงานผล ตามรูปแบบที่ตั้งไว้ให้ผู้ใช้งาน อ่านง่าย

    Operational Directives (กฎเหล็กในการปฏิบัติงาน):
    1. 100% Full Scan Only: ห้ามใช้วิธีสุ่มตรวจ (Sampling) หรือดูแค่ส่วนหัว/ท้ายไฟล์เด็ดขาด ต้องตรวจสอบละเอียด "ทุกบรรทัดและทุกช่อง" (Every Row & Cell) 
    2. Exact Row Mapping (ห้ามสลับบรรทัด): การตรวจสอบราคา (Col Q, AE, AF, AG) ต้อง Map ข้อมูลให้ตรงกับ "ชื่อประเภทห้องพัก" (Room Type - Col AL) ทุกครั้ง ห้ามอ่านข้อมูลเหลื่อมบรรทัดหรือสลับประเภทห้องเด็ดขาด
    3. Exception Scanning (สแกนหาคำยกเว้นอย่างเข้มงวด): ทุกครั้งที่มีเงื่อนไขโปรโมชั่น, Early Bird, หรือ Long Stay ให้สแกนหาคำว่า "not including", "excluding", "except", "only" ในสัญญา PDF ทันที หากพบข้อยกเว้น (เช่น ไม่รวมเตียงเสริม, ไม่รวมอาหารเช้า) ห้ามเหมารวมว่าส่วนลดนั้นครอบคลุมทั้งหมด
    4. Double-Check Before Flagging: ก่อนที่จะตัดสินว่าไฟล์ Excel ใส่ข้อมูลผิด (🔴) หรือคำนวณถูก (🟢) โดยเฉพาะเรื่อง 'ส่วนลด' ให้กลับไปอ่านประโยคเงื่อนไขใน PDF ซ้ำเป็นรอบที่ 2 เสมอเพื่อยืนยันความถูกต้อง
    5. Child Policy Logic : หากสัญญาระบุ Under 12 ให้ใช้เป็น 11.99, ระบุ 5 - 11 years old ควรเป็น 5 - 11.99 years old 
    6. Discount Logic : 
    6.1 หาก Column K เป็น Early Bird หรือ Promotion ให้ตรวจสอบการลดราคา Column Q ว่าลดถูกต้องจากราคาหลักไหม 
    6.2 ***ระวังเงื่อนไขการหักส่วนลดกับ Extra Bed / Child Policy หากสัญญาไม่ระบุให้ลด ห้ามนำส่วนลดไปคำนวณเด็ดขาด***
    7. Typo Detection: ตรวจสอบตัวสะกดอย่างเข้มงวด รวมถึงการใส่วันที่ที่เป็นไปไม่ได้ (เช่น 31 เมษายน)
    8. Cancellation : สกัดตัวเลขจำนวนวันยกเลิกออกจากข้อความ HTML (Col AA) และตรวจสอบว่าจำนวนวันที่ระบุ ตรงตามช่วงเวลาฤดูกาลในสัญญา PDF หรือไม่
    9. MIN. NIGHTS STAY : Column O หากสัญญาระบุว่า Period ไหนมี MIN NIGHTS ให้ตรวจสอบอย่างละเอียด
    10. ห้ามใช้ความรู้พื้นฐาน หรือมาตรฐานทั่วไปของอุตสาหกรรมโรงแรม (Industry Standard) มาตัดสิน ทุกการตรวจสอบต้องอ้างอิงจาก "ตัวอักษรและตัวเลขที่ระบุในไฟล์ PDF สัญญาเท่านั้น" 
    11. ในการตรวจสอบ HTML (Col AA, AD, AO) ให้โฟกัสที่ 'ความถูกต้องของเนื้อหา ตัวเลข และเงื่อนไข' เป็นหลัก 
    12. หากข้อมูลมีหลายบรรทัด แบ่ง Batch ได้ แต่ต้องแจ้งผลลัพธ์ครอบคลุมทั้งหมด

    Process:
    1. นำข้อมูลไฟล์ PDF มาตรวจสอบในไฟล์ Excel
    2. Column Mapping : ข้อมูลพื้นฐาน Excel
    2.1 Col G (start_date): วันที่เริ่มต้น Period (Format: วัน/เดือน/ปี )
    2.2 Col H (end_date): วันที่สิ้นสุด Period (Format: วัน/เดือน/ปี )
    2.3 Col I (refundable): FALSE เท่านั้น
    2.4 Col J (abf): ไม่ต้องตรวจสอบ
    2.5 Col K (contract_type): ระบุค่าเป็นค่าใดค่าหนึ่งต่อไปนี้เท่านั้น: Main Contract, Early Bird, Promotion, POR
    2.6 Col L (cutoff_date) : ตัวเลข ตามสัญญา
    2.7 Col M (hotel_supplier) : ชื่อโรงแรม 
    2.8 Col O (min_nights_stay): ตัวเลขจำนวนคืน ตามสัญญา (หากไม่มีเงื่อนไข เว้นว่าง)
    2.9 Col P (min_advance_days): ตัวเลข ตรวจสอบเฉพาะ Early Bird
    2.10 Col Q (net_price): ตัวเลขตามราคาสัญญา
    2.11 Col U (promo_book_till) : วันที่สิ้นสุด Booking (Format: ปี/เดือน/วัน 23:59:59 ) 
    2.12 Col V (promo_code) : Code Promotion / ถ้าเป็น Early Bird และ long stay offer ที่มีลดราคาจะใช้แพทเทิล E.B xx DAYS, LONG STAY OFFER
    2.13 Col W (promo_note): ใส่เฉพาะ Condition สำคัญโดยเฉพาะ MIN. xx NIGHTS, COMPULSORY NEW YEAR GALA DINNER on xx, NOT ALLOWED CHECK OUT 
    2.14 Col X (room_allotment): ระบุค่าดังนี้: Free Sales, On Request, หรือ ตัวเลขจำนวนห้อง (หากไม่มีเว้นว่าง)
    2.15 Col AA (cancellation_policy): จัดรูปแบบข้อความ HTML ตาม Pattern นี้ (แพทเทิลไม่ตายตัวเน้นให้ข้อมูลถูกตามไฟล์สัญญา PDF ) :
    CANCELLATION : LOW/SHOULDER/HIGH/PEAK SEASON หรือ ระบุเป็นวันที่ [ แพทเทิลวันที่ คือ DD MM YY : 1 MAR 27 ]
    Cancellation made (รายละเอียด policy)
    NO-SHOW/Early Check Out : (รายละเอียด policy)
    *Remark : (ถ้ามี)
    2.16 Col AD (child_policy): จัดรูปแบบข้อความ HTML ตาม Pattern นี้ (แพทเทิลไม่ตายตัวเน้นให้ข้อมูลถูกตามไฟล์สัญญา PDF ) 
    Room rate includes ABF for xx persons (ไม่จำเป็นต้องใส่ก็ได้)
    Maximum Occupancy : XA / XA+XC (มีหรือไม่มีก็ได้หากไฟล์สัญญา PDF ระบุไว้ แต่ Excel ไม่ได้ระบุให้แนะนำให้ผู้ใช้กรอก)
    Child ?? years old Sharing Bed + ABF = ?? THB
    Child ?? years old Extra Bed + ABF = ?? THB
    Adult Extra Bed + ABF = ?? THB
    in case have 2 Children; (หากห้องให้มีเด็ก 2 คนขึ้นไป แนะนำแพทเทิลนี้ )
    if have 2 children (xx-xx.99 years old) stay in room, subject to charge as policy below;
    1st child Sharing Bed + ABF = ?? THB
    2nd child Extra Bed + ABF = ?? THB
    Adult Extra Bed + ABF = ?? THB
    *Remark : Maximum X Extra bed / *CANNOT ADD EXTRA BED*
    2.17 Col AE (child_share_bed_abf) ระบุเป็นตัวเลขโดยอิงราคา Child Sharing Bed+ABF ตามในสัญญา
    2.18 Col AF (child_extra_bed_abf) ระบุเป็นตัวเลขโดยอิงราคา Child Extra Bed +ABF ตามในสัญญา 
    2.19 Col AG (extra_bed_abf) ระบุเป็นตัวเลขโดยอิงราคา Adult Extra Bed +ABF ตามในสัญญา 
    2.20 Col AI (full_board): ตัวเลข ตามที่ระบุไว้ในสัญญา
    2.21 Col AJ (half_board): ตัวเลข ตามที่ระบุไว้ในสัญญา
    2.22 Col AL (room_name): ระบุชื่อห้องพักตามสัญญา “ยืดยุ่นได้” เช่น สัญญาระบุ Deluxe room ไฟล์ Excel สามารถระบุ Deluxe ได้
    2.23 Col AO (meals_and_info): รวมข้อมูล Condition, Meal, Transfer, E.B, Long Stay ให้จัดรูปแบบ HTML ตาม Pattern นี้ (หัวข้อไหนไม่มีข้อมูล ให้ตัดออก): (แพทเทิลไม่ตายตัวเน้นให้ข้อมูลถูกตามไฟล์สัญญา PDF ):
    [ #MIN. XX NIGHTS - COMPULSORY CHRISTMAS/NEW YEAR’S GALA DINNER on xx - xx
    **NOT ALLOWED CHECK OUT on xx** บรรทัดนี้ใส่เฉพาะช่วงที่ Period ระบุไว้ 
    ※ MEAL RATES
    ※ Benefits 
    ※ Transfer (ระบุรายละเอียดตามสัญญาเฉพาะโรงแรมเกาะ/ทะเล)
    ※ Early Bird
    ※ Long Stay Offer / Minimum Nights Stay Offer
    ※ Honeymooner / Anniversary ]
    
    ข้อมูลสัญญา PDF:
    ```
    {pdf_text}
    ```
    
    ข้อมูล Excel ระบบ:
    ```csv
    {csv_data}
    ```
    
    **คำสั่งสำหรับการตอบกลับ (Output Requirement):**
    ลดคำชมแสดงคำตอบล้วนๆ คุณต้องตอบกลับมาเป็น **JSON format เท่านั้น** ห้ามมีข้อความอื่นนอกเหนือจาก JSON block เริ่มด้วย {{ และจบด้วย }}
    โครงสร้าง JSON ที่ต้องการคือ:
    {{
        "accuracy": <ตัวเลข 0-100 ระบุเปอร์เซ็นต์ความถูกต้องโดยรวม>,
        "summary": "<ข้อความสรุปสิ่งที่ต้องแก้ไขสั้นๆ ครบทุกข้อ>",
        "correct_items": [
            "<รายการสิ่งที่ถูกต้อง 1>",
            "<รายการสิ่งที่ถูกต้อง 2>"
        ],
        "wrong_items": [
            {{
                "issue": "<คำอธิบายสิ่งที่ผิดพลาด *หมายเหตุ : หากข้อมูลเป็นโค้ด HTML บน Excel ให้อ่านเป็นข้อมูลแล้วหากข้อมูลตรงไหนผิด ให้แสดงคำตอบมาเป็นตัวอักษรแทน>",
                "should_be": "<สิ่งที่ควรจะเป็นตาม Logic/สัญญา>",
                "action": "<คำแนะนำสั้นๆ ในการแก้ไข>"
            }}
        ],
        "confuse_items": [
            "<จุดที่ข้อมูลขัดแย้งกัน หรือไม่แน่ใจ หรือไม่มีระบุในสัญญา เช่น As per main contract แต่หาไม่เจอ>"
        ]
    }}
    """

    
    try:
        response = model.generate_content(prompt)
        response_text = response.text.strip()
        
        # Clean up if the model wrapped it in markdown code blocks
        if response_text.startswith("```json"):
            response_text = response_text[7:]
        if response_text.startswith("```"):
            response_text = response_text[3:]
        if response_text.endswith("```"):
            response_text = response_text[:-3]
            
        result_json = json.loads(response_text.strip())
        return result_json
        
    except json.JSONDecodeError:
        return {
            "accuracy": 0,
            "summary": "เกิดข้อผิดพลาดในการอ่านผลลัพธ์จาก AI (JSON Parse Error)",
            "correct_items": [],
            "wrong_items": [{"issue": "AI ไม่ได้ส่งข้อมูลกลับมาเป็น JSON ตามที่กำหนด", "should_be": "JSON Format", "action": "ลองกดตรวจสอบใหม่อีกครั้ง"}],
            "confuse_items": [f"Raw response: {response.text}"]
        }
    except Exception as e:
        return {
            "accuracy": 0,
            "summary": f"เกิดข้อผิดพลาดในการเชื่อมต่อ API: {str(e)}",
            "correct_items": [],
            "wrong_items": [],
            "confuse_items": []
        }
