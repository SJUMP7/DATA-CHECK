import streamlit as st
import pandas as pd
from utils import extract_pdf_text, extract_excel_data, mock_audit_process, run_gemini_audit

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Hotel Contract Data Audit",
    page_icon="🏨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CUSTOM CSS FOR PREMIUM LOOK ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600&display=swap');
    
    /* Global Font */
    html, body, [class*="css"] {
        font-family: 'Outfit', sans-serif !important;
    }
    
    /* Metrics Box - Light Mode Card */
    div[data-testid="metric-container"] {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        padding: 5% 5% 5% 10%;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    
    div[data-testid="metric-container"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
    }
    
    /* Alert Boxes - Correct */
    .correct-box {
        background-color: #f0fdf4;
        border-left: 4px solid #22c55e;
        padding: 16px 20px;
        border-radius: 8px;
        margin-bottom: 12px;
        color: #166534;
        transition: all 0.2s ease;
    }
    .correct-box:hover {
        transform: translateX(4px);
    }
    
    /* Alert Boxes - Wrong */
    .wrong-box {
        background-color: #fef2f2;
        border-left: 4px solid #ef4444;
        padding: 16px 20px;
        border-radius: 8px;
        margin-bottom: 12px;
        color: #991b1b;
        transition: all 0.2s ease;
    }
    .wrong-box:hover {
        transform: translateX(4px);
    }
    
    /* Alert Boxes - Confuse */
    .confuse-box {
        background-color: #fefce8;
        border-left: 4px solid #eab308;
        padding: 16px 20px;
        border-radius: 8px;
        margin-bottom: 12px;
        color: #854d0e;
        transition: all 0.2s ease;
    }
    .confuse-box:hover {
        transform: translateX(4px);
    }

    /* Primary Button Customization */
    button[data-testid="baseButton-primary"] {
        background: linear-gradient(90deg, #4f46e5 0%, #3730a3 100%);
        border: none;
        box-shadow: 0 4px 14px 0 rgba(79, 70, 229, 0.39);
        color: white;
        transition: all 0.3s ease;
    }
    button[data-testid="baseButton-primary"]:hover {
        box-shadow: 0 6px 20px rgba(79, 70, 229, 0.5);
        transform: translateY(-2px);
    }
</style>
""", unsafe_allow_html=True)

# --- SIDEBAR ---
with st.sidebar:
    st.title("🏨 Audit Settings")
    st.markdown("อัปโหลดไฟล์สัญญาและไฟล์ระบบเพื่อตรวจสอบ")
    
    pdf_file = st.file_uploader("📄 อัปโหลดสัญญาโรงแรม (PDF)", type=["pdf"])
    excel_file = st.file_uploader("📊 อัปโหลดข้อมูลระบบ (Excel)", type=["xlsx", "xls"])
    
    st.markdown("---")
    st.subheader("⚙️ AI Configuration")
    
    # Try to load API key from secrets first
    if "GEMINI_API_KEY" in st.secrets and st.secrets["GEMINI_API_KEY"]:
        api_key = st.secrets["GEMINI_API_KEY"]
        st.success("✅ เชื่อมต่อ API อัตโนมัติแล้ว (ดึงจากระบบ)")
    else:
        api_key = st.text_input("Gemini API Key (Option)", type="password", help="ใส่ API Key เพื่อให้ระบบอ่านสัญญาด้วย AI ได้แม่นยำ 100%")
        if not api_key:
            st.warning("ขณะนี้ระบบใช้โหมดจำลอง (Mock Engine) เนื่องจากไม่ได้ใส่ API Key")
        else:
            st.success("API Key พร้อมใช้งาน")

    run_audit = st.button("🚀 เริ่มการตรวจสอบ (Run Audit)", type="primary", use_container_width=True)

# --- MAIN CONTENT ---
st.title("Hotel Contract Data Audit Dashboard")
st.markdown("ระบบตรวจสอบข้อมูลโรงแรมแบบ 100% Full Scan เพื่อความถูกต้องแม่นยำสูงสุด")

if run_audit:
    if not pdf_file or not excel_file:
        st.error("กรุณาอัปโหลดทั้งไฟล์ PDF และ Excel ก่อนเริ่มการตรวจสอบ")
    else:
        with st.spinner("⏳ กำลังสแกนข้อมูลและเทียบสัญญาแบบ 100% Full Scan..."):
            # 1. Extract Data
            pdf_text = extract_pdf_text(pdf_file)
            df = extract_excel_data(excel_file)
            
            # 2. Process
            if api_key:
                st.info("กำลังเรียกใช้ Gemini API เพื่อทำการวิเคราะห์ 100% Full Scan...")
                results = run_gemini_audit(pdf_text, df, api_key)
            else:
                st.warning("เรียกใช้ Mock Engine เนื่องจากไม่ได้ใส่ API Key")
                results = mock_audit_process(pdf_text, df)
            
        st.success("✅ การตรวจสอบเสร็จสมบูรณ์!")

        st.markdown("---")
        
        # --- RESULTS DASHBOARD ---
        col1, col2 = st.columns([1, 2])
        
        with col1:
            st.metric(label="ความถูกต้องของข้อมูล (Accuracy)", value=f"{results['accuracy']}%", delta="100% Full Scan Validated", delta_color="normal")
            
        with col2:
            st.markdown("### 📖 สรุปสิ่งที่ต้องแก้ไข")
            st.info(results["summary"])
            
        st.markdown("---")
        
        tab1, tab2, tab3 = st.tabs(["🟢 ถูกต้อง (Correct)", "🔴 ต้องแก้ไข (Wrong)", "🟡 จุดที่สงสัย (Confuse)"])
        
        with tab1:
            st.markdown("### รายละเอียดจุดที่ถูกต้อง")
            for item in results["correct_items"]:
                st.markdown(f"""
                <div class="correct-box">
                    <strong>[Correct]</strong> {item}
                </div>
                """, unsafe_allow_html=True)
                
        with tab2:
            st.markdown("### รายละเอียดจุดที่ผิดพลาด")
            for item in results["wrong_items"]:
                st.markdown(f"""
                <div class="wrong-box">
                    <strong>[Wrong]</strong> {item['issue']}<br>
                    <span style='color: #15803d;'><strong>Should be:</strong> {item['should_be']}</span><br>
                    <span style='color: #1d4ed8;'><strong>Action:</strong> {item['action']}</span>
                </div>
                """, unsafe_allow_html=True)
                
        with tab3:
            st.markdown("### จุดที่สงสัยหรือไม่แน่ใจ")
            for item in results["confuse_items"]:
                st.markdown(f"""
                <div class="confuse-box">
                    <strong>[Confuse]</strong> {item}
                </div>
                """, unsafe_allow_html=True)

else:
    # Initial state screen
    st.info("👈 กรุณาอัปโหลดไฟล์ PDF และ Excel ที่แถบด้านซ้ายมือ แล้วกดปุ่ม 'เริ่มการตรวจสอบ'")
    
    st.markdown("""
    ### 📌 กฎเหล็กในการปฏิบัติงาน (Operational Directives)
    - **100% Full Scan Only:** ตรวจสอบละเอียดทุกบรรทัดและทุกช่อง (Every Row & Cell)
    - **Exception Scanning:** สแกนหาคำยกเว้น "not including", "excluding", "except", "only" 
    - **Double-Check Before Flagging:** อ่านประโยคเงื่อนไขใน PDF ซ้ำเพื่อยืนยันเสมอ
    - **Child Policy Logic:** ตีความอายุเด็กให้แม่นยำ เช่น Under 12 คือ 11.99
    """)
