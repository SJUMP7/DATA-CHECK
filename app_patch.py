# ================================================================
#  APP.PY — CHANGES REQUIRED TO MIGRATE FROM GEMINI → CLAUDE
# ================================================================
#
# ไฟล์นี้อธิบาย diff ที่ต้องแก้ใน app.py ของคุณ
# ไม่ต้องแตะ logic อื่น — เปลี่ยนแค่ 4 จุดด้านล่าง
#
# ================================================================


# ── CHANGE 1: Import  (บรรทัดบนสุดของ app.py) ─────────────────
# BEFORE:
#   from utils import stream_recheck_analysis, validate_api_key
#
# AFTER:
from utils import (
    stream_recheck_analysis,
    validate_api_key,
    generate_excel_from_pdf,   # ← เพิ่ม Feature 2
)


# ── CHANGE 2: Sidebar — API Key label ──────────────────────────
# ค้นหา "Gemini API Key" ใน app.py แล้วเปลี่ยนเป็น "Claude API Key"
#
# BEFORE:
#   api_key = st.text_input("Gemini API Key", ...)
#   st.caption("Get your key at makersuite.google.com")
#
# AFTER:
#   api_key = st.text_input(
#       "Claude API Key",
#       type="password",
#       placeholder="sk-ant-api03-...",
#       help="Get your key at console.anthropic.com",
#   )
#   st.caption("Get your key at console.anthropic.com/settings/api-keys")


# ── CHANGE 3: API Key validation call ──────────────────────────
# validate_api_key() signature ไม่เปลี่ยน — ใช้ได้เลย
# แต่ return format เปลี่ยน:
#
# BEFORE (Gemini):  validate_api_key(key) → True/False
# AFTER (Claude):   validate_api_key(key) → (bool, message_string)
#
# ถ้าโค้ดเดิมเป็น:
#   if validate_api_key(api_key):
#       st.success("Valid")
#
# ให้เปลี่ยนเป็น:
#   is_valid, msg = validate_api_key(api_key)
#   if is_valid:
#       st.success(msg)
#   else:
#       st.error(msg)


# ── CHANGE 4: Feature 2 — Generate Excel button ────────────────
# ใส่ส่วนนี้ต่อจาก "Start Audit" button ใน app.py
# (ในบล็อกที่เช็ค has_pdf)

def render_generate_excel_section(pdf_file, api_key, hotel_id, room_id_map):
    """
    วางฟังก์ชันนี้ใน app.py และเรียกใช้หลัง Start Audit button
    """
    import streamlit as st

    st.markdown("---")
    st.markdown(
        '<div style="font-size:10px;font-weight:800;letter-spacing:.15em;'
        'text-transform:uppercase;color:#94a3b8;margin-bottom:12px">'
        'FEATURE 2 — GENERATE EXCEL FROM PDF</div>',
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns([2, 1])
    with col1:
        hotel_id_input = st.text_input(
            "Hotel ID (จาก dashboard)",
            value=hotel_id or "",
            placeholder="เช่น 1288055",
        )
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        generate_btn = st.button(
            "📥 Generate Upload Excel from PDF",
            use_container_width=True,
            disabled=not (pdf_file and api_key),
        )

    if generate_btn:
        with st.spinner("AI กำลังสกัดข้อมูลจาก PDF..."):
            pdf_bytes = pdf_file.getvalue() if hasattr(pdf_file, "getvalue") else pdf_file
            xlsx_bytes, error = generate_excel_from_pdf(
                pdf_bytes=pdf_bytes,
                api_key=api_key,
                hotel_id=hotel_id_input,
                room_id_map=room_id_map or {},
            )

        if error:
            st.error(f"เกิดข้อผิดพลาด: {error}")
        else:
            st.success(f"สร้างไฟล์สำเร็จ — {len(xlsx_bytes):,} bytes")
            st.warning(
                "⚠️ คอลัมน์ A–F (_id, hotel_id, room_id, status, dates) "
                "ต้องกรอกเพิ่มเองหรือ match กับ dashboard ก่อน import"
            )
            st.download_button(
                label="📥 Download Excel",
                data=xlsx_bytes,
                file_name=f"upload_{hotel_id_input or 'hotel'}_{__import__('datetime').date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
