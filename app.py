import os
import re
from datetime import datetime
import streamlit as st
from utils import stream_recheck_analysis, validate_api_key

# ─── Dynamic Honest Milestone Parser ─────────────────────────────────────────
def get_honest_milestone(full_text, chunk_count):
    if not full_text:
        return 5, "Initializing AI Engine..."
    if "[SECTION_FAIL]" not in full_text:
        # Phase 1: Analyzing PDF structure
        perc = min(10 + (chunk_count * 1.5), 30)
        return perc, "Reading Contract PDF..."
    if "[SECTION_REVIEW]" not in full_text:
        # Phase 2: Cross-referencing pricing rows
        perc = min(30 + (chunk_count * 0.8), 65)
        return perc, "Cross-Referencing Excel Rates..."
    if "คะแนนความถูกต้อง" not in full_text:
        # Phase 3: Validating cancellation and child policies
        perc = min(65 + (chunk_count * 0.5), 90)
        return perc, "Validating Booking Policies..."
    # Phase 4: Finalizing report assembly
    perc = min(90 + (chunk_count * 0.3), 98)
    return perc, "Assembling Final Audit Score..."

# ─── Interactive Cancellation Interrupt (C2 & H2) ───────────────────────────
if "cancel_requested" not in st.session_state:
    st.session_state.cancel_requested = False

if st.session_state.get("cancel_btn_trigger") or st.session_state.get("cancel_requested"):
    st.session_state.is_auditing = False
    st.session_state.cancel_requested = False
    st.session_state.pop("audit_done", None)
    st.session_state.pop("cancel_btn_trigger", None)
    st.rerun()

# ─── Destructive Start Fresh Confirmation Modal (C1) ─────────────────────────
if st.session_state.get("confirm_reset", False):
    st.markdown("""
        <div class="fixed-overlay"></div>
        <div class="fixed-modal" style="border-color: rgba(239, 68, 68, 0.4) !important; background: rgba(11, 15, 25, 0.95) !important;">
            <h3 style="color: #ef4444 !important;">ยืนยันการล้างข้อมูล</h3>
            <p style="margin-bottom: 24px; color: #94a3b8 !important;">คุณแน่ใจหรือไม่ว่าต้องการล้างข้อมูลและไฟล์ที่อัปโหลดทั้งหมด? การดำเนินการนี้ไม่สามารถย้อนกลับได้</p>
        </div>
    """, unsafe_allow_html=True)
    
    # Render two action buttons inline
    _, btn1, btn2, _ = st.columns([1, 2, 2, 1])
    with btn1:
        if st.button("ยืนยันการล้างข้อมูล", type="primary", use_container_width=True, key="confirm_reset_yes"):
            # Clear all states
            for key in ["audit_done", "is_auditing", "_audit_result", "prev_focus", "show_upload", "pdf", "excel", 
                        "cached_pdf_bytes", "cached_pdf_name", "cached_excel_bytes", "cached_excel_name", "confirm_reset"]:
                st.session_state.pop(key, None)
            st.rerun()
    with btn2:
        if st.button("ยกเลิก", use_container_width=True, key="confirm_reset_no"):
            st.session_state.confirm_reset = False
            st.rerun()
    st.stop()

# ─── Session State Key Constants (prevents silent typo bugs) ─────────────────
_KEY_AUDIT_DONE   = "audit_done"
_KEY_IS_AUDITING  = "is_auditing"
_KEY_SHOW_UPLOAD  = "show_upload"
_KEY_AUDIT_RESULT = "_audit_result"
_KEY_FOCUS_LIST   = "focus_list"
_KEY_PREV_FOCUS   = "prev_focus"


st.set_page_config(
    page_title="Recheck Excel Data",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── CSS INJECTION ───────────────────────────────────────────────────────────
# ─── CSS INJECTION ───────────────────────────────────────────────────────────
def load_css():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Outfit:wght@500;600;700;800&display=swap');
    
    html { scroll-behavior: smooth !important; }
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif !important;
        color: var(--text-color) !important;
    }

    /* Base App Background & Glow Effect */
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {
        background-color: var(--background-color) !important;
        background-image: 
            radial-gradient(at 0% 0%, color-mix(in srgb, #10b981 4%, transparent) 0px, transparent 50%),
            radial-gradient(at 100% 100%, color-mix(in srgb, #3b82f6 4%, transparent) 0px, transparent 50%) !important;
        color: var(--text-color) !important;
    }
    .block-container { padding: 2.5rem 3rem 4rem !important; max-width: 1400px !important; }

    /* Sidebar Custom Styling */
    section[data-testid="stSidebar"] {
        background-color: var(--secondary-background-color) !important;
        border-right: 1px solid color-mix(in srgb, var(--text-color) 8%, transparent) !important;
    }
    section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] {
        padding: 2.5rem 1.5rem !important;
    }

    /* Sidebar premium navigation radio styles (Emoji-free) */
    div[data-testid="stSidebar"] div[role="radiogroup"] {
        display: flex !important;
        flex-direction: column !important;
        gap: 10px !important;
        background-color: transparent !important;
        padding: 0 !important;
    }
    div[data-testid="stSidebar"] div[role="radiogroup"] label {
        display: flex !important;
        align-items: center !important;
        background-color: color-mix(in srgb, var(--text-color) 2%, transparent) !important;
        border: 1px solid color-mix(in srgb, var(--text-color) 6%, transparent) !important;
        border-radius: 8px !important;
        padding: 14px 18px !important;
        color: color-mix(in srgb, var(--text-color) 60%, transparent) !important;
        font-size: 11px !important;
        font-weight: 700 !important;
        letter-spacing: 0.08em !important;
        text-transform: uppercase !important;
        transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1) !important;
        cursor: pointer !important;
        margin: 0 0 4px 0 !important;
    }
    div[data-testid="stSidebar"] div[role="radiogroup"] label:hover {
        background-color: color-mix(in srgb, var(--text-color) 4%, transparent) !important;
        border-color: color-mix(in srgb, var(--text-color) 12%, transparent) !important;
        color: var(--text-color) !important;
        transform: translateY(-1px) !important;
    }
    div[data-testid="stSidebar"] div[role="radiogroup"] label[data-checked="true"] {
        background: linear-gradient(135deg, color-mix(in srgb, #10b981 8%, transparent), color-mix(in srgb, #3b82f6 8%, transparent)) !important;
        border: 1px solid color-mix(in srgb, #3b82f6 40%, transparent) !important;
        color: #3b82f6 !important;
        box-shadow: 0 4px 20px color-mix(in srgb, #3b82f6 12%, transparent) !important;
    }
    div[data-testid="stSidebar"] div[role="radiogroup"] [data-testid="stRadioButtonCircle"] {
        display: none !important;
    }
    div[data-testid="stSidebar"] div[role="radiogroup"] label div[data-testid="stMarkdownContainer"] {
        padding-left: 0 !important;
        width: 100% !important;
    }

    /* Sidebar clean buttons */
    section[data-testid="stSidebar"] .stDownloadButton button {
        background-color: color-mix(in srgb, var(--text-color) 2%, transparent) !important;
        border: 1px solid color-mix(in srgb, var(--text-color) 6%, transparent) !important;
        color: var(--text-color) !important;
        border-radius: 8px !important;
        font-size: 11px !important;
        font-weight: 700 !important;
        letter-spacing: 0.06em !important;
        text-transform: uppercase !important;
        justify-content: center !important;
        padding: 10px 16px !important;
        width: 100% !important;
        transition: all 0.2s ease !important;
    }
    section[data-testid="stSidebar"] .stDownloadButton button:hover {
        background-color: color-mix(in srgb, var(--text-color) 6%, transparent) !important;
        border-color: color-mix(in srgb, var(--text-color) 15%, transparent) !important;
        color: var(--text-color) !important;
    }

    /* Connection Settings Collapsible Expander in Sidebar */
    .stExpander {
        background-color: color-mix(in srgb, var(--text-color) 1%, transparent) !important;
        border: 1px solid color-mix(in srgb, var(--text-color) 6%, transparent) !important;
        border-radius: 8px !important;
        margin: 12px 0 !important;
        overflow: hidden !important;
        box-shadow: none !important;
    }
    .stExpander > details > summary {
        background: color-mix(in srgb, var(--text-color) 2%, transparent) !important;
        color: color-mix(in srgb, var(--text-color) 60%, transparent) !important;
        font-weight: 700 !important;
        font-size: 10px !important;
        letter-spacing: 0.08em !important;
        text-transform: uppercase !important;
        padding: 12px 16px !important;
        cursor: pointer !important;
    }
    .stExpander > details > summary:hover {
        color: var(--text-color) !important;
        background: color-mix(in srgb, var(--text-color) 4%, transparent) !important;
    }

    /* Hero Typography */
    .hero { text-align: center; padding: 60px 20px 40px; }
    .h1 {
        font-family: 'Outfit', sans-serif !important;
        font-size: 56px; font-weight: 800; letter-spacing: 0.04em; line-height: 1.1; margin-bottom: 20px;
        text-transform: uppercase;
        background: linear-gradient(135deg, #10b981, #3b82f6);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        display: inline-block;
    }
    .sub { font-size: 16px; color: color-mix(in srgb, var(--text-color) 60%, transparent); max-width: 600px; margin: 0 auto; line-height: 1.65; font-weight: 400; letter-spacing: 0.01em; }

    /* Upload Area & Cards */
    .unified-card {
        background: color-mix(in srgb, var(--secondary-background-color) 45%, transparent) !important;
        border: 1px solid color-mix(in srgb, var(--text-color) 8%, transparent) !important;
        border-radius: 12px !important;
        padding: 24px 28px !important;
        backdrop-filter: blur(12px) !important;
        -webkit-backdrop-filter: blur(12px) !important;
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.05) !important;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
        margin-bottom: 16px;
    }
    .unified-card:hover { 
        border-color: color-mix(in srgb, #3b82f6 25%, transparent) !important;
        box-shadow: 0 12px 40px color-mix(in srgb, #3b82f6 8%, transparent) !important;
        transform: translateY(-2px) !important;
    }
    .c-eye {
        font-family: 'Outfit', sans-serif !important;
        font-size: 10px; font-weight: 700; letter-spacing: 0.15em; text-transform: uppercase;
        color: #3b82f6; margin-bottom: 8px;
    }
    .c-ttl {
        font-family: 'Outfit', sans-serif !important;
        font-size: 16px; font-weight: 700; color: var(--text-color); margin-bottom: 20px;
        letter-spacing: 0.02em;
        border-left: 2px solid #3b82f6; padding-left: 12px; line-height: 1.3;
    }

    div[data-testid="stFileUploader"] { width: 100% !important; }
    div[data-testid="stFileUploader"] > section {
        background: color-mix(in srgb, var(--text-color) 1%, transparent) !important;
        border: 1px dashed color-mix(in srgb, var(--text-color) 10%, transparent) !important;
        border-radius: 8px !important;
        padding: 16px !important;
        transition: all 0.25s ease !important;
    }
    div[data-testid="stFileUploader"] > section:hover {
        border-color: #3b82f6 !important;
        background: color-mix(in srgb, #3b82f6 2%, transparent) !important;
    }
    div[data-testid="stFileUploader"] small { display: none !important; }

    /* Gradient Divider */
    .divider {
        height: 1px; margin: 20px 0 40px;
        background: linear-gradient(90deg, transparent, color-mix(in srgb, var(--text-color) 6%, transparent), transparent);
    }

    /* Primary Actions */
    button[data-testid="baseButton-primary"] {
        background: linear-gradient(135deg, #10b981, #3b82f6) !important; color: #fff !important; border: none !important;
        border-radius: 8px !important; font-size: 13px !important; font-weight: 700 !important; padding: 12px 28px !important;
        box-shadow: 0 4px 20px color-mix(in srgb, #3b82f6 20%, transparent) !important; transition: all 0.3s cubic-bezier(0.2, 0.8, 0.2, 1) !important;
        letter-spacing: 0.08em !important; text-transform: uppercase !important;
    }
    button[data-testid="baseButton-primary"]:hover { transform: translateY(-2px); box-shadow: 0 8px 30px color-mix(in srgb, #3b82f6 30%, transparent) !important; filter: brightness(1.1) !important; }
    button[data-testid="baseButton-primary"]:disabled { background: color-mix(in srgb, var(--text-color) 4%, transparent) !important; color: color-mix(in srgb, var(--text-color) 30%, transparent) !important; box-shadow: none !important; transform: none !important; border: 1px solid color-mix(in srgb, var(--text-color) 5%, transparent) !important; }

    button[data-testid="baseButton-secondary"] {
        background: color-mix(in srgb, var(--text-color) 2%, transparent) !important;
        color: var(--text-color) !important;
        border: 1px solid color-mix(in srgb, var(--text-color) 8%, transparent) !important;
        border-radius: 8px !important;
        font-size: 11px !important;
        font-weight: 700 !important;
        letter-spacing: 0.08em !important;
        text-transform: uppercase !important;
        transition: all 0.25s ease !important;
    }
    button[data-testid="baseButton-secondary"]:hover {
        background: color-mix(in srgb, var(--text-color) 5%, transparent) !important;
        border-color: color-mix(in srgb, var(--text-color) 15%, transparent) !important;
        color: var(--text-color) !important;
    }

    /* Multiselect / Dropdown Overrides */
    span[data-baseweb="tag"] {
        background-color: color-mix(in srgb, #3b82f6 10%, transparent) !important;
        color: #3b82f6 !important;
        border: 1px solid color-mix(in srgb, #3b82f6 25%, transparent) !important;
        border-radius: 4px !important;
        font-weight: 600 !important;
        letter-spacing: 0.02em;
    }
    span[data-baseweb="tag"] span { color: #3b82f6 !important; }
    li[role="option"] {
        color: var(--text-color) !important;
        background-color: var(--background-color) !important;
    }
    li[role="option"]:hover, li[role="option"][aria-selected="false"]:hover {
        background-color: color-mix(in srgb, var(--text-color) 4%, transparent) !important;
        color: var(--text-color) !important;
    }
    li[role="option"][aria-selected="true"] {
        background-color: color-mix(in srgb, #3b82f6 12%, transparent) !important;
        color: #3b82f6 !important;
        font-weight: 700 !important;
    }
    ul[data-testid="stMultiSelectDropdown"] {
        border: 1px solid color-mix(in srgb, var(--text-color) 8%, transparent) !important;
        border-radius: 8px !important;
        background-color: var(--background-color) !important;
        box-shadow: 0 12px 48px rgba(0, 0, 0, 0.15) !important;
        overflow: hidden;
    }

    /* Output Card & Tables */
    .output-card {
        background: color-mix(in srgb, var(--secondary-background-color) 45%, transparent) !important;
        border: 1px solid color-mix(in srgb, var(--text-color) 6%, transparent) !important; 
        border-radius: 16px; padding: 40px 36px; margin-top: 24px;
        backdrop-filter: blur(16px) !important;
        -webkit-backdrop-filter: blur(16px) !important;
        box-shadow: 0 16px 48px rgba(0, 0, 0, 0.08) !important;
        line-height: 1.8;
    }
    .output-card h3 {
        font-family: 'Outfit', sans-serif !important;
        font-size: 16px; font-weight: 700; margin: 32px 0 16px;
        color: var(--text-color);
        border-bottom: 1px solid color-mix(in srgb, var(--text-color) 8%, transparent); padding-bottom: 10px;
        letter-spacing: 0.03em;
        text-transform: uppercase;
    }
    .output-card p, .output-card li {
        font-size: 14px; color: color-mix(in srgb, var(--text-color) 85%, transparent); margin-bottom: 8px;
    }
    
    /* Policy Details Dropdown */
    details {
        background: color-mix(in srgb, var(--secondary-background-color) 40%, transparent) !important;
        border: 1px solid color-mix(in srgb, var(--text-color) 8%, transparent) !important;
        border-radius: 8px !important;
        margin-top: 14px !important;
        overflow: hidden !important;
        transition: all 0.3s ease !important;
    }
    details[open] {
        border-color: color-mix(in srgb, #3b82f6 30%, transparent) !important;
        box-shadow: 0 4px 24px color-mix(in srgb, #3b82f6 8%, transparent) !important;
    }
    summary {
        padding: 14px 20px !important;
        cursor: pointer !important;
        font-weight: 700 !important;
        font-size: 12px !important;
        color: #3b82f6 !important;
        background: color-mix(in srgb, #3b82f6 4%, transparent) !important;
        display: flex !important;
        align-items: center !important;
        gap: 10px !important;
        outline: none !important;
        list-style: none !important;
        user-select: none !important;
        transition: background 0.2s ease !important;
        letter-spacing: 0.05em !important;
        text-transform: uppercase !important;
    }
    summary:hover { background: color-mix(in srgb, #3b82f6 8%, transparent) !important; }
    summary::-webkit-details-marker { display: none; }
    details > *:not(summary) {
        padding: 16px 20px 20px !important;
        animation: slideDown 0.25s ease-out;
    }
    @keyframes slideDown {
        from { opacity: 0; transform: translateY(-8px); }
        to   { opacity: 1; transform: translateY(0); }
    }

    /* Outline/glass badges */
    .badge {
        display: inline-block; padding: 3px 10px; border-radius: 4px;
        font-size: 10px; font-weight: 700; text-transform: uppercase;
        letter-spacing: 0.06em;
    }
    .badge-fail     { background: rgba(239, 68, 68, 0.1) !important; color: #ef4444 !important; border: 1px solid rgba(239, 68, 68, 0.25) !important; }
    .badge-review   { background: rgba(245, 158, 11, 0.1) !important; color: #d97706 !important; border: 1px solid rgba(245, 158, 11, 0.25) !important; }
    .badge-verified { background: rgba(16, 185, 129, 0.1) !important; color: #059669 !important; border: 1px solid rgba(16, 185, 129, 0.25) !important; }

    /* Thin line section accents */
    .section-accent {
        border-left: 3px solid transparent !important;
        padding: 18px 22px !important; margin: 24px 0 !important;
        border-radius: 0 8px 8px 0 !important;
    }
    .accent-fail     { border-left-color: #ef4444 !important; background: rgba(239, 68, 68, 0.02) !important; }
    .accent-review   { border-left-color: #f59e0b !important; background: rgba(245, 158, 11, 0.02) !important; }
    .accent-verified { border-left-color: #10b981 !important; background: rgba(16, 185, 129, 0.02) !important; }

    /* Output sections inside Report */
    .output-section {
        background: color-mix(in srgb, var(--background-color) 40%, transparent) !important;
        border: 1px solid color-mix(in srgb, var(--text-color) 4%, transparent) !important;
        border-radius: 12px !important;
        padding: 24px !important;
        margin-bottom: 12px !important;
        line-height: 1.8;
    }
    .output-section h3 {
        background: color-mix(in srgb, var(--text-color) 2%, transparent) !important;
        border-left: 3px solid #3b82f6 !important;
        color: var(--text-color) !important;
        font-family: 'Outfit', sans-serif !important;
        font-size: 13px !important;
        font-weight: 700 !important;
        letter-spacing: 0.05em !important;
        text-transform: uppercase !important;
        padding: 8px 16px !important;
        margin: 28px 0 16px 0 !important;
        border-bottom: none !important;
        border-radius: 0 4px 4px 0 !important;
        display: block !important;
    }
    .output-section h3:first-child { margin-top: 0 !important; }
    .output-section p { color: color-mix(in srgb, var(--text-color) 85%, transparent); margin: 8px 0; font-size: 14px; }
    .output-section li { color: color-mix(in srgb, var(--text-color) 85%, transparent); margin: 6px 0; font-size: 14px; }
    .output-section table {
        width: 100% !important; border-collapse: collapse !important;
        border: 1px solid color-mix(in srgb, var(--text-color) 8%, transparent) !important;
        background: color-mix(in srgb, var(--background-color) 60%, transparent) !important;
        border-radius: 8px !important; overflow: hidden !important; margin: 16px 0 !important;
    }
    .output-section th {
        background: color-mix(in srgb, var(--text-color) 3%, transparent) !important; color: color-mix(in srgb, var(--text-color) 60%, transparent) !important;
        font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em; font-size: 10px;
        padding: 12px 16px; border-bottom: 1px solid color-mix(in srgb, var(--text-color) 8%, transparent) !important;
    }
    .output-section td {
        padding: 12px 16px; border-bottom: 1px solid color-mix(in srgb, var(--text-color) 6%, transparent) !important;
        vertical-align: top; color: var(--text-color); line-height: 1.65;
    }
    .output-section tr:last-child td { border-bottom: none; }
    .output-section tr:nth-child(even) td { background: color-mix(in srgb, var(--text-color) 0.5%, transparent); }
    .output-section code {
        background: rgba(239, 68, 68, 0.05) !important; color: #ef4444 !important;
        padding: 2px 6px; border-radius: 4px;
        font-family: 'JetBrains Mono', monospace; font-size: 12.5px;
        border: 1px solid rgba(239, 68, 68, 0.15) !important;
    }

    /* Score banner */
    .score-card {
        background: color-mix(in srgb, var(--secondary-background-color) 45%, transparent) !important;
        border: 1px solid color-mix(in srgb, var(--text-color) 6%, transparent) !important;
        border-top: 3px solid #3b82f6 !important;
        border-radius: 12px !important;
        padding: 20px 24px !important;
        margin-bottom: 16px !important;
    }
    .score-number { font-family: 'Outfit', sans-serif !important; font-size: 40px; font-weight: 800; color: #3b82f6; letter-spacing: -0.02em; line-height: 1; }
    .score-label { font-family: 'Outfit', sans-serif !important; font-size: 10px; font-weight: 700; letter-spacing: 0.12em; text-transform: uppercase; color: #3b82f6; margin-bottom: 4px; }
    
    .audit-done-banner {
        display: flex; align-items: center; gap: 12px;
        background: rgba(16, 185, 129, 0.05) !important; border: 1px solid rgba(16, 185, 129, 0.2) !important;
        border-left: 4px solid #10b981 !important;
        border-radius: 8px; padding: 12px 18px;
        color: #059669; font-size: 13.5px; font-weight: 600;
        margin-bottom: 20px; letter-spacing: 0.01em;
    }
    .audit-done-dot {
        width: 8px; height: 8px; border-radius: 50%;
        background: #10b981; flex-shrink: 0;
    }

    /* Code block inside details */
    [data-testid="stCodeBlock"] {
        background-color: var(--secondary-background-color) !important;
        border-radius: 8px !important;
        border: 1px solid color-mix(in srgb, var(--text-color) 8%, transparent) !important;
        margin: 0 !important;
    }
    [data-testid="stCodeBlock"] code {
        color: var(--text-color) !important;
        font-family: 'JetBrains Mono', monospace !important;
        font-size: 12.5px !important; line-height: 1.75 !important;
    }

    /* Modal Overlay (z-index 99999) */
    .fixed-overlay {
        position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
        background: color-mix(in srgb, var(--background-color) 70%, transparent);
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        z-index: 999998;
        animation: overlayFadeIn 0.35s ease-out forwards;
    }
    .fixed-modal {
        position: fixed; top: 45%; left: 50%; transform: translate(-50%, -50%);
        border-radius: 16px; padding: 48px; width: 500px; max-width: 90vw;
        box-shadow: 0 40px 80px rgba(0,0,0,0.15), inset 0 1px 0 color-mix(in srgb, var(--text-color) 8%, transparent) !important;
        text-align: center; z-index: 999999;
        background: color-mix(in srgb, var(--secondary-background-color) 90%, transparent) !important;
        border: 1px solid color-mix(in srgb, var(--text-color) 8%, transparent) !important;
        color: var(--text-color) !important;
        animation: modalSlideIn 0.4s cubic-bezier(0.16, 1, 0.3, 1) forwards;
    }
    .fixed-modal h3 { 
        font-family: 'Outfit', sans-serif !important;
        color: var(--text-color) !important; margin: 0 0 12px; font-weight: 700; font-size: 20px;
        letter-spacing: 0.03em; text-transform: uppercase;
    }
    .fixed-modal p  { color: color-mix(in srgb, var(--text-color) 60%, transparent) !important; margin: 0; font-size: 13px; letter-spacing: 0.02em; }
    .spinner-loader {
        border: 2px solid color-mix(in srgb, var(--text-color) 5%, transparent); border-top: 2px solid #3b82f6; border-radius: 50%;
        width: 40px; height: 40px; animation: spin 0.8s cubic-bezier(0.4, 0, 0.2, 1) infinite; margin: 0 auto 24px auto;
    }
    .progress-container {
        width: 100%; height: 4px; background: color-mix(in srgb, var(--text-color) 5%, transparent); border-radius: 10px; margin: 24px 0 12px; overflow: hidden;
    }
    .progress-fill {
        height: 100%; background: linear-gradient(90deg, #10b981, #3b82f6); transition: width 0.4s cubic-bezier(0.1, 0.8, 0.1, 1);
    }
    .perc-text { font-family: 'Outfit', sans-serif !important; font-size: 32px; font-weight: 800; color: #3b82f6; margin-bottom: 4px; }

    /* Cancellation Button Float Container */
    .cancel-btn-container {
        position: fixed !important;
        top: calc(45% + 120px) !important;
        left: 50% !important;
        transform: translateX(-50%) !important;
        z-index: 1000000 !important;
        width: 220px !important;
        text-align: center !important;
    }
    .cancel-btn-container button {
        background: rgba(239, 68, 68, 0.1) !important;
        color: #ef4444 !important;
        border: 1px solid rgba(239, 68, 68, 0.3) !important;
        border-radius: 8px !important;
        font-size: 11px !important;
        font-weight: 700 !important;
        letter-spacing: 0.08em !important;
        text-transform: uppercase !important;
        padding: 10px 20px !important;
        transition: all 0.2s ease !important;
        width: 100% !important;
    }
    .cancel-btn-container button:hover {
        background: rgba(239, 68, 68, 0.2) !important;
        border-color: #ef4444 !important;
        color: #ffffff !important;
        box-shadow: 0 0 15px rgba(239, 68, 68, 0.25) !important;
    }

    /* Entrance Animation */
    @keyframes slideInUp {
        from { opacity: 0; transform: translateY(40px); }
        to { opacity: 1; transform: translateY(0); }
    }
    .first-run-anim {
        opacity: 0;
        animation: slideInUp 0.8s cubic-bezier(0.2, 0.8, 0.2, 1) forwards;
    }
    .anim-delay-1 { animation-delay: 0.1s; }
    .anim-delay-2 { animation-delay: 0.25s; }
    .anim-delay-3 { animation-delay: 0.4s; }

    @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
    @keyframes overlayFadeIn { 0% { opacity: 0; } 100% { opacity: 1; } }
    @keyframes modalSlideIn { 0% { opacity: 0; transform: translate(-50%, -40%) scale(0.95); } 100% { opacity: 1; transform: translate(-50%, -45%) scale(1); } }
    
    [data-testid="stStatusWidget"] { display: none !important; }
    </style>
    """, unsafe_allow_html=True)

# Call CSS setup
load_css()

# --- API Key Management (C3 Secure Key Environment) ---
# Read cloud secret ONCE at startup (avoids double-warning on local machines)
_CLOUD_KEY = ""
try:
    _CLOUD_KEY = st.secrets.get("GEMINI_KEY", "") or ""
except Exception:
    pass  # No secrets.toml — running locally, that's fine

def load_key():
    if _CLOUD_KEY:
        return _CLOUD_KEY                       # Cloud deployment
    if os.path.exists(".env"):
        try:
            with open(".env", "r", encoding="utf-8") as f:
                content = f.read()
            match = re.search(r'^\s*GEMINI_KEY\s*=\s*(["\']?)(.*?)\1\s*$', content, re.MULTILINE)
            if match:
                return match.group(2).strip()
        except Exception:
            pass
    # Backward compatibility fallback to .gemini_key
    if os.path.exists(".gemini_key"):
        try:
            with open(".gemini_key", "r", encoding="utf-8") as f:
                return f.read().strip()
        except Exception:
            pass
    return ""

def save_key(k):
    try:
        lines = []
        if os.path.exists(".env"):
            with open(".env", "r", encoding="utf-8") as f:
                lines = f.readlines()
        
        key_found = False
        new_lines = []
        for line in lines:
            if re.match(r'^\s*GEMINI_KEY\s*=', line):
                new_lines.append(f"GEMINI_KEY={k.strip()}\n")
                key_found = True
            else:
                new_lines.append(line)
        
        if not key_found:
            new_lines.append(f"GEMINI_KEY={k.strip()}\n")
            
        with open(".env", "w", encoding="utf-8") as f:
            f.writelines(new_lines)
            
        # Clean up old .gemini_key to keep environment clean
        if os.path.exists(".gemini_key"):
            try:
                os.remove(".gemini_key")
            except Exception:
                pass
    except Exception:
        pass  # Read-only filesystem (Streamlit Cloud) — ignored

def is_cloud_key():
    return bool(_CLOUD_KEY)

# ─── apply_badges — top-level utility (used by both streaming & saved-report) ─
def apply_badges(text):
    import re as _re
    text = text.replace("&lt;div class=\"section-accent accent-fail\"&gt;", '<div class="section-accent accent-fail">')
    text = text.replace("&lt;div class=\"section-accent accent-review\"&gt;", '<div class="section-accent accent-review">')
    text = text.replace("&lt;div class=\"section-accent accent-verified\"&gt;", '<div class="section-accent accent-verified">')
    text = text.replace("&lt;/div&gt;", '</div>')
    text = text.replace("&lt;span class=\"badge", '<span class="badge')
    text = text.replace("&lt;/span&gt;", '</span>')
    text = text.replace("[SECTION_FAIL]", '\n\n<div class="section-accent accent-fail">\n\n')
    text = text.replace("[SECTION_REVIEW]", '\n\n</div>\n\n<div class="section-accent accent-review">\n\n')
    text = text.replace("[SECTION_VERIFIED]", '\n\n</div>\n\n<div class="section-accent accent-verified">\n\n')
    text = text.replace("[FAIL]", '<span class="badge badge-fail">FAIL</span>')
    text = text.replace("[REVIEW]", '<span class="badge badge-review">REVIEW</span>')
    text = text.replace("[VERIFIED]", '<span class="badge badge-verified">VERIFIED</span>')
    if '<div class="section-accent' in text and text.rstrip()[-6:] != '</div>':
        text += "\n\n</div>\n\n"
    # Spacing fix: ensure blank line before each bullet
    text = _re.sub(r'(?<!\n)\n(• |\* |-  ?(?=\S))', r'\n\n\1', text)
    return text

with st.sidebar:
    st.markdown("<div style='font-size: 20px; font-weight: 800; background: linear-gradient(90deg, #10b981, #3b82f6); -webkit-background-clip: text; -webkit-text-fill-color: transparent;'>EXCEL AUDITOR</div>", unsafe_allow_html=True)
    st.markdown("<div style='font-size: 12px; font-weight: 500; opacity: 0.5; margin-bottom: 24px;'>Data Recheck Tool</div>", unsafe_allow_html=True)
    
    st.markdown("<div style='font-size: 10px; font-weight: 700; letter-spacing: 0.1em; text-transform: uppercase; opacity: 0.4; margin-bottom: 8px;'>NAVIGATION</div>", unsafe_allow_html=True)
    page_options = ["CONTRACT AUDITOR", "AI EXCEL GENERATOR"]
    selected_page = st.radio("Navigation", page_options, label_visibility="collapsed")
    
    st.markdown("<br>", unsafe_allow_html=True)
    saved_key = load_key()

    with st.expander("CONNECTION SETTINGS", expanded=not saved_key):
        if is_cloud_key():
            api_key = saved_key
            st.caption("API KEY: CONFIGURED VIA SECRETS")
        else:
            api_key = st.text_input("GEMINI API KEY", type="password", value=saved_key, placeholder="Enter API Key...", label_visibility="collapsed")
            if api_key:
                if api_key != saved_key:
                    save_key(api_key)
                ok, msg = validate_api_key(api_key)
                if ok:
                    st.caption(f"STATUS: {msg}")
                else:
                    st.error("INVALID API KEY")
            else:
                st.caption("API KEY REQUIRED FOR ANALYSIS")



# ─── State Control for Streamlit Re-runs ──────────────────────────────────────
if "has_loaded" not in st.session_state:
    st.session_state.has_loaded = True
    anim_class = "first-run-anim"
else:
    anim_class = ""

# ─── Hero ─────────────────────────────────────────────────────────────────────
if selected_page == "CONTRACT AUDITOR":
    st.markdown(f"""
    <div class="hero {anim_class}">
      <div class="h1">DATA AUDITOR</div>
      <div class="sub">Precision Hotel Contract Verification.</div>
    </div>
    <div class="divider"></div>
    """, unsafe_allow_html=True)
else:
    st.markdown(f"""
    <div class="hero {anim_class}">
      <div class="h1">EXCEL GENERATOR</div>
      <div class="sub">Generate Hotel Upload Files from PDF.</div>
    </div>
    <div class="divider"></div>
    """, unsafe_allow_html=True)

# ─── PAGE ROUTING ──────────────────────────────────────────────────────────────
if selected_page == "AI EXCEL GENERATOR":
    col1, col2 = st.columns(2, gap="large")

    with col1:
        st.markdown(f'<div class="unified-card {anim_class} anim-delay-1"><div class="c-eye">STEP 1</div><div class="c-ttl">Contract PDF</div>', unsafe_allow_html=True)
        pdf_file_gen = st.file_uploader("Upload PDF for Generation", type=["pdf"], key="pdf_gen", label_visibility="collapsed")
        if pdf_file_gen:
            st.session_state.cached_pdf_bytes = pdf_file_gen.getvalue()
            st.session_state.cached_pdf_name = pdf_file_gen.name
            st.markdown(f'<div style="font-size:12px;color:#10b981;font-weight:600;margin-top:8px;">{pdf_file_gen.name}</div>', unsafe_allow_html=True)
        elif st.session_state.get("cached_pdf_name"):
            st.markdown(f'<div style="font-size:12px;color:#3b82f6;font-weight:600;margin-top:8px;">CACHED: {st.session_state.cached_pdf_name}</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with col2:
        st.markdown(f'<div class="unified-card {anim_class} anim-delay-2"><div class="c-eye">STEP 2 (OPTIONAL)</div><div class="c-ttl">Reference Excel</div>', unsafe_allow_html=True)
        excel_ref_file = st.file_uploader("Upload Reference Excel", type=["xlsx", "xls"], key="excel_gen", label_visibility="collapsed")
        if excel_ref_file:
            st.session_state.cached_excel_bytes = excel_ref_file.getvalue()
            st.session_state.cached_excel_name = excel_ref_file.name
            st.markdown(f'<div style="font-size:12px;color:#10b981;font-weight:600;margin-top:8px;">{excel_ref_file.name}</div>', unsafe_allow_html=True)
        elif st.session_state.get("cached_excel_name"):
            st.markdown(f'<div style="font-size:12px;color:#3b82f6;font-weight:600;margin-top:8px;">CACHED: {st.session_state.cached_excel_name}</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    
    # Process Button
    has_pdf_gen = bool(pdf_file_gen or st.session_state.get("cached_pdf_bytes"))
    gen_ready = bool(has_pdf_gen and api_key)
    
    _, btn_col, _ = st.columns([1.5, 3, 1.5])
    with btn_col:
        if st.button("Generate Upload Excel from PDF", type="primary", use_container_width=True, disabled=not gen_ready):
            with st.spinner("AI is analyzing and generating Excel..."):
                try:
                    from utils import extract_pdf_to_excel_json, create_upload_excel
                    pdf_bytes_for_gen = pdf_file_gen.getvalue() if pdf_file_gen else st.session_state.get("cached_pdf_bytes")
                    excel_bytes_for_gen = excel_ref_file.getvalue() if excel_ref_file else st.session_state.get("cached_excel_bytes")
                    
                    result = extract_pdf_to_excel_json(pdf_bytes_for_gen, api_key, excel_bytes=excel_bytes_for_gen)
                    
                    extracted_data = result[0] if isinstance(result, tuple) else result
                    error_detail = result[1] if isinstance(result, tuple) and len(result) > 1 else None

                    if extracted_data:
                        excel_data = create_upload_excel(extracted_data)
                        st.success("Excel generated successfully!")
                        st.download_button(
                            label="DOWNLOAD GENERATED UPLOAD FILE",
                            data=excel_data,
                            file_name=f"Generated_Upload_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    else:
                        st.error("AI could not extract data. Please try again.")
                        if error_detail:
                            st.caption("TECHNICAL DETAILS (DEBUG):")
                            st.code(error_detail, language="bash")
                except Exception as e:
                    st.error(f"Generation failed: {str(e)}")

    if not gen_ready:
        hint = "Upload PDF file to continue" if not has_pdf_gen else "Enter API Key in sidebar"
        st.markdown(f"<p style='text-align:center;color:#94a3b8;font-size:12px;margin-top:6px;letter-spacing:0.02em'>{hint}</p>", unsafe_allow_html=True)
        
    st.stop()

# ─── Upload Area ───────────────────────────────────────────────────────────────
_audit_done = st.session_state.get("audit_done", False) and bool(st.session_state.get("_audit_result"))
_is_auditing = st.session_state.get("is_auditing", False)

# Render clean action buttons in top-right when audit is done
if _audit_done:
    _left, _right1, _right2 = st.columns([4, 1.5, 1.5])
    with _right1:
        if st.button("EDIT SCOPE / RE-AUDIT", use_container_width=True):
            for key in ["audit_done", "is_auditing", "_audit_result"]:
                st.session_state.pop(key, None)
            st.rerun()
    with _right2:
        if st.button("START FRESH", use_container_width=True):
            st.session_state.confirm_reset = True
            st.rerun()

# ─── Decoupled Upload Panel & Scope Selector (H5) ───────────────────────────
with st.expander("UPLOAD & SCOPE SETTINGS", expanded=not _audit_done):
    col1, col2 = st.columns(2, gap="large")

    with col1:
        st.markdown(f'<div class="unified-card {anim_class} anim-delay-1"><div class="c-eye">STEP 1</div><div class="c-ttl">Contract PDF</div>', unsafe_allow_html=True)
        pdf_file = st.file_uploader("Upload PDF", type=["pdf"], key="pdf", label_visibility="collapsed")
        if pdf_file:
            st.session_state.cached_pdf_bytes = pdf_file.getvalue()
            st.session_state.cached_pdf_name = pdf_file.name
            st.markdown(f'<div style="font-size:12px;color:#10b981;font-weight:600;margin-top:8px;">{pdf_file.name}</div>', unsafe_allow_html=True)
        elif st.session_state.get("cached_pdf_name"):
            st.markdown(f'<div style="font-size:12px;color:#3b82f6;font-weight:600;margin-top:8px;">CACHED: {st.session_state.cached_pdf_name}</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with col2:
        st.markdown(f'<div class="unified-card {anim_class} anim-delay-2"><div class="c-eye">STEP 2</div><div class="c-ttl">Data Excel</div>', unsafe_allow_html=True)
        excel_file = st.file_uploader("Upload Excel", type=["xlsx", "xls"], key="excel", label_visibility="collapsed")
        if excel_file:
            st.session_state.cached_excel_bytes = excel_file.getvalue()
            st.session_state.cached_excel_name = excel_file.name
            st.markdown(f'<div style="font-size:12px;color:#10b981;font-weight:600;margin-top:8px;">{excel_file.name}</div>', unsafe_allow_html=True)
        elif st.session_state.get("cached_excel_name"):
            st.markdown(f'<div style="font-size:12px;color:#3b82f6;font-weight:600;margin-top:8px;">CACHED: {st.session_state.cached_excel_name}</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    # ─── Audit Focus ──────────────────────────────────────────────────────────
    _, focus_col, _ = st.columns([0.5, 5, 0.5])
    with focus_col:
        st.markdown(f'<div class="unified-card {anim_class} anim-delay-3"><div class="c-eye">STEP 3</div><div class="c-ttl">Audit Scope</div>', unsafe_allow_html=True)
        individual_options = [
            "Net Price & Extra Beds",
            "Cancellation Policy",
            "Child Policy",
            "Period & Seasons",
            "Meals & Info",
        ]
        all_in_one = "All-in-One Full Scan"
        focus_options = individual_options + [all_in_one]

        if "prev_focus" not in st.session_state:
            st.session_state.prev_focus = [all_in_one]

        selected_focus = st.multiselect(
            "Audit scope", options=focus_options,
            default=st.session_state.prev_focus, label_visibility="collapsed"
        )
        prev = st.session_state.prev_focus
        if all_in_one in selected_focus and all_in_one not in prev:
            selected_focus = [all_in_one]
            st.session_state.prev_focus = selected_focus
            st.rerun()
        elif all_in_one in selected_focus and any(o in selected_focus for o in individual_options) and any(o not in prev for o in selected_focus if o != all_in_one):
            selected_focus = [o for o in selected_focus if o != all_in_one]
            st.session_state.prev_focus = selected_focus
            st.rerun()
        st.session_state.prev_focus = selected_focus
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ─── Process Button ────────────────────────────────────────────────────────
    has_pdf = bool(pdf_file or st.session_state.get("cached_pdf_bytes"))
    has_excel = bool(excel_file or st.session_state.get("cached_excel_bytes"))
    ready = bool(has_pdf and has_excel and api_key)
    
    _, btn_col, _ = st.columns([1.5, 3, 1.5])
    with btn_col:
        if st.button("Start Audit", type="primary", use_container_width=True, disabled=not ready):
            st.session_state.is_auditing = True
            st.session_state.audit_done = False
            st.session_state.focus_list = selected_focus
            st.rerun()

    if not ready and not st.session_state.get("is_auditing"):
        hint = "Upload PDF and Excel files to continue" if not (has_pdf and has_excel) else "Enter API Key in sidebar"
        st.markdown(f"<p style='text-align:center;color:#94a3b8;font-size:12px;margin-top:6px;letter-spacing:0.02em'>{hint}</p>", unsafe_allow_html=True)




st.markdown('<div class="divider"></div>', unsafe_allow_html=True)




# ─── Analysis Output ──────────────────────────────────────────────────────────
if st.session_state.get("is_auditing"):
    st.markdown('<div style="font-size:10px;font-weight:800;letter-spacing:0.15em;text-transform:uppercase;color:#94a3b8;margin-bottom:20px;padding-bottom:12px;border-bottom:1px solid #e8edf2">AUDIT REPORT</div>', unsafe_allow_html=True)
    
    report_placeholder = st.empty()
    modal_placeholder = st.empty()
    cancel_placeholder = st.empty()
    
    # Initial Loading State
    focus_text = ", ".join(st.session_state.get("focus_list", [])) or "Full Scan"

    def render_modal(perc, phase):
        modal_placeholder.markdown(f"""
            <div class="fixed-overlay"></div>
            <div class="fixed-modal">
                <div class="spinner-loader"></div>
                <div class="perc-text">{int(perc)}%</div>
                <h3>{phase}</h3>
                <p>Target: {focus_text}</p>
                <div class="progress-container"><div class="progress-fill" style="width: {perc}%"></div></div>
            </div>
        """, unsafe_allow_html=True)

    render_modal(5, "Initializing AI Engine...")
    
    # Centered floating cancel button above overlay
    cancel_placeholder.markdown('<div class="cancel-btn-container">', unsafe_allow_html=True)
    if cancel_placeholder.button("CANCEL AUDIT", key="cancel_btn_trigger"):
        st.session_state.cancel_requested = True
        st.rerun()

    full_response = ""
    chunk_count = 0
    try:
        # Resolve files dynamically
        pdf_bytes_to_use = st.session_state.get("cached_pdf_bytes")
        excel_bytes_to_use = st.session_state.get("cached_excel_bytes")
        
        for chunk in stream_recheck_analysis(pdf_bytes_to_use, excel_bytes_to_use, api_key, st.session_state.focus_list):
            chunk_count += 1
            
            # Dynamic honest milestone parsing
            current_perc, current_phase = get_honest_milestone(full_response, chunk_count)
            render_modal(current_perc, current_phase)
            
            # Re-draw the cancel button to keep it clickable in the event-loop
            cancel_placeholder.markdown('<div class="cancel-btn-container">', unsafe_allow_html=True)
            if cancel_placeholder.button("CANCEL AUDIT", key="cancel_btn_trigger_" + str(chunk_count)):
                st.session_state.cancel_requested = True
                st.rerun()

            if chunk == "[RESET_STREAM]":
                full_response = ""
                continue
            full_response += chunk
            
            # Transform and display
            display_response = apply_badges(full_response)
            report_placeholder.markdown(f'<div class="output-card">\n\n{display_response}▌\n\n</div>', unsafe_allow_html=True)
        
        render_modal(100, "Audit Complete")
        modal_placeholder.empty()
        cancel_placeholder.empty()
        
        # Detect if the response is an error (quota exceeded, invalid key, etc.)
        _is_error = any(kw in full_response for kw in [
            "API Key ไม่ถูกต้อง", "โควต้า API", "ไม่สามารถเชื่อมต่อ",
            "RESOURCE_EXHAUSTED", "INVALID_ARGUMENT", "API_KEY_INVALID"
        ])
        
        if _is_error:
            # Show error inline — keep upload section visible so user can retry
            report_placeholder.empty()
            st.session_state.is_auditing = False
            st.error(full_response.replace("**", "").strip())
        else:
            # Successful audit — collapse upload and show report
            st.session_state.audit_done = True
            st.session_state.is_auditing = False
            st.session_state._audit_result = full_response
            st.rerun()
        
    except Exception as e:
        modal_placeholder.empty()
        st.error(f"Error: {str(e)}")
        st.session_state.is_auditing = False

# ─── Show Saved Report (after rerun with audit_done=True) ────────────────────
if st.session_state.get("audit_done") and not st.session_state.get("is_auditing"):
    saved_result = st.session_state.get("_audit_result", "")
    if saved_result:
        st.markdown('<div style="font-size:10px;font-weight:800;letter-spacing:0.15em;text-transform:uppercase;color:#94a3b8;margin-bottom:20px;padding-bottom:12px;border-bottom:1px solid #e8edf2">AUDIT REPORT</div>', unsafe_allow_html=True)
        # apply_badges is now a top-level function — accessible here directly
        
        def _show_saved_report(full_text):
            st.markdown(
                '<div class="audit-done-banner"><div class="audit-done-dot"></div>Audit complete — report ready for review.</div>',
                unsafe_allow_html=True
            )
            st.toast("Audit Process Completed Successfully")
            st.markdown('<a id="audit-result-section"></a>', unsafe_allow_html=True)
            
            score_match = re.search(r'คะแนนความถูกต้อง.*?(\d+(?:\.\d+)?)\s*%', full_text)
            summary_match_text = re.search(r'บทสรุป.*?:\s*(.+)', full_text)
            if score_match:
                score = float(score_match.group(1))
                color = "#10b981" if score >= 90 else "#f59e0b" if score >= 70 else "#ef4444"
                summary_txt = summary_match_text.group(1).strip() if summary_match_text else ""
                st.markdown(f"""
                <div style="background: linear-gradient(135deg, {color}15, {color}08); border: 2px solid {color}40; border-radius: 16px; padding: 28px 32px; margin: 24px 0; text-align: center;">
                    <div style="font-size: 52px; font-weight: 900; color: {color}; line-height: 1;">{score:.0f}%</div>
                    <div style="font-size: 13px; font-weight: 700; color: {color}; margin-top: 4px; text-transform: uppercase; letter-spacing: 0.08em;">Accuracy Score</div>
                    {"<div style='font-size:14px;color:#64748b;margin-top:12px;'>"+summary_txt+"</div>" if summary_txt else ""}
                </div>""", unsafe_allow_html=True)
            
            processed = apply_badges(full_text)
            parts = re.split(r'(?=<details>)|(?<=</details>)', processed)
            for part in parts:
                stripped = part.strip()
                if not stripped:
                    continue
                if stripped.startswith('<details>'):
                    summary_m = re.search(r'<summary>([\s\S]*?)</summary>', part)
                    label = "HTML Code — คลิกเพื่อดู / Copy"
                    if summary_m:
                        raw_label = summary_m.group(1)
                        label = re.sub(r'<[^>]+>', '', raw_label).strip()
                    code_match = re.search(r'```html\s*([\s\S]*?)```', part)
                    html_code = code_match.group(1).strip() if code_match else ""
                    if not html_code:
                        fallback = re.search(r'</summary>([\s\S]*?)</details>', part)
                        html_code = fallback.group(1).strip() if fallback else ""
                    with st.expander(label, expanded=False):
                        if html_code:
                            st.code(html_code, language="html")
                else:
                    st.markdown(f'<div class="output-section">{stripped}</div>', unsafe_allow_html=True)
        
        _show_saved_report(saved_result)
    
    st.markdown("<br>", unsafe_allow_html=True)
    _, reaudit_btn, reset_btn, _ = st.columns([1, 2, 2, 1])
    with reaudit_btn:
        if st.button("EDIT SCOPE / RE-AUDIT", use_container_width=True):
            for key in ["audit_done", "is_auditing", "_audit_result"]:
                st.session_state.pop(key, None)
            st.rerun()
    with reset_btn:
        if st.button("RESET / NEW UPLOAD", use_container_width=True):
            st.session_state.confirm_reset = True
            st.rerun()