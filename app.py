import os
import re
import json
import copy
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

# ─── Clean & Repair JSON helper ───────────────────────────────────────────────
def _clean_json(raw: str) -> str:
    text = re.sub(r"```(?:json)?", "", raw).strip()
    s, e = text.find("{"), text.rfind("}")
    if s != -1 and e > s:
        text = text[s: e + 1]
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", text)
    return text

def _repair_json(text: str) -> str:
    if not text: return text
    text = re.sub(r',\s*([\]\}])', r'\1', text)
    text = re.sub(r',?\s*"[^"]*"\s*:\s*[^,}\]]*$', '', text)
    text = re.sub(r',?\s*"[^"]*"\s*:\s*$', '', text)
    text = re.sub(r',?\s*$', '', text)
    
    stack = []
    for char in text:
        if char == '{': stack.append('}')
        elif char == '[': stack.append(']')
        elif char in ('}', ']'):
            if stack and stack[-1] == char:
                stack.pop()
    if stack:
        text += "".join(reversed(stack))
    return text

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
        <div class="fixed-modal" style="border-color: rgba(239, 68, 68, 0.4) !important; background: var(--secondary-background-color) !important; background: color-mix(in srgb, var(--secondary-background-color) 95%, transparent) !important;">
            <h3 style="color: #ef4444 !important;">ยืนยันการล้างข้อมูล</h3>
            <p style="margin-bottom: 24px; color: rgba(130, 130, 130, 0.6) !important;">คุณแน่ใจหรือไม่ว่าต้องการล้างข้อมูลและไฟล์ที่อัปโหลดทั้งหมด? การดำเนินการนี้ไม่สามารถย้อนกลับได้</p>
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

# ─── Dynamic Module Loading for COMPARE CONTRACT ─────────────────────────────
import importlib.util

if "active_app" not in st.session_state:
    st.session_state.active_app = "DATA AUDITOR"

if "app_selector_open" not in st.session_state:
    st.session_state.app_selector_open = False

compare_utils = None
compare_excel = None
try:
    _base_dir = os.path.dirname(__file__)
    # Search for the utils.py file in multiple possible deployment structures
    _possible_dirs = [
        os.path.abspath(os.path.join(_base_dir, "../COMPARE CONTRACT")), # Local Windows setup
        os.path.abspath(os.path.join(_base_dir, "COMPARE CONTRACT")),    # Cloud subfolder setup
        os.path.abspath(_base_dir)                                       # Cloud root dump setup
    ]
    
    _compare_dir = None
    for d in _possible_dirs:
        if os.path.exists(os.path.join(d, "utils2.py")):
            _compare_dir = d
            break
            
    if _compare_dir:
        _utils_path = os.path.join(_compare_dir, "utils2.py")
        _spec_utils = importlib.util.spec_from_file_location("compare_utils", _utils_path)
        compare_utils = importlib.util.module_from_spec(_spec_utils)
        _spec_utils.loader.exec_module(compare_utils)
        
        _excel_path = os.path.join(_compare_dir, "excel_generator.py")
        if os.path.exists(_excel_path):
            _spec_excel = importlib.util.spec_from_file_location("compare_excel", _excel_path)
            compare_excel = importlib.util.module_from_spec(_spec_excel)
            _spec_excel.loader.exec_module(compare_excel)
except Exception as e:
    pass

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
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:ital,wght@0,300;0,400;0,500;0,600;0,700;0,800;1,400&family=JetBrains+Mono:wght@400;500;600&display=swap');

    /* ── Reset & Base ──────────────────────────────────────────── */
    html { scroll-behavior: smooth !important; }
    *, *::before, *::after { box-sizing: border-box; }
    html, body, p, li, th, td, h1, h2, h3, h4, h5, h6,
    div[data-testid="stMarkdownContainer"] p {
        font-family: 'Plus Jakarta Sans', sans-serif !important;
        color: var(--text-color);
    }
    ::-webkit-scrollbar { width: 4px; }
    ::-webkit-scrollbar-track { background: transparent; }
    ::-webkit-scrollbar-thumb { background: rgba(100,116,139,0.2); border-radius: 99px; }

    /* ── App Background ────────────────────────────────────────── */
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {
        background-color: #f1f5f9 !important;
        color: var(--text-color) !important;
    }
    .block-container { padding: 2rem 2.5rem 4rem !important; max-width: 1100px !important; }
    [data-testid="stStatusWidget"] { display: none !important; }
    div[data-testid="stSpinner"], div[class*="stSpinner"] p, .stSpinner,
    div[data-testid="stSpinner"] > div > div { color: var(--text-color) !important; }

    /* ── Sidebar ───────────────────────────────────────────────── */
    section[data-testid="stSidebar"] {
        background-color: #ffffff !important;
        border-right: 1px solid #f1f5f9 !important;
    }
    section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] {
        padding: 1.25rem 0.75rem !important;
    }

    /* ── Nav Buttons (replaces broken Radio) ──────────────────── */
    .nav-label {
        font-size: 10px; font-weight: 700; letter-spacing: 0.08em;
        text-transform: uppercase; color: #cbd5e1;
        padding: 0 4px; margin-bottom: 4px; margin-top: 4px;
    }
    .nav-btn-wrap div.stButton > button {
        display: flex !important; align-items: center !important;
        gap: 8px !important; padding: 9px 12px !important;
        border-radius: 10px !important; font-size: 13px !important;
        font-weight: 500 !important; color: #94a3b8 !important;
        background: transparent !important; border: none !important;
        width: 100% !important; text-align: left !important;
        justify-content: flex-start !important; transition: all 0.15s !important;
        letter-spacing: -0.01em !important; box-shadow: none !important;
        font-family: 'Plus Jakarta Sans', sans-serif !important;
    }
    .nav-btn-wrap div.stButton > button:hover {
        background: #f8fafc !important; color: #334155 !important; transform: none !important;
    }
    .nav-btn-active div.stButton > button {
        background: linear-gradient(135deg, #eef2ff, #e0e7ff) !important;
        color: #4f46e5 !important; font-weight: 600 !important;
    }
    .nav-btn-active div.stButton > button:hover {
        background: linear-gradient(135deg, #e0e7ff, #c7d2fe) !important;
        color: #4338ca !important;
    }

    /* ── Sidebar Download Buttons ──────────────────────────────── */
    section[data-testid="stSidebar"] .stDownloadButton button {
        background-color: #f8fafc !important; border: 1px solid #f1f5f9 !important;
        color: #64748b !important; border-radius: 8px !important;
        font-size: 11px !important; font-weight: 600 !important;
        letter-spacing: 0.03em !important; text-transform: uppercase !important;
        padding: 8px 14px !important; width: 100% !important;
        transition: all 0.15s !important; font-family: 'Plus Jakarta Sans', sans-serif !important;
    }
    section[data-testid="stSidebar"] .stDownloadButton button:hover {
        background-color: #eef2ff !important; border-color: #c7d2fe !important; color: #4f46e5 !important;
    }

    /* ── Expander ──────────────────────────────────────────────── */
    .stExpander {
        background-color: #fafafa !important; border: 1px solid #f1f5f9 !important;
        border-radius: 10px !important; margin: 10px 0 !important;
        overflow: hidden !important; box-shadow: none !important;
    }
    .stExpander > details > summary {
        background: #f8fafc !important; color: #94a3b8 !important;
        font-weight: 700 !important; font-size: 10px !important;
        letter-spacing: 0.08em !important; text-transform: uppercase !important;
        padding: 12px 16px !important; cursor: pointer !important;
        font-family: 'Plus Jakarta Sans', sans-serif !important;
    }
    .stExpander > details > summary:hover { color: #334155 !important; background: #f1f5f9 !important; }

    /* ── File Uploader ─────────────────────────────────────────── */
    div[data-testid="stFileUploader"] { width: 100% !important; }
    div[data-testid="stFileUploader"] > section {
        background: #fafafa !important; border: 1.5px dashed #e2e8f0 !important;
        border-radius: 12px !important; transition: all 0.2s !important;
    }
    div[data-testid="stFileUploader"] > section:hover {
        border-color: #a5b4fc !important; background: #f5f3ff !important;
    }
    div[data-testid="stFileUploader"] small { display: none !important; }

    /* ── Primary Button ────────────────────────────────────────── */
    button[kind="primary"] {
        background: linear-gradient(135deg, #4f46e5, #6366f1) !important;
        color: #fff !important; border: none !important; border-radius: 10px !important;
        font-size: 13px !important; font-weight: 700 !important; padding: 11px 26px !important;
        box-shadow: 0 4px 14px rgba(79,70,229,0.25) !important;
        transition: all 0.15s !important; letter-spacing: -0.01em !important;
        font-family: 'Plus Jakarta Sans', sans-serif !important;
    }
    button[kind="primary"]:hover {
        filter: brightness(1.08) !important; transform: translateY(-1px) !important;
        box-shadow: 0 6px 20px rgba(79,70,229,0.35) !important;
    }
    button[kind="primary"]:disabled {
        background: #e2e8f0 !important; color: #94a3b8 !important;
        box-shadow: none !important; transform: none !important; border: 1px solid #f1f5f9 !important;
    }
    button[kind="secondary"] {
        background: #ffffff !important; color: #64748b !important;
        border: 1.5px solid #e2e8f0 !important; border-radius: 9px !important;
        font-size: 12px !important; font-weight: 500 !important;
        transition: all 0.15s !important; font-family: 'Plus Jakarta Sans', sans-serif !important;
    }
    button[data-testid="baseButton-secondary"]:hover { border-color: #94a3b8 !important; color: #334155 !important; }

    /* ── Tags ──────────────────────────────────────────────────── */
    span[data-baseweb="tag"] {
        background-color: rgba(79,70,229,0.08) !important; color: #4f46e5 !important;
        border: 1px solid rgba(79,70,229,0.2) !important; border-radius: 5px !important; font-weight: 600 !important;
    }
    span[data-baseweb="tag"] span { color: #4f46e5 !important; }
    li[role="option"] { color: var(--text-color) !important; background-color: var(--background-color) !important; }
    li[role="option"]:hover { background-color: rgba(79,70,229,0.04) !important; }
    li[role="option"][aria-selected="true"] { background-color: rgba(79,70,229,0.1) !important; color: #4f46e5 !important; font-weight: 700 !important; }
    ul[data-testid="stMultiSelectDropdown"] {
        border: 1px solid #e2e8f0 !important; border-radius: 10px !important;
        background-color: var(--background-color) !important; box-shadow: 0 12px 40px rgba(0,0,0,0.1) !important; overflow: hidden;
    }

    /* ── Hero ──────────────────────────────────────────────────── */
    .hero { text-align: center; padding: 48px 20px 32px; }
    .h1 {
        font-family: 'Plus Jakarta Sans', sans-serif !important;
        font-size: 46px; font-weight: 800; letter-spacing: -0.03em;
        line-height: 1.05; margin-bottom: 14px; text-transform: uppercase;
        background: linear-gradient(135deg, #4f46e5, #818cf8);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent; display: inline-block;
    }
    .sub { font-size: 15px; color: #94a3b8 !important; max-width: 520px; margin: 0 auto; line-height: 1.6; font-weight: 400; }

    /* ── Cards ─────────────────────────────────────────────────── */
    .unified-card {
        background: #ffffff !important; border: 1px solid #f1f5f9 !important;
        border-radius: 14px !important; padding: 20px 24px !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04) !important;
        transition: all 0.2s !important; margin-bottom: 14px;
    }
    .unified-card:hover {
        border-color: rgba(79,70,229,0.18) !important;
        box-shadow: 0 4px 16px rgba(79,70,229,0.06) !important;
        transform: translateY(-1px) !important;
    }
    .c-eye { font-family: 'Plus Jakarta Sans', sans-serif !important; font-size: 10px; font-weight: 700; letter-spacing: 0.1em; text-transform: uppercase; color: #4f46e5; margin-bottom: 6px; }
    .c-ttl { font-family: 'Plus Jakarta Sans', sans-serif !important; font-size: 15px; font-weight: 700; color: var(--text-color); margin-bottom: 16px; letter-spacing: -0.01em; border-left: 2.5px solid #4f46e5; padding-left: 10px; line-height: 1.3; }
    .divider { height: 1px; margin: 16px 0 32px; background: linear-gradient(90deg, transparent, rgba(79,70,229,0.08), transparent); }

    /* ── Output Cards ──────────────────────────────────────────── */
    .output-card {
        background: #ffffff !important; border: 1px solid #f1f5f9 !important;
        border-radius: 16px; padding: 36px 32px; margin-top: 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05) !important; line-height: 1.8;
    }
    .output-card h3 { font-family: 'Plus Jakarta Sans', sans-serif !important; font-size: 15px; font-weight: 700; margin: 28px 0 14px; color: var(--text-color); border-bottom: 1px solid #f1f5f9; padding-bottom: 8px; letter-spacing: -0.01em; text-transform: uppercase; }
    .output-card p, .output-card li { font-size: 14px; color: #64748b; margin-bottom: 6px; }

    /* ── Details/Summary ───────────────────────────────────────── */
    details { background: #ffffff !important; border: 1px solid #f1f5f9 !important; border-radius: 10px !important; margin-top: 12px !important; overflow: hidden !important; transition: all 0.25s !important; }
    details[open] { border-color: rgba(79,70,229,0.22) !important; box-shadow: 0 4px 16px rgba(79,70,229,0.06) !important; }
    summary { padding: 13px 18px !important; cursor: pointer !important; font-weight: 700 !important; font-size: 12px !important; color: #4f46e5 !important; background: rgba(79,70,229,0.03) !important; display: flex !important; align-items: center !important; gap: 8px !important; outline: none !important; list-style: none !important; user-select: none !important; transition: background 0.15s !important; letter-spacing: 0.04em !important; text-transform: uppercase !important; }
    summary:hover { background: rgba(79,70,229,0.06) !important; }
    summary::-webkit-details-marker { display: none; }
    details > *:not(summary) { padding: 14px 18px 18px !important; animation: slideDown 0.2s ease-out; }
    @keyframes slideDown { from { opacity:0; transform:translateY(-6px); } to { opacity:1; transform:translateY(0); } }

    /* ── Badges ────────────────────────────────────────────────── */
    .badge { display: inline-block; padding: 3px 9px; border-radius: 5px; font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em; font-family: 'JetBrains Mono', monospace !important; }
    .badge-fail     { background: #fef2f2 !important; color: #dc2626 !important; border: 1px solid #fca5a5 !important; }
    .badge-review   { background: #fffbeb !important; color: #d97706 !important; border: 1px solid #fcd34d !important; }
    .badge-verified { background: #f0fdf4 !important; color: #16a34a !important; border: 1px solid #86efac !important; }
    .section-accent { border-left: 3px solid transparent !important; padding: 16px 20px !important; margin: 20px 0 !important; border-radius: 0 8px 8px 0 !important; }
    .accent-fail     { border-left-color: #dc2626 !important; background: rgba(220,38,38,0.02) !important; }
    .accent-review   { border-left-color: #f59e0b !important; background: rgba(245,158,11,0.02) !important; }
    .accent-verified { border-left-color: #22c55e !important; background: rgba(34,197,94,0.02) !important; }

    /* ── Output Section Tables ─────────────────────────────────── */
    .output-section { background: #f8fafc !important; border: 1px solid #f1f5f9 !important; border-radius: 12px !important; padding: 22px !important; margin-bottom: 10px !important; line-height: 1.8; }
    .output-section h3 { background: rgba(79,70,229,0.04) !important; border-left: 3px solid #4f46e5 !important; color: var(--text-color) !important; font-family: 'Plus Jakarta Sans', sans-serif !important; font-size: 12px !important; font-weight: 700 !important; letter-spacing: 0.05em !important; text-transform: uppercase !important; padding: 7px 14px !important; margin: 24px 0 14px 0 !important; border-bottom: none !important; border-radius: 0 4px 4px 0 !important; display: block !important; }
    .output-section h3:first-child { margin-top: 0 !important; }
    .output-section p, .output-section li { color: #64748b; margin: 6px 0; font-size: 14px; }
    .output-section table { width: 100% !important; border-collapse: collapse !important; border: 1px solid #e2e8f0 !important; border-radius: 8px !important; overflow: hidden !important; margin: 14px 0 !important; background: #ffffff !important; }
    .output-section th { background: #f8fafc !important; color: #94a3b8 !important; font-weight: 700; text-transform: uppercase; letter-spacing: 0.07em; font-size: 10px; padding: 11px 14px; border-bottom: 1px solid #e2e8f0 !important; }
    .output-section td { padding: 11px 14px; border-bottom: 1px solid #f1f5f9 !important; vertical-align: top; color: var(--text-color); line-height: 1.6; }
    .output-section tr:last-child td { border-bottom: none; }
    .output-section tr:nth-child(even) td { background: #fafafa; }
    .output-section code { background: rgba(220,38,38,0.05) !important; color: #dc2626 !important; padding: 2px 6px; border-radius: 4px; font-family: 'JetBrains Mono', monospace; font-size: 12.5px; border: 1px solid rgba(220,38,38,0.12) !important; }

    /* ── Score Card ────────────────────────────────────────────── */
    .score-card { background: #ffffff !important; border: 1px solid #f1f5f9 !important; border-top: 3px solid #4f46e5 !important; border-radius: 12px !important; padding: 18px 22px !important; margin-bottom: 14px !important; box-shadow: 0 1px 4px rgba(0,0,0,0.04) !important; }
    .score-number { font-family: 'JetBrains Mono', monospace !important; font-size: 38px; font-weight: 700; color: #4f46e5; letter-spacing: -0.03em; line-height: 1; }
    .score-label  { font-family: 'Plus Jakarta Sans', sans-serif !important; font-size: 10px; font-weight: 700; letter-spacing: 0.1em; text-transform: uppercase; color: #94a3b8; margin-bottom: 4px; }
    .audit-done-banner { display: flex; align-items: center; gap: 10px; background: #f0fdf4 !important; border: 1px solid #86efac !important; border-left: 4px solid #22c55e !important; border-radius: 8px; padding: 11px 16px; color: #16a34a; font-size: 13px; font-weight: 600; margin-bottom: 18px; }
    .audit-done-dot { width: 7px; height: 7px; border-radius: 50%; background: #22c55e; flex-shrink: 0; }

    /* ── Code Block ────────────────────────────────────────────── */
    [data-testid="stCodeBlock"] { background-color: #f8fafc !important; border-radius: 8px !important; border: 1px solid #e2e8f0 !important; margin: 0 !important; }
    [data-testid="stCodeBlock"] code { color: var(--text-color) !important; font-family: 'JetBrains Mono', monospace !important; font-size: 12.5px !important; line-height: 1.7 !important; }

    /* ── Processing Modal ──────────────────────────────────────── */
    .fixed-overlay {
        position: fixed; top: 0; left: 0; width: 100vw; height: 100dvh;
        background: rgba(15,23,42,0.55); backdrop-filter: blur(8px);
        -webkit-backdrop-filter: blur(8px); z-index: 999998;
        animation: overlayFadeIn 0.3s ease-out forwards;
    }
    .fixed-modal {
        position: fixed; top: 50%; left: 50%;
        transform: translate(-50%, -50%);
        border-radius: 16px; padding: 44px; width: 440px; max-width: 90vw;
        box-shadow: 0 24px 48px rgba(0,0,0,0.18) !important;
        text-align: center; z-index: 999999;
        background: #ffffff !important; border: 1px solid #e2e8f0 !important;
        color: var(--text-color) !important; opacity: 0;
        animation: modalSlideIn 0.35s cubic-bezier(0.34,1.56,0.64,1) forwards;
    }
    .fixed-modal h3 { font-family: 'Plus Jakarta Sans', sans-serif !important; color: #0f172a !important; margin: 0 0 10px; font-weight: 800; font-size: 18px; letter-spacing: -0.02em; }
    .fixed-modal p  { color: #94a3b8 !important; margin: 0; font-size: 13px; }
    .spinner-loader { border: 3px solid #e0e7ff; border-top: 3px solid #4f46e5; border-radius: 50%; width: 44px; height: 44px; animation: spin 0.85s linear infinite; margin: 0 auto 22px auto; }
    .progress-container { width: 100%; height: 5px; background: #e0e7ff; border-radius: 99px; margin: 22px 0 10px; overflow: hidden; }
    .progress-fill { height: 100%; background: linear-gradient(90deg, #4f46e5, #818cf8); transition: width 0.4s cubic-bezier(0.1,0.8,0.1,1); border-radius: 99px; }
    .perc-text { font-family: 'JetBrains Mono', monospace !important; font-size: 30px; font-weight: 600; color: #4f46e5; margin-bottom: 4px; }

    .cancel-btn-container { position: fixed !important; top: calc(50% + 148px) !important; left: 50% !important; transform: translateX(-50%) !important; z-index: 1000000 !important; width: 200px !important; text-align: center !important; }
    .cancel-btn-container button { background: #ffffff !important; color: #dc2626 !important; border: 1.5px solid #fee2e2 !important; border-radius: 9px !important; font-size: 12px !important; font-weight: 600 !important; padding: 9px 18px !important; transition: all 0.15s !important; width: 100% !important; font-family: 'Plus Jakarta Sans', sans-serif !important; }
    .cancel-btn-container button:hover { background: #fef2f2 !important; border-color: #fca5a5 !important; }

    /* ── Animations ────────────────────────────────────────────── */
    @keyframes fadeUp { from { opacity:0; transform:translateY(10px); } to { opacity:1; transform:translateY(0); } }
    @keyframes spin { to { transform: rotate(360deg); } }
    @keyframes overlayFadeIn { from { opacity:0; } to { opacity:1; } }
    @keyframes modalSlideIn { 0% { opacity:0; transform:translate(-50%,-50%) scale(0.94); } 100% { opacity:1; transform:translate(-50%,-50%) scale(1); } }
    .first-run-anim { opacity:0; animation: fadeUp 0.4s cubic-bezier(0.2,0.8,0.2,1) forwards; }
    .anim-delay-1 { animation-delay: 0.08s; }
    .anim-delay-2 { animation-delay: 0.18s; }
    .anim-delay-3 { animation-delay: 0.28s; }

    /* ── App Selector Sidebar Header ───────────────────────────── */
    .app-selector-header-container { background: #f8fafc !important; border: 1px solid #f1f5f9 !important; border-radius: 12px; padding: 11px 14px; margin-bottom: 20px; transition: all 0.15s !important; cursor: pointer; position: relative; z-index: 5; }
    .app-selector-header-container:hover { background: #f1f5f9 !important; border-color: #e2e8f0 !important; }
    .app-selector-header { display: flex; align-items: center; gap: 10px; width: 100%; }
    .app-selector-header .logo-circle { width: 34px; height: 34px; border-radius: 9px; display: flex; align-items: center; justify-content: center; font-family: 'Plus Jakarta Sans', sans-serif !important; font-weight: 800; font-size: 12px; color: #ffffff !important; flex-shrink: 0; }
    .da-logo { background: linear-gradient(135deg, #4f46e5, #818cf8) !important; }
    .cc-logo { background: linear-gradient(135deg, #7c3aed, #818cf8) !important; }
    .app-selector-header .text-block { flex-grow: 1; text-align: left; line-height: 1.2; }
    .app-selector-header .main-title { font-family: 'Plus Jakarta Sans', sans-serif !important; font-weight: 700; font-size: 13px; letter-spacing: -0.01em; color: var(--text-color) !important; }
    .app-selector-header .sub-title { font-size: 10px; color: #94a3b8 !important; font-weight: 500; margin-top: 1px; }
    .app-selector-header .chevron { font-size: 10px; color: #cbd5e1 !important; flex-shrink: 0; }

    /* ── Toggle Button Overlay ─────────────────────────────────── */
    div[data-testid="stVerticalBlock"]:has(.toggle-btn-wrapper) { position: relative; margin-top: -72px; height: 56px; margin-bottom: 20px; z-index: 10; }
    div[data-testid="stVerticalBlock"]:has(.toggle-btn-wrapper) div.stButton button { background: transparent !important; border: none !important; opacity: 0 !important; width: 100% !important; height: 56px !important; cursor: pointer !important; }

    /* ── Popup Card ────────────────────────────────────────────── */
    .app-popup-card { position: fixed; top: 90px; left: 16px; width: 306px; background: #ffffff !important; border: 1px solid #e2e8f0 !important; border-radius: 14px; padding: 14px; box-shadow: 0 20px 48px rgba(0,0,0,0.12) !important; z-index: 999995; opacity: 0; animation: popupSlideIn 0.3s cubic-bezier(0.34,1.56,0.64,1) forwards; }
    @keyframes popupSlideIn { 0% { opacity:0; transform:translateY(-10px) scale(0.96); } 100% { opacity:1; transform:translateY(0) scale(1); } }
    .popup-header { display: flex; align-items: center; gap: 9px; padding: 3px 5px 10px 5px; }
    .popup-header .building-icon { font-size: 15px; color: #4f46e5 !important; }
    .popup-header .header-text { line-height: 1.2; text-align: left; }
    .popup-header .title { font-family: 'Plus Jakarta Sans', sans-serif !important; font-size: 12px; font-weight: 800; color: var(--text-color) !important; letter-spacing: 0.04em; text-transform: uppercase; }
    .popup-header .subtitle { font-size: 10.5px; color: #94a3b8 !important; }
    .divider-line { height: 1px; background: #f1f5f9 !important; margin: 0 0 8px 0; }
    .app-item-card { display: flex; align-items: center; gap: 11px; padding: 11px; border-radius: 10px; background: transparent; border: 1px solid transparent; transition: all 0.15s !important; margin-bottom: 6px; position: relative; }
    .app-item-card:hover { background: #f8fafc !important; border-color: #f1f5f9 !important; }
    .app-item-card.active-item { background: rgba(79,70,229,0.06) !important; border-color: rgba(79,70,229,0.2) !important; }
    .app-item-card .item-text { flex-grow: 1; text-align: left; line-height: 1.3; }
    .app-item-card .item-title { font-family: 'Plus Jakarta Sans', sans-serif !important; font-weight: 700; font-size: 12px; color: #1e293b !important; }
    .app-item-card.active-item .item-title { color: #4f46e5 !important; }
    .app-item-card .item-subtitle { font-size: 10.5px; color: #94a3b8 !important; }
    .app-item-card .active-badge { background: #4f46e5 !important; color: #ffffff !important; font-size: 8.5px !important; font-weight: 700 !important; padding: 3px 8px !important; border-radius: 20px !important; text-transform: uppercase !important; letter-spacing: 0.05em !important; flex-shrink: 0; }

    /* ── Popup Click Overlays ──────────────────────────────────── */
    div[data-testid="stVerticalBlock"]:has(.popup-da-btn-wrapper) { position: fixed; top: 148px; left: 28px; width: 278px; height: 58px; z-index: 999997; }
    div[data-testid="stVerticalBlock"]:has(.popup-da-btn-wrapper) div.stButton button { opacity: 0 !important; background: transparent !important; border: none !important; height: 58px !important; cursor: pointer !important; }
    div[data-testid="stVerticalBlock"]:has(.popup-cc-btn-wrapper) { position: fixed; top: 214px; left: 28px; width: 278px; height: 58px; z-index: 999997; }
    div[data-testid="stVerticalBlock"]:has(.popup-cc-btn-wrapper) div.stButton button { opacity: 0 !important; background: transparent !important; border: none !important; height: 58px !important; cursor: pointer !important; }
    div[data-testid="stVerticalBlock"]:has(.close-overlay-wrapper) { position: fixed; top: 0; left: 0; width: 100vw; height: 100dvh; z-index: 999990; }
    div[data-testid="stVerticalBlock"]:has(.close-overlay-wrapper) div.stButton button { opacity: 0 !important; background: transparent !important; border: none !important; width: 100vw !important; height: 100dvh !important; cursor: pointer !important; }
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
    import html
    text = html.unescape(text)
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
    # ─── Sidebar Selector Markup and Overlays ───
    active_app = st.session_state.active_app
    if active_app == "DATA AUDITOR":
        active_logo = "DA"
        logo_class = "da-logo"
        active_title = "DATA AUDITOR"
        active_subtitle = "EXCEL AUDITING PORTAL"
    else:
        active_logo = "CC"
        logo_class = "cc-logo"
        active_title = "CONTRACT COMPARE"
        active_subtitle = "CONTRACT COMPARISON"

    # 1. Overlay Backdrop wrapper if open
    if st.session_state.get("app_selector_open", False):
        with st.container():
            st.markdown('<div class="close-overlay-wrapper"></div>', unsafe_allow_html=True)
            if st.button("Close Selector", key="close_selector_overlay", use_container_width=True):
                st.session_state.app_selector_open = False
                st.rerun()

    # 2. Launcher Header UI
    chevron_symbol = "▼"
    st.markdown(f"""
    <div class="app-selector-header-container">
        <div class="app-selector-header">
            <div class="logo-circle {logo_class}">{active_logo}</div>
            <div class="text-block">
                <div class="main-title">{active_title}</div>
                <div class="sub-title">{active_subtitle}</div>
            </div>
            <div class="chevron">{chevron_symbol}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # 3. Transparent Overlay Button for Launcher Header
    with st.container():
        st.markdown('<div class="toggle-btn-wrapper"></div>', unsafe_allow_html=True)
        if st.button("Toggle Selector Menu", key="toggle_selector_btn", use_container_width=True):
            st.session_state.app_selector_open = not st.session_state.app_selector_open
            st.rerun()

    # 4. Floating Popup Card Menu
    if st.session_state.get("app_selector_open", False):
        da_active = "active-item" if active_app == "DATA AUDITOR" else ""
        cc_active = "active-item" if active_app == "CONTRACT COMPARE" else ""
        da_badge = '<span class="active-badge">Active</span>' if active_app == "DATA AUDITOR" else ""
        cc_badge = '<span class="active-badge">Active</span>' if active_app == "CONTRACT COMPARE" else ""
        
        html_str = f"""<div class="app-popup-card">
<div class="popup-header">
<div class="building-icon">
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="color: #3b82f6; vertical-align: middle;"><rect x="4" y="2" width="16" height="20" rx="2" ry="2"></rect><line x1="9" y1="22" x2="9" y2="16"></line><line x1="15" y1="22" x2="15" y2="16"></line><line x1="9" y1="16" x2="15" y2="16"></line><path d="M8 6h.01"></path><path d="M16 6h.01"></path><path d="M12 6h.01"></path><path d="M12 10h.01"></path><path d="M8 10h.01"></path><path d="M16 10h.01"></path><path d="M8 14h.01"></path><path d="M16 14h.01"></path></svg>
</div>
<div class="header-text">
<div class="title">APPLICATIONS</div>
<div class="subtitle">Select active auditor application</div>
</div>
</div>
<div class="divider-line"></div>
<!-- DATA AUDITOR ITEM -->
<div class="app-item-card {da_active}">
<div class="logo-circle da-logo">DA</div>
<div class="item-text">
<div class="item-title">DATA AUDITOR</div>
<div class="item-subtitle">Excel Rate Auditing</div>
</div>
{da_badge}
</div>
<!-- CONTRACT COMPARE ITEM -->
<div class="app-item-card {cc_active}" style="margin-top: 8px;">
<div class="logo-circle cc-logo">CC</div>
<div class="item-text">
<div class="item-title">CONTRACT COMPARE</div>
<div class="item-subtitle">PDF Comparison Engine</div>
</div>
{cc_badge}
</div>
</div>"""
        st.markdown(html_str, unsafe_allow_html=True)

        # 5. Floating Popup Card Click Overlays
        with st.container():
            st.markdown('<div class="popup-da-btn-wrapper"></div>', unsafe_allow_html=True)
            if st.button("Select Data Auditor App", key="select_da_app", use_container_width=True):
                st.session_state.active_app = "DATA AUDITOR"
                st.session_state.app_selector_open = False
                st.rerun()

        with st.container():
            st.markdown('<div class="popup-cc-btn-wrapper"></div>', unsafe_allow_html=True)
            if st.button("Select Contract Compare App", key="select_cc_app", use_container_width=True):
                st.session_state.active_app = "CONTRACT COMPARE"
                st.session_state.app_selector_open = False
                st.rerun()

    # ─── Navigation Button Page Routing ───
    st.markdown('<div class="nav-label">NAVIGATION</div>', unsafe_allow_html=True)
    
    if "selected_page" not in st.session_state:
        st.session_state.selected_page = "CONTRACT AUDITOR"

    if active_app == "DATA AUDITOR":
        page_options = ["CONTRACT AUDITOR", "AI EXCEL GENERATOR"]
        if st.session_state.selected_page == "CONTRACT COMPARE":
            st.session_state.selected_page = "CONTRACT AUDITOR"
    else:
        page_options = ["CONTRACT COMPARE"]
        if st.session_state.selected_page not in page_options:
            st.session_state.selected_page = "CONTRACT COMPARE"
        
    for opt in page_options:
        btn_class = "nav-btn-active" if st.session_state.selected_page == opt else "nav-btn-wrap"
        
        icon = "📄"
        if opt == "AI EXCEL GENERATOR": icon = "⚡"
        elif opt == "CONTRACT COMPARE": icon = "🔄"
            
        st.markdown(f'<div class="{btn_class}">', unsafe_allow_html=True)
        if st.button(f"{icon}  {opt}", key=f"nav_btn_{opt}", use_container_width=True):
            st.session_state.selected_page = opt
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
        
    selected_page = st.session_state.selected_page
    
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

    # ─── Recent Comparison History (Sidebar) ───
    if active_app == "CONTRACT COMPARE":
        st.markdown("<div style='font-size: 10px; font-weight: 700; letter-spacing: 0.1em; text-transform: uppercase; opacity: 0.4; margin-bottom: 8px; margin-top: 32px;'>RECENT AUDITS</div>", unsafe_allow_html=True)
        os.makedirs("history", exist_ok=True)
        history_files = sorted([f for f in os.listdir("history") if f.endswith(".xlsx")], reverse=True)
        if not history_files:
            st.markdown("<div style='font-size: 11px; opacity: 0.5; padding-left: 4px;'>No recent audits found.</div>", unsafe_allow_html=True)
        else:
            for hf in history_files[:8]:
                if hf.startswith("Comparison_"):
                    display_name = hf.split('_vs_')[0].replace('Comparison_', '')
                    if len(display_name) > 22:
                        display_name = display_name[:20] + "..."
                else:
                    parts = hf.rsplit("_", 2)
                    display_name = parts[0][:22] if len(parts) >= 3 else hf[:22]
                
                try:
                    with open(os.path.join("history", hf), "rb") as f:
                        st.download_button(
                            label=f"{display_name.upper()}",
                            data=f.read(),
                            file_name=hf,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="history_" + hf,
                            use_container_width=True
                        )
                except Exception:
                    pass



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
elif selected_page == "AI EXCEL GENERATOR":
    st.markdown(f"""
    <div class="hero {anim_class}">
      <div class="h1">EXCEL GENERATOR</div>
      <div class="sub">Generate Hotel Upload Files from PDF.</div>
    </div>
    <div class="divider"></div>
    """, unsafe_allow_html=True)
elif selected_page == "CONTRACT COMPARE":
    st.markdown(f"""
    <div class="hero {anim_class}">
      <div class="h1">CONTRACT COMPARE</div>
      <div class="sub">AI-Powered Contract Comparison & Revision Verification.</div>
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
        st.markdown(f"<p style='text-align:center;color:rgba(130, 130, 130, 0.5);font-size:12px;margin-top:6px;letter-spacing:0.02em'>{hint}</p>", unsafe_allow_html=True)
        
    st.stop()

elif selected_page == "CONTRACT COMPARE":
    # ─── Interrupt Check for Cancellation ───
    if "cc_cancel_requested" not in st.session_state:
        st.session_state.cc_cancel_requested = False

    if st.session_state.get("cc_cancel_btn_trigger") or st.session_state.get("cc_cancel_requested"):
        st.session_state.cc_started = False
        st.session_state.cc_cancel_requested = False
        st.session_state.pop("cc_cancel_btn_trigger", None)
        st.rerun()

    is_focus_mode = st.session_state.get("cc_review_mode") or st.session_state.get("cc_report_ready")

    up1 = None
    up2 = None

    if not is_focus_mode:
        st.markdown("<div style='font-weight:700;font-size:12px;letter-spacing:0.1em;text-transform:uppercase;color:var(--text-color);opacity:0.5;margin-bottom:16px;'>Upload Contracts</div>", unsafe_allow_html=True)
        c1, c2 = st.columns(2, gap="large")
        with c1:
            st.markdown('<div class="unified-card"><div class="c-eye">STEP 1</div><div class="c-ttl">Previous Contract</div>', unsafe_allow_html=True)
            up1 = st.file_uploader("Contract 1", type=["pdf"], key="pdf1", label_visibility="collapsed")
            if up1:
                st.markdown(f'<div style="font-size:12px;color:#10b981;font-weight:600;margin-top:8px;">{up1.name}</div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        
        with c2:
            st.markdown('<div class="unified-card"><div class="c-eye">STEP 2</div><div class="c-ttl">New Contract</div>', unsafe_allow_html=True)
            up2 = st.file_uploader("Contract 2", type=["pdf"], key="pdf2", label_visibility="collapsed")
            if up2:
                st.markdown(f'<div style="font-size:12px;color:#10b981;font-weight:600;margin-top:8px;">{up2.name}</div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

    cta_placeholder = st.empty()

    # ─── CTA ──────────────────────────────────────────────────────────────────
    ready = bool(up1 and up2 and api_key)

    if not is_focus_mode:
        up1_name = up1.name if up1 else ""
        up2_name = up2.name if up2 else ""
        if "cc_last_up1" not in st.session_state or st.session_state.cc_last_up1 != up1_name or st.session_state.cc_last_up2 != up2_name:
            st.session_state.cc_started = False
            st.session_state.cc_review_mode = False
            st.session_state.cc_report_ready = False
            st.session_state.cc_extracted_data = None
            st.session_state.cc_last_up1 = up1_name
            st.session_state.cc_last_up2 = up2_name

        with cta_placeholder.container():
            _, btn_col, _ = st.columns([1.5, 3, 1.5])
            with btn_col:
                if st.button("Compare Contracts  →", type="primary", use_container_width=True, disabled=not ready, key="start_compare_btn"):
                    st.session_state.cc_started = True
                    st.rerun()
                
                if not ready:
                    hint = "Upload both contracts to continue" if not (up1 and up2) else "Add API Key in sidebar"
                    st.markdown(f"<p style='text-align:center;color:#9ca3af;font-size:13px;margin-top:6px'>{hint}</p>",
                                unsafe_allow_html=True)

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    else:
        # FOCUS MODE is active: Show a Back/Reset button at the top instead of the uploaders
        _, reset_col, _ = st.columns([1.5, 3, 1.5])
        with reset_col:
            if st.button("← Upload Different Contracts", use_container_width=True, key="reset_compare_btn"):
                st.session_state.cc_started = False
                st.session_state.cc_review_mode = False
                st.session_state.cc_report_ready = False
                st.session_state.cc_extracted_data = None
                st.rerun()
        st.markdown("<br>", unsafe_allow_html=True)

    # ─── Processing ───────────────────────────────────────────────────────────
    if st.session_state.get("cc_started"):
        placeholder = st.empty()
        
        def render_modal(pct):
            placeholder.markdown(f"""
                <div class="fixed-overlay"></div>
                <div class="fixed-modal">
                    <div class="spinner-loader" style="margin-bottom:20px;"></div>
                    <h3 style="margin:0 0 8px; font-weight:700;">Analyzing Contracts...</h3>
                    <p style="margin:0; font-size:14px; opacity:0.8;">Extracting policy rules and prices. Please wait.</p>
                    <div style="margin-top:24px; background:var(--secondary-background-color); border-radius:10px; height:6px; overflow:hidden;">
                        <div style="background: linear-gradient(90deg, #3b82f6, #8b5cf6); width: {pct}%; height: 100%; transition: width 0.3s ease;"></div>
                    </div>
                    <div style="text-align:right; font-size:12px; margin-top:6px; font-weight:600; color:#3b82f6;">{pct}%</div>
                </div>
            """, unsafe_allow_html=True)
            
        render_modal(10)
        
        if not up1 or not up2:
            placeholder.empty()
            st.error("Upload files not found. Please upload both contracts and try again.")
            st.session_state.cc_started = False
            st.stop()
            
        if compare_utils is None:
            placeholder.empty()
            st.error("The CONTRACT COMPARE module (`utils2.py`) was not found in your cloud deployment. Please ensure you uploaded all files to your GitHub repository.")
            st.session_state.cc_started = False
            st.stop()
        
        pdf1_bytes = up1.getvalue()
        pdf2_bytes = up2.getvalue()
        
        chunks = []
        char_count = 0
        EXPECTED_CHARS = 6000 

        cancel_compare_placeholder = st.empty()
        cancel_compare_placeholder.markdown('<div class="cancel-btn-container">', unsafe_allow_html=True)
        if cancel_compare_placeholder.button("CANCEL COMPARISON", key="cc_cancel_btn_trigger"):
            st.session_state.cc_cancel_requested = True
            st.rerun()

        try:
            for chunk in compare_utils.stream_contract_comparison(pdf1_bytes, pdf2_bytes, api_key):
                if st.session_state.get("cc_cancel_requested"):
                    break

                if chunk == "[RESET_STREAM]":
                    chunks = []
                    char_count = 0
                    render_modal(10)
                    continue
                    
                chunks.append(chunk)
                char_count += len(chunk)
                pct = min(15 + int(char_count / EXPECTED_CHARS * 80), 98)
                render_modal(pct)

            cancel_compare_placeholder.empty()
            
            if st.session_state.get("cc_cancel_requested"):
                st.session_state.cc_started = False
                st.session_state.cc_cancel_requested = False
                placeholder.empty()
                st.rerun()
                
            render_modal(100)
            placeholder.empty()
            
            result_raw = "".join(chunks)
            st.session_state.cc_started = False

            if "429" in result_raw or "quota" in result_raw.lower() or "quota exceeded" in result_raw.lower():
                st.error("**API Quota Exceeded (429)**")
                st.warning(
                    "Free-tier limit reached (20 req/day).\n\n"
                    "**Fix:** Enable Billing at aistudio.google.com "
                    "to increase quota."
                )
                st.stop()

            cleaned = _clean_json(result_raw)
            try:
                data = json.loads(cleaned)
            except json.JSONDecodeError:
                try:
                    repaired = _repair_json(cleaned)
                    data = json.loads(repaired)
                except json.JSONDecodeError as ex:
                    st.error(f"**JSON Parse Error:** {ex}")
                    st.info("The AI response was slightly malformed.")
                    st.stop()

            if "error" in data:
                err = data["error"]
                st.error("Quota Exceeded" if ("429" in err or "quota" in err.lower()) else f"AI Error: {err}")
                st.stop()
                
            st.session_state.cc_extracted_data = data
            st.session_state.cc_started = False
            st.session_state.cc_review_mode = True
            st.rerun()
        except Exception as e:
            placeholder.empty()
            cancel_compare_placeholder.empty()
            st.error(f"Comparison failed: {str(e)}")
            st.session_state.cc_started = False

    # ─── Review & Edit Mode ───────────────────────────────────────────────────
    if st.session_state.get("cc_review_mode"):
        st.markdown("""
            <div style="display:flex; align-items:center; gap:12px; margin-bottom: 20px;">
                <div style="background:linear-gradient(135deg, #3b82f6, #8b5cf6); border-radius:6px; padding:6px 12px; color:white; font-weight:700; font-size:12px; letter-spacing:1px; box-shadow:0 4px 6px -1px rgba(59, 130, 246, 0.3);">DATA VERIFICATION</div>
                <div style="font-size:24px; font-weight:700; color:var(--text-color); letter-spacing:-0.03em;">Review & Edit Prices</div>
            </div>
            <div style="background:var(--secondary-background-color); border-left: 4px solid #3b82f6; border-radius:4px 8px 8px 4px; padding:16px 20px; margin-bottom: 32px;">
                <p style="margin:0; font-size:14px; color:var(--text-color); opacity:0.9; line-height:1.6;">
                    AI extraction is complete. Please verify the extracted prices below. You can <b>click any cell to edit</b> the value before finalizing the Excel report.
                </p>
            </div>
        """, unsafe_allow_html=True)
        
        edited_data = copy.deepcopy(st.session_state.cc_extracted_data)
        
        seasons_data = edited_data.get("seasons") or []
        for i, season in enumerate(seasons_data):
            s_name = season.get("season_name") or f"Season {i+1}"
            p1 = season.get("period_1", "")
            p2 = season.get("period_2", "")
            
            p1_display = p1 if p1 and p1.strip() and p1 != "N/A" else "Not Specified"
            p2_display = p2 if p2 and p2.strip() and p2 != "N/A" else "Not Specified"
            
            st.markdown(f"""
                <div style="margin-top:32px; margin-bottom:16px; display:flex; align-items:baseline; flex-wrap:wrap; gap:12px; border-bottom:2px solid var(--secondary-background-color); padding-bottom:8px;">
                    <div style="font-size:17px; font-weight:700; color:var(--text-color);">{s_name}</div>
                    <div style="font-size:13px; color:var(--text-color);">
                        <span style="opacity:0.6;">Prev:</span> <span style="font-weight:600; opacity:0.9;">{p1_display}</span> 
                        <span style="margin:0 8px;opacity:0.2;">|</span> 
                        <span style="opacity:0.6;">New:</span> <span style="font-weight:600; color:#3b82f6;">{p2_display}</span>
                    </div>
                </div>
            """, unsafe_allow_html=True)
            
            rooms = season.get("rooms", [])
            if rooms:
                edited_rooms = st.data_editor(
                    rooms,
                    column_config={
                        "room_name": st.column_config.TextColumn("Room Name", width="large"),
                        "price_1": st.column_config.TextColumn("Contract 1 Price"),
                        "price_2": st.column_config.TextColumn("Contract 2 Price"),
                    },
                    hide_index=True,
                    key=f"cc_editor_season_{i}",
                    use_container_width=True
                )
                edited_data["seasons"][i]["rooms"] = edited_rooms
                
        st.markdown("<br>", unsafe_allow_html=True)
        _, btn1, btn2, _ = st.columns([1, 1.5, 1.5, 1])
        with btn1:
            if st.button("Cancel & Start Over", use_container_width=True, key="cancel_review_btn"):
                st.session_state.cc_started = False
                st.session_state.cc_review_mode = False
                st.session_state.cc_extracted_data = None
                st.rerun()
        with btn2:
            if st.button("Confirm & Generate Excel", type="primary", use_container_width=True, key="confirm_excel_btn"):
                st.session_state.cc_final_data = edited_data
                st.session_state.cc_review_mode = False
                st.session_state.cc_report_ready = True
                st.rerun()

    # ─── Generate Excel & Downloader ──────────────────────────────────────────
    if st.session_state.get("cc_report_ready"):
        data = st.session_state.cc_final_data
        
        try:
            excel_bytes = compare_excel.generate_comparison_excel(data)
            
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            hotel_name_raw = data.get("hotel_name")
            if not hotel_name_raw or not str(hotel_name_raw).strip():
                hotel_name = "Unknown_Hotel"
            else:
                hotel_name = str(hotel_name_raw).strip()
            
            if hotel_name.upper() == "HOTEL NAME":
                hotel_name = "Unknown_Hotel"
            
            hotel_name_safe = re.sub(r'[\\/*?:"<>|]', "", hotel_name)
            history_file_name = f"{hotel_name_safe}_{timestamp}.xlsx"
            
            try:
                os.makedirs("history", exist_ok=True)
                with open(os.path.join("history", history_file_name), "wb") as f:
                    f.write(excel_bytes)
            except Exception as e:
                pass
                
        except Exception as ex:
            st.error(f"Excel generation failed: {ex}")
            st.stop()

        st.markdown("""
            <div style="background: linear-gradient(135deg, rgba(16,185,129,0.1), rgba(59,130,246,0.1));
                        border: 1px solid rgba(16,185,129,0.3); border-radius: 12px;
                        padding: 16px 24px; margin: 24px 0; display:flex; align-items:center; gap:12px;">
                <div style="width:8px;height:8px;border-radius:50%;background:#10b981;box-shadow:0 0 8px #10b981;"></div>
                <div style="font-size:14px;font-weight:600;color:var(--text-color);">Analysis complete — your comparison report is ready to download.</div>
            </div>
        """, unsafe_allow_html=True)
        
        recommendation = str(data.get("recommendation") or "").strip()
        if recommendation:
            st.markdown(
                f'<div style="background:var(--secondary-background-color);border-radius:10px;'
                f'padding:14px 20px;margin:8px 0 16px;font-size:15px;font-weight:600;color:var(--text-color)">'
                f'{recommendation}</div>',
                unsafe_allow_html=True
            )

        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                "Download Excel Report",
                data=excel_bytes,
                file_name="Contract_Comparison.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
                key="download_compare_report_btn"
            )
            st.caption("In Google Sheets → File → Import → Upload .xlsx")
            
        with col2:
            if st.button("Compare Another", use_container_width=True, key="compare_another_btn"):
                st.session_state.cc_started = False
                st.session_state.cc_review_mode = False
                st.session_state.cc_report_ready = False
                st.session_state.cc_extracted_data = None
                st.rerun()
                
    st.stop()

# ─── Upload Area ───────────────────────────────────────────────────────────────
_audit_done = st.session_state.get("audit_done", False) and bool(st.session_state.get("_audit_result"))
_is_auditing = st.session_state.get("is_auditing", False)

# Render clean action buttons in top-right when audit is done
if _audit_done:
    _left, _right1, _right2 = st.columns([4, 1.5, 1.5])
    with _right1:
        if st.button("EDIT SCOPE / RE-AUDIT", use_container_width=True, key="edit_scope_top_btn"):
            for key in ["audit_done", "is_auditing", "_audit_result"]:
                st.session_state.pop(key, None)
            st.rerun()
    with _right2:
        if st.button("START FRESH", use_container_width=True, key="start_fresh_top_btn"):
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
    st.markdown('<div style="font-size:10px;font-weight:800;letter-spacing:0.15em;text-transform:uppercase;color:rgba(130, 130, 130, 0.6);margin-bottom:20px;padding-bottom:12px;border-bottom:1px solid rgba(130, 130, 130, 0.12)">AUDIT REPORT</div>', unsafe_allow_html=True)
    
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
        st.markdown('<div style="font-size:10px;font-weight:800;letter-spacing:0.15em;text-transform:uppercase;color:rgba(130, 130, 130, 0.6);margin-bottom:20px;padding-bottom:12px;border-bottom:1px solid rgba(130, 130, 130, 0.12)">AUDIT REPORT</div>', unsafe_allow_html=True)
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
                    {"<div style='font-size:14px;color:rgba(130, 130, 130, 0.65);margin-top:12px;'>"+summary_txt+"</div>" if summary_txt else ""}
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
        if st.button("EDIT SCOPE / RE-AUDIT", use_container_width=True, key="edit_scope_btm_btn"):
            for key in ["audit_done", "is_auditing", "_audit_result"]:
                st.session_state.pop(key, None)
            st.rerun()
    with reset_btn:
        if st.button("RESET / NEW UPLOAD", use_container_width=True, key="reset_btm_btn"):
            st.session_state.confirm_reset = True
            st.rerun()