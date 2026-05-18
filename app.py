import os
import re
from datetime import datetime
import streamlit as st
from utils import stream_recheck_analysis, validate_api_key

st.set_page_config(
    page_title="Recheck Excel Data",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── CSS INJECTION ───────────────────────────────────────────────────────────
def load_css():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
    html,body,[class*="css"]{font-family:'Inter',sans-serif!important}

    /* Background & Main Structure */
    .stApp { background-color: transparent !important; }
    .block-container { padding: 2rem 2.5rem 3rem !important; }
    section[data-testid="stSidebar"] { border-right: 1px solid var(--secondary-background-color) !important; }

    /* Sidebar clean buttons */
    section[data-testid="stSidebar"] .stDownloadButton button {
        background-color: transparent !important;
        border: 1px solid transparent !important;
        color: var(--text-color) !important;
        border-radius: 8px !important;
        font-size: 13px !important;
        font-weight: 500 !important;
        justify-content: flex-start !important;
        padding: 6px 12px !important;
        transition: background-color 0.2s ease !important;
        box-shadow: none !important;
    }
    section[data-testid="stSidebar"] .stDownloadButton button:hover {
        background-color: var(--secondary-background-color) !important;
    }

    /* Hero Typography */
    .hero { text-align: center; padding: 50px 16px 32px; }
    .h1 {
        font-size: 72px; font-weight: 800; letter-spacing: -0.05em; line-height: 1; margin-bottom: 16px;
        background: linear-gradient(135deg, #10b981, #3b82f6);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        display: inline-block;
    }
    .sub { font-size: 20px; color: var(--text-color); opacity: 0.6; max-width: 600px; margin: 0 auto; line-height: 1.6; font-weight: 400; }

    /* Upload Area & Cards */
    .unified-card { background: var(--background-color); border: 1px solid var(--secondary-background-color); border-radius: 16px; padding: 24px; transition: all .3s ease; box-shadow: 0 4px 6px -1px rgba(0,0,0,.02); margin-bottom: 20px; }
    .unified-card:hover { border-color: rgba(59, 130, 246, 0.4); }

    .c-eye { 
        font-size: 10px; font-weight: 800; letter-spacing: 0.12em; text-transform: uppercase; 
        color: #10b981; margin-bottom: 8px; opacity: 0.7;
    }
    .c-ttl { 
        font-size: 20px; font-weight: 800; color: var(--text-color); margin-bottom: 24px; 
        letter-spacing: -0.03em; border-left: 4px solid #10b981; padding-left: 14px; line-height: 1.2;
    }

    div[data-testid="stFileUploader"] { width: 100% !important; }
    div[data-testid="stFileUploader"] > section {
        background: var(--secondary-background-color) !important;
        border: 1px dashed rgba(156, 163, 175, 0.4) !important;
        border-radius: 12px !important;
        padding: 16px !important;
        transition: all 0.2s ease !important;
    }
    div[data-testid="stFileUploader"] > section:hover {
        border-color: #3b82f6 !important;
        background: rgba(59, 130, 246, 0.05) !important;
    }
    div[data-testid="stFileUploader"] small { display: none !important; }

    /* Gradient Divider */
    .divider {
        height: 1px; margin: 10px 0 30px; 
        background: linear-gradient(90deg, transparent, rgba(59,130,246,0.3), rgba(147,51,234,0.3), transparent); 
        box-shadow: 0 4px 16px rgba(59,130,246,0.15), 0 1px 8px rgba(147,51,234,0.1);
    }

    /* Button */
    button[data-testid="baseButton-primary"] {
        background: linear-gradient(135deg, #0ea5e9, #3b82f6) !important; color: #fff !important; border: none !important;
        border-radius: 12px !important; font-size: 16px !important; font-weight: 600 !important; padding: 12px 24px !important;
        box-shadow: 0 4px 14px rgba(59,130,246,.3) !important; transition: all .2s ease !important;
    }
    button[data-testid="baseButton-primary"]:hover { transform: translateY(-2px); box-shadow: 0 8px 25px rgba(59,130,246,.4) !important; }
    button[data-testid="baseButton-primary"]:disabled { background: var(--secondary-background-color) !important; color: gray !important; box-shadow: none !important; transform: none; }

    /* ─── Multiselect / Dropdown Fix ─────────────────────────────────── */
    /* Selected tag pills */
    span[data-baseweb="tag"] {
        background-color: #3b82f6 !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
    }
    span[data-baseweb="tag"] span { color: #ffffff !important; }
    /* Dropdown option: default */
    li[role="option"] {
        color: #1e293b !important;
        background-color: #ffffff !important;
    }
    /* Dropdown option: hover */
    li[role="option"]:hover, li[role="option"][aria-selected="false"]:hover {
        background-color: #eff6ff !important;
        color: #1d4ed8 !important;
    }
    /* Dropdown option: selected (the dark background bug) */
    li[role="option"][aria-selected="true"] {
        background-color: #2563eb !important;
        color: #ffffff !important;
        font-weight: 700 !important;
    }
    /* Dropdown container */
    ul[data-testid="stMultiSelectDropdown"] {
        border: 1px solid #e2e8f0 !important;
        border-radius: 12px !important;
        box-shadow: 0 8px 24px rgba(0,0,0,0.12) !important;
        overflow: hidden;
    }

    /* ─── Output Card & Table ─────────────────────────────────────────── */
    .output-card {
        background: var(--background-color);
        border: 1px solid var(--secondary-background-color); 
        border-radius: 16px; padding: 36px 32px; margin-top: 20px;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,.05);
        line-height: 1.75;
    }
    .output-card h3 {
        font-size: 18px; font-weight: 800; margin: 28px 0 12px;
        color: var(--text-color);
        border-bottom: 2px solid #10b98122; padding-bottom: 8px;
    }
    .output-card p, .output-card li {
        font-size: 15px; color: var(--text-color); margin-bottom: 6px;
    }
    .output-card table {
        width: 100%; border-collapse: collapse; margin: 16px 0;
        font-size: 14px; border-radius: 12px; overflow: hidden;
        border: 1px solid #e2e8f0;
    }
    .output-card th {
        background: #f1f5f9; color: #475569;
        font-weight: 700; text-transform: uppercase;
        letter-spacing: 0.05em; font-size: 11px;
        text-align: left; padding: 12px 16px;
        border-bottom: 2px solid #e2e8f0;
    }
    .output-card td {
        padding: 12px 16px; border-bottom: 1px solid #f1f5f9;
        vertical-align: top; color: var(--text-color);
    }
    .output-card tr:last-child td { border-bottom: none; }
    .output-card tr:hover td { background: #f8fafc; }
    .output-card code {
        background: #f1f5f9; color: #dc2626;
        padding: 2px 8px; border-radius: 4px;
        font-family: 'JetBrains Mono', monospace; font-size: 13px;
    }
    .output-card hr {
        border: none; border-top: 2px solid #e2e8f0;
        margin: 24px 0;
    }

    /* ─── Details / Summary Dropdown ─────────────────────────────────── */
    details {
        background: #f8fafc;
        border: 1.5px solid #3b82f6;
        border-radius: 10px;
        margin-top: 14px;
        overflow: hidden;
        transition: box-shadow 0.3s ease;
    }
    details[open] {
        box-shadow: 0 4px 20px rgba(59, 130, 246, 0.15);
    }
    summary {
        padding: 12px 18px;
        cursor: pointer;
        font-weight: 700;
        font-size: 14px;
        color: #1d4ed8;
        background: #eff6ff;
        display: flex;
        align-items: center;
        gap: 8px;
        outline: none;
        list-style: none;
        user-select: none;
        transition: background 0.2s ease;
    }
    summary:hover { background: #dbeafe; }
    summary::-webkit-details-marker { display: none; }
    details > *:not(summary) {
        padding: 0 16px 16px;
        animation: slideDown 0.25s ease-out;
    }
    @keyframes slideDown {
        from { opacity: 0; transform: translateY(-8px); }
        to   { opacity: 1; transform: translateY(0); }
    }

    /* ─── Badge System ────────────────────────────────────────────────── */
    .badge {
        display: inline-block; padding: 3px 10px; border-radius: 6px;
        font-size: 11px; font-weight: 800; text-transform: uppercase;
        letter-spacing: 0.04em;
    }
    .badge-fail     { background: #fee2e2; color: #991b1b; border: 1px solid #fca5a5; }
    .badge-review   { background: #fef3c7; color: #92400e; border: 1px solid #fcd34d; }
    .badge-verified { background: #dcfce7; color: #166534; border: 1px solid #86efac; }

    /* ─── Section Accents ─────────────────────────────────────────────── */
    .section-accent {
        border-left: 5px solid transparent;
        padding-left: 20px; margin: 28px 0;
        border-radius: 4px; transition: all 0.3s ease;
    }
    .accent-fail     { border-left-color: #ef4444; background: rgba(239,68,68,0.03); }
    .accent-review   { border-left-color: #f59e0b; background: rgba(245,158,11,0.03); }
    .accent-verified { border-left-color: #10b981; background: rgba(16,185,129,0.03); }

    /* ─── Audit Output Sections ──────────────────────────────────────────── */
    .output-section {
        background: var(--background-color);
        border: 1px solid #e8edf2;
        border-radius: 10px;
        padding: 20px 24px;
        margin-bottom: 6px;
        line-height: 1.8;
    }
    /* Section category heading (### หมวดหมู่: ...) */
    .output-section h3 {
        font-size: 10px !important;
        font-weight: 800 !important;
        letter-spacing: 0.13em !important;
        text-transform: uppercase !important;
        color: #94a3b8 !important;
        margin: 24px 0 12px !important;
        padding: 0 0 0 10px !important;
        border-left: 3px solid #3b82f6 !important;
        border-bottom: none !important;
    }
    .output-section h3:first-child { margin-top: 0 !important; }
    .output-section p { color: var(--text-color); margin: 3px 0; font-size: 14px; }
    .output-section strong { color: var(--text-color); }
    .output-section hr { border: none; border-top: 1px solid #f1f5f9; margin: 18px 0; }
    /* Table */
    .output-section table {
        width: 100%; border-collapse: collapse;
        border: 1px solid #e8edf2; border-radius: 8px; overflow: hidden; margin: 12px 0;
        font-size: 13.5px;
    }
    .output-section th {
        background: #f8fafc; color: #64748b;
        font-weight: 700; text-transform: uppercase; letter-spacing: 0.07em; font-size: 10px;
        text-align: left; padding: 9px 14px; border-bottom: 1px solid #e8edf2;
    }
    .output-section td {
        padding: 10px 14px; border-bottom: 1px solid #f8fafc;
        vertical-align: top; color: var(--text-color); line-height: 1.65;
    }
    .output-section tr:last-child td { border-bottom: none; }
    .output-section tr:nth-child(even) td { background: #fcfcfd; }
    /* Code inline — muted red, not aggressive */
    .output-section code {
        background: #fff1f2; color: #be123c;
        padding: 1px 6px; border-radius: 4px;
        font-family: 'JetBrains Mono', monospace; font-size: 12.5px;
        border: 1px solid #fecdd3;
    }
    /* Audit score card */
    .score-card {
        background: var(--background-color);
        border: 1px solid #e8edf2;
        border-top: 3px solid #3b82f6;
        border-radius: 10px;
        padding: 18px 22px;
        margin-bottom: 6px;
        display: flex; align-items: center; gap: 20px;
    }
    .score-number { font-size: 40px; font-weight: 800; color: #3b82f6; letter-spacing: -0.04em; line-height: 1; }
    .score-meta { flex: 1; }
    .score-label { font-size: 10px; font-weight: 800; letter-spacing: 0.12em; text-transform: uppercase; color: #94a3b8; margin-bottom: 4px; }
    .score-summary { font-size: 13.5px; color: var(--text-color); line-height: 1.55; }
    /* Done banner */
    .audit-done-banner {
        display: flex; align-items: center; gap: 10px;
        background: #f0fdf4; border: 1px solid #bbf7d0;
        border-left: 4px solid #22c55e;
        border-radius: 8px; padding: 11px 16px;
        color: #166534; font-size: 13px; font-weight: 600;
        margin-bottom: 16px; letter-spacing: 0.01em;
    }
    .audit-done-dot {
        width: 8px; height: 8px; border-radius: 50%;
        background: #22c55e; flex-shrink: 0;
    }

    /* ─── Expander (HTML Code Button) ────────────────────────────────── */
    .stExpander {
        border: 1px solid #e2e8f0 !important;
        border-left: 3px solid #64748b !important;
        border-radius: 8px !important;
        margin: 6px 0 !important;
        overflow: hidden !important;
        box-shadow: none !important;
    }
    .stExpander > details > summary {
        background: #f8fafc !important;
        color: #334155 !important;
        font-weight: 600 !important;
        font-size: 13px !important;
        padding: 10px 16px !important;
        border-bottom: 1px solid #f1f5f9;
        cursor: pointer !important;
    }
    .stExpander > details > summary::after {
        content: "" !important;
        display: inline-block !important;
        width: 0; height: 0 !important;
        border-left: 4px solid transparent !important;
        border-right: 4px solid transparent !important;
        border-top: 5px solid #94a3b8 !important;
        margin-left: 8px !important;
        transition: transform 0.2s ease !important;
        vertical-align: middle !important;
    }
    .stExpander > details[open] > summary::after {
        transform: rotate(180deg) !important;
    }
    .stExpander > details > summary:hover {
        background: #f1f5f9 !important;
        border-left-color: #3b82f6 !important;
    }
    .stExpander > details[open] > summary {
        border-bottom: 1px solid #e2e8f0 !important;
        color: #1e293b !important;
    }
    /* Code block — dark editor theme */
    [data-testid="stCodeBlock"] {
        background-color: #0d1117 !important;
        border-radius: 6px !important;
        border: 1px solid #21262d !important;
        margin: 0 !important;
    }
    [data-testid="stCodeBlock"] code {
        color: #c9d1d9 !important;
        font-family: 'JetBrains Mono', 'SF Mono', 'Consolas', monospace !important;
        font-size: 12.5px !important; line-height: 1.75 !important;
    }
    [data-testid="stCodeBlock"] .token.tag { color: #7ee787 !important; }
    [data-testid="stCodeBlock"] .token.attr-name { color: #79c0ff !important; }
    [data-testid="stCodeBlock"] .token.attr-value,
    [data-testid="stCodeBlock"] .token.string { color: #a5d6ff !important; }
    [data-testid="stCodeBlock"] .token.punctuation { color: #8b949e !important; }


    /* Modal CSS */
    .fixed-overlay {
        position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
        background: rgba(0, 0, 0, 0.55);
        backdrop-filter: blur(8px);
        -webkit-backdrop-filter: blur(8px);
        z-index: 999998;
        animation: overlayFadeIn 0.35s ease-out forwards;
    }
    .fixed-modal {
        position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%);
        border-radius: 24px; padding: 40px; width: 520px; max-width: 92vw;
        box-shadow: 0 32px 64px rgba(0,0,0,0.5), inset 0 1px 0 rgba(255,255,255,0.08) !important;
        text-align: center; z-index: 999999;
        background: rgba(15, 23, 42, 0.5) !important;
        border: 1px solid rgba(255,255,255,0.1) !important;
        color: #f1f5f9 !important;
        animation: modalSlideIn 0.4s cubic-bezier(0.16, 1, 0.3, 1) forwards;
    }
    .fixed-modal h3 { color: #f1f5f9 !important; margin:0 0 8px; font-weight:700; }
    .fixed-modal p  { color: #94a3b8 !important; margin:0; font-size:14px; opacity:0.8; }
    .spinner-loader {
        border: 3px solid rgba(255,255,255,0.1); border-top: 3px solid #10b981; border-radius: 50%;
        width: 32px; height: 32px; animation: spin 0.8s linear infinite; margin: 0 auto 16px auto;
    }
    .progress-container {
        width: 100%; height: 6px; background: rgba(255,255,255,0.1); border-radius: 10px; margin: 20px 0 10px; overflow: hidden;
    }
    .progress-fill {
        height: 100%; background: linear-gradient(90deg, #10b981, #3b82f6); transition: width 0.4s ease;
    }
    .perc-text { font-size: 28px; font-weight: 800; color: #10b981; margin-bottom: 4px; }

    /* Entrance Animation - only fires when class is present */
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


    /* ─── Animations & Keyframes ──────────────────────────────────────── */
    @keyframes badgePulse {
        0%   { box-shadow: 0 0 0 0   rgba(16, 185, 129, 0.4); }
        70%  { box-shadow: 0 0 0 10px rgba(16, 185, 129, 0); }
        100% { box-shadow: 0 0 0 0   rgba(16, 185, 129, 0); }
    }
    @keyframes successGlow {
        0%   { border-color: #10b98122; box-shadow: 0 4px 20px rgba(0,0,0,.06); }
        50%  { border-color: #10b981;   box-shadow: 0 0 25px rgba(16, 185, 129, 0.2); }
        100% { border-color: #10b98122; box-shadow: 0 4px 20px rgba(0,0,0,.06); }
    }
    .report-ready { animation: successGlow 2s ease-out; border: 1px solid #10b981 !important; }
    @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
    @keyframes overlayFadeIn { 0% { opacity: 0; } 100% { opacity: 1; } }
    @keyframes modalSlideIn { 0% { opacity: 0; transform: translate(-50%, -44%) scale(0.95); } 100% { opacity: 1; transform: translate(-50%, -50%) scale(1); } }
    [data-testid="stStatusWidget"] { display: none !important; }
    </style>
    """, unsafe_allow_html=True)

# Call CSS setup
load_css()

# ─── API Key Management ───────────────────────────────────────────────────────
KEY_FILE = ".gemini_key"
def load_key():
    if os.path.exists(KEY_FILE):
        with open(KEY_FILE, "r") as f: return f.read().strip()
    return ""

def save_key(k):
    with open(KEY_FILE, "w") as f: f.write(k.strip())

with st.sidebar:
    st.markdown("<div style='font-size: 20px; font-weight: 800; background: linear-gradient(90deg, #10b981, #3b82f6); -webkit-background-clip: text; -webkit-text-fill-color: transparent;'>EXCEL AUDITOR</div>", unsafe_allow_html=True)
    st.markdown("<div style='font-size: 12px; font-weight: 500; opacity: 0.5; margin-bottom: 32px;'>Data Recheck Tool</div>", unsafe_allow_html=True)

    st.markdown("<div style='font-size: 10px; font-weight: 700; letter-spacing: 0.1em; text-transform: uppercase; opacity: 0.4; margin-bottom: 8px;'>AI Engine</div>", unsafe_allow_html=True)
    saved_key = load_key()
    
    api_key = st.text_input("GEMINI API KEY", type="password", value=saved_key, placeholder="Enter API Key...", label_visibility="collapsed")
    
    if api_key:
        if api_key != saved_key:
            save_key(api_key)
        ok, msg = validate_api_key(api_key)
        if ok:
            st.caption(f"STATUS: {msg}")
        else:
            st.error("Invalid API Key")
    else:
        st.caption("API Key required to run analysis.")

# ─── State Control for Streamlit Re-runs ──────────────────────────────────────
if "has_loaded" not in st.session_state:
    st.session_state.has_loaded = True
    anim_class = "first-run-anim"
else:
    anim_class = ""

# ─── Hero ─────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="hero {anim_class}">
  <div class="h1">DATA AUDITOR</div>
  <div class="sub">Precision Hotel Contract Verification.</div>
</div>
<div class="divider"></div>
""", unsafe_allow_html=True)

# ─── Upload Area ──────────────────────────────────────────────────────────────
col1, col2 = st.columns(2, gap="large")

with col1:
    st.markdown(f'<div class="unified-card {anim_class} anim-delay-1"><div class="c-eye">STEP 1</div><div class="c-ttl">Upload Contract (PDF)</div>', unsafe_allow_html=True)
    pdf_file = st.file_uploader("Upload PDF", type=["pdf"], key="pdf", label_visibility="collapsed")
    if pdf_file: st.success(f"File Ready: {pdf_file.name}")
    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    st.markdown(f'<div class="unified-card {anim_class} anim-delay-2"><div class="c-eye">STEP 2</div><div class="c-ttl">Upload Data (Excel)</div>', unsafe_allow_html=True)
    excel_file = st.file_uploader("Upload Excel", type=["xlsx", "xls"], key="excel", label_visibility="collapsed")
    if excel_file: st.success(f"File Ready: {excel_file.name}")
    st.markdown('</div>', unsafe_allow_html=True)

# ─── Audit Focus Area ────────────────────────────────────────────────────────
_, focus_col, _ = st.columns([0.5, 5, 0.5])
with focus_col:
    st.markdown(f'<div class="unified-card {anim_class} anim-delay-3"><div class="c-eye">STEP 3</div><div class="c-ttl">Audit Focus & Scope</div>', unsafe_allow_html=True)
    focus_options = [
        "Net Price & Extra Beds",
        "Cancellation Policy",
        "Child Policy",
        "Period & Seasons",
        "Meals & Info",
        "All-in-One Full Scan"
    ]
    selected_focus = st.multiselect("Select what you want to audit", options=focus_options, default=["All-in-One Full Scan"], label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ─── Process Button ───────────────────────────────────────────────────────────
ready = bool(pdf_file and excel_file and api_key)

_, btn_col, _ = st.columns([1.5, 3, 1.5])
with btn_col:
    if st.button("Start Specialized Audit  →", type="primary", use_container_width=True, disabled=not ready):
        st.session_state.is_auditing = True
        st.session_state.focus_list = selected_focus

if not ready and not st.session_state.get("is_auditing"):
    hint = "Upload both PDF and Excel files to continue" if not (pdf_file and excel_file) else "Add API Key in sidebar"
    st.markdown(f"<p style='text-align:center;color:#9ca3af;font-size:13px;margin-top:6px'>{hint}</p>", unsafe_allow_html=True)

# ─── Experimental Generator Section [IN TEST] ────────────────────────────────
st.markdown("<br><br>", unsafe_allow_html=True)
exp_col_1, exp_col_2, exp_col_3 = st.columns([1, 4, 1])
with exp_col_2:
    with st.expander("✨ Experimental: AI Excel Generator [IN TEST]", expanded=False):
        st.info("ระบบสร้างไฟล์ Upload อัตโนมัติจาก PDF โดยตรง (เบต้า)")
        gen_ready = bool(pdf_file and api_key)
        if st.button("Generate Upload Excel from PDF →", use_container_width=True, disabled=not gen_ready):
            with st.spinner("AI is analyzing and generating Excel..."):
                try:
                    from utils import extract_pdf_to_excel_json, create_upload_excel
                    result = extract_pdf_to_excel_json(pdf_file.getvalue(), api_key)
                    
                    # Result is now a tuple (data, error_msg)
                    extracted_data = result[0] if isinstance(result, tuple) else result
                    error_detail = result[1] if isinstance(result, tuple) and len(result) > 1 else None

                    if extracted_data:
                        excel_data = create_upload_excel(extracted_data)
                        st.success("Excel generated successfully!")
                        st.download_button(
                            label="📥 Download Generated Upload File",
                            data=excel_data,
                            file_name=f"Generated_Upload_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    else:
                        st.error("AI could not extract data. Please try again.")
                        if error_detail:
                            st.caption("🛠️ Technical Details (Debug):")
                            st.code(error_detail, language="bash")
                except Exception as e:
                    st.error(f"Generation failed: {str(e)}")

st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

# ─── Analysis Output ──────────────────────────────────────────────────────────
if st.session_state.get("is_auditing"):
    st.markdown('<div style="font-size:10px;font-weight:800;letter-spacing:0.15em;text-transform:uppercase;color:#94a3b8;margin-bottom:20px;padding-bottom:12px;border-bottom:1px solid #e8edf2">AUDIT REPORT</div>', unsafe_allow_html=True)
    
    report_placeholder = st.empty()
    modal_placeholder = st.empty()
    
    # Initial Loading State
    focus_text = ", ".join(st.session_state.get("focus_list", [])) or "Full Scan"
    # Phases for progress display
    phases = [
        "Analyzing PDF Structure...",
        "Mapping Excel Columns...",
        "Cross-referencing Data Rows...",
        "Validating Business Logic...",
        "Finalizing Report..."
    ]

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
    
    def apply_badges(text):
        # Restore our specialized UI tags
        text = text.replace("&lt;div class=\"section-accent accent-fail\"&gt;", '<div class="section-accent accent-fail">')
        text = text.replace("&lt;div class=\"section-accent accent-review\"&gt;", '<div class="section-accent accent-review">')
        text = text.replace("&lt;div class=\"section-accent accent-verified\"&gt;", '<div class="section-accent accent-verified">')
        text = text.replace("&lt;/div&gt;", '</div>')
        text = text.replace("&lt;span class=\"badge", '<span class="badge')
        text = text.replace("&lt;/span&gt;", '</span>')
        # Header Accents
        text = text.replace("[SECTION_FAIL]", '\n\n<div class="section-accent accent-fail">\n\n')
        text = text.replace("[SECTION_REVIEW]", '\n\n</div>\n\n<div class="section-accent accent-review">\n\n')
        text = text.replace("[SECTION_VERIFIED]", '\n\n</div>\n\n<div class="section-accent accent-verified">\n\n')
        # Inline Badges
        text = text.replace("[FAIL]", '<span class="badge badge-fail">FAIL</span>')
        text = text.replace("[REVIEW]", '<span class="badge badge-review">REVIEW</span>')
        text = text.replace("[VERIFIED]", '<span class="badge badge-verified">VERIFIED</span>')
        if '<div class="section-accent' in text and text.rstrip()[-6:] != '</div>':
            text += "\n\n</div>\n\n"
        return text

    def render_final_report(full_text):
        """Parse AI response and render HTML code blocks as st.expander() — the ONLY way that works in Streamlit."""
        # 1. Done banner
        st.markdown(
            '<div class="audit-done-banner"><div class="audit-done-dot"></div>Audit complete — report ready for review.</div>',
            unsafe_allow_html=True
        )

        # 2. Auto-detect score from AI text and render as score card
        score_match = re.search(r'\u0e04\u0e30\u0e41\u0e19\u0e19\u0e04\u0e27\u0e32\u0e21\u0e16\u0e39\u0e01\u0e15\u0e49\u0e2d\u0e07.*?(\d+(?:\.\d+)?)\s*%', full_text)
        summary_match_text = re.search(r'\u0e1a\u0e17\u0e2a\u0e23\u0e38\u0e1b.*?:\s*(.+)', full_text)
        if score_match:
            score_val = score_match.group(1)
            score_color = "#22c55e" if float(score_val) >= 90 else "#f59e0b" if float(score_val) >= 70 else "#ef4444"
            summary_text = summary_match_text.group(1).strip() if summary_match_text else ""
            st.markdown(f"""
            <div class="score-card">
                <div class="score-number" style="color:{score_color}">{score_val}%</div>
                <div class="score-meta">
                    <div class="score-label">Accuracy Score</div>
                    <div class="score-summary">{summary_text}</div>
                </div>
            </div>""", unsafe_allow_html=True)

        processed = apply_badges(full_text)
        # Split on <details>...</details> blocks
        parts = re.split(r'(<details>[\s\S]*?</details>)', processed)

        for part in parts:
            stripped = part.strip()
            if not stripped:
                continue

            if stripped.startswith('<details>'):
                # Extract button label from <summary>
                summary_m = re.search(r'<summary>([\s\S]*?)</summary>', part)
                label = "HTML Code — คลิกเพื่อดู / Copy"
                if summary_m:
                    raw_label = summary_m.group(1)
                    label = re.sub(r'<[^>]+>', '', raw_label).strip()

                # Extract HTML code from ```html ... ``` block
                code_match = re.search(r'```html\s*([\s\S]*?)```', part)
                html_code = code_match.group(1).strip() if code_match else ""
                if not html_code:
                    fallback = re.search(r'</summary>([\s\S]*?)</details>', part)
                    html_code = fallback.group(1).strip() if fallback else ""

                with st.expander(label, expanded=False):
                    if html_code:
                        st.code(html_code, language="html")
            else:
                st.markdown(
                    f'<div class="output-section">{stripped}</div>',
                    unsafe_allow_html=True
                )

    full_response = ""
    chunk_count = 0
    try:
        for chunk in stream_recheck_analysis(pdf_file.getvalue(), excel_file.getvalue(), api_key, st.session_state.focus_list):
            chunk_count += 1
            
            # Smart Progress Logic
            phase_idx = min(chunk_count // 12, len(phases) - 1)
            current_perc = min(10 + (chunk_count * 1.5), 98)
            render_modal(current_perc, phases[phase_idx])
            
            if chunk == "[RESET_STREAM]":
                full_response = ""
                continue
            full_response += chunk
            
            # Transform and display
            display_response = apply_badges(full_response)
            report_placeholder.markdown(f'<div class="output-card">\n\n{display_response}▌\n\n</div>', unsafe_allow_html=True)
        
        render_modal(100, "Audit Complete")
        modal_placeholder.empty()
        report_placeholder.empty()
        render_final_report(full_response)
    except Exception as e:
        modal_placeholder.empty()
        st.error(f"Error: {str(e)}")
        
    st.session_state.is_auditing = False
    
    _, reset_btn, _ = st.columns([1.5, 3, 1.5])
    with reset_btn:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("RESET / NEW UPLOAD", use_container_width=True):
            st.rerun()
