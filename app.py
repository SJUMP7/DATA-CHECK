import os
import streamlit as st
from utils import stream_recheck_analysis, validate_api_key

st.set_page_config(
    page_title="Recheck Excel Data",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── CSS ─────────────────────────────────────────────────────────────────────
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

/* Output Card */
.output-card {
    background: var(--background-color); border: 1px solid var(--secondary-background-color); 
    border-radius: 16px; padding: 36px 32px; margin-top: 20px;
    box-shadow: 0 4px 6px -1px rgba(0,0,0,.05);
}

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

.output-card {
    background: var(--background-color); border: 1px solid var(--secondary-background-color); 
    border-radius: 16px; padding: 32px; margin-top: 20px;
    box-shadow: 0 4px 20px rgba(0,0,0,.06); line-height: 1.7;
}
.output-card h3 { font-family: 'Inter', sans-serif !important; font-size: 20px; font-weight: 800; margin: 32px 0 16px 0; color: var(--text-color); border-bottom: 2px solid #10b98122; padding-bottom: 8px; }
.output-card p, .output-card li { font-size: 15px; color: var(--text-color); opacity: 0.85; margin-bottom: 8px; }
.output-card table { width: 100%; border-collapse: collapse; margin: 24px 0; font-size: 14px; border-radius: 12px; overflow: hidden; border: 1px solid #10b98111; }
.output-card th { background: #f8fafc; color: #64748b; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em; font-size: 11px; text-align: left; padding: 14px; border-bottom: 1px solid #e2e8f0; }
.output-card td { padding: 14px; border-bottom: 1px solid #f1f5f9; vertical-align: top; }
.output-card tr:last-child td { border-bottom: none; }
.output-card code { background: #f1f5f9; color: #475569; padding: 2px 6px; border-radius: 4px; font-family: 'JetBrains Mono', monospace; font-size: 13px; }

/* Badge System */
.badge { display: inline-block; padding: 4px 10px; border-radius: 6px; font-size: 12px; font-weight: 800; text-transform: uppercase; letter-spacing: 0.02em; }
.badge-fail { background: #fee2e2; color: #b91c1c; border: 1px solid #fecaca; }
.badge-review { background: #fef3c7; color: #b45309; border: 1px solid #fde68a; }
.badge-verified { 
    background: #dcfce7; color: #15803d; border: 1px solid #bbf7d0; 
    animation: badgePulse 2s infinite;
}

@keyframes badgePulse {
    0% { box-shadow: 0 0 0 0 rgba(16, 185, 129, 0.4); }
    70% { box-shadow: 0 0 0 10px rgba(16, 185, 129, 0); }
    100% { box-shadow: 0 0 0 0 rgba(16, 185, 129, 0); }
}

/* Success Glow for Final Report */
@keyframes successGlow {
    0% { border-color: #10b98122; box-shadow: 0 4px 20px rgba(0,0,0,.06); }
    50% { border-color: #10b981; box-shadow: 0 0 25px rgba(16, 185, 129, 0.2); }
    100% { border-color: #10b98122; box-shadow: 0 4px 20px rgba(0,0,0,.06); }
}
.report-ready {
    animation: successGlow 2s ease-out;
    border: 1px solid #10b981 !important;
}

/* Section Accents */
.section-accent {
    border-left: 5px solid transparent;
    padding-left: 24px;
    margin: 32px 0;
    border-radius: 4px;
    transition: all 0.3s ease;
}
.accent-fail { border-left-color: #ef4444; background: rgba(239, 68, 68, 0.02); }
.accent-review { border-left-color: #f59e0b; background: rgba(245, 158, 11, 0.02); }
.accent-verified { border-left-color: #10b981; background: rgba(16, 185, 129, 0.02); }

@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
@keyframes overlayFadeIn { 0% { opacity: 0; } 100% { opacity: 1; } }
@keyframes modalSlideIn { 0% { opacity: 0; transform: translate(-50%, -44%) scale(0.95); } 100% { opacity: 1; transform: translate(-50%, -50%) scale(1); } }
[data-testid="stStatusWidget"] { display: none !important; }
</style>
""", unsafe_allow_html=True)

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
    st.markdown("### AUDIT REPORT")
    
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
        # Basic protection: escape brackets to prevent raw HTML injection from AI
        # except for the tags we intentionally use.
        text = text.replace("<", "&lt;").replace(">", "&gt;")
        
        # Restore our specialized UI tags
        text = text.replace("&lt;div class=\"section-accent accent-fail\"&gt;", '<div class="section-accent accent-fail">')
        text = text.replace("&lt;div class=\"section-accent accent-review\"&gt;", '<div class="section-accent accent-review">')
        text = text.replace("&lt;div class=\"section-accent accent-verified\"&gt;", '<div class="section-accent accent-verified">')
        text = text.replace("&lt;/div&gt;", '</div>')
        text = text.replace("&lt;span class=\"badge", '<span class="badge')
        text = text.replace("&lt;/span&gt;", '</span>')

        # Header Accents Replacement
        text = text.replace("[SECTION_FAIL]", '<div class="section-accent accent-fail">')
        text = text.replace("[SECTION_REVIEW]", '</div><div class="section-accent accent-review">')
        text = text.replace("[SECTION_VERIFIED]", '</div><div class="section-accent accent-verified">')
        
        # Inline Badges Replacement
        text = text.replace("[FAIL]", '<span class="badge badge-fail">FAIL</span>')
        text = text.replace("[REVIEW]", '<span class="badge badge-review">REVIEW</span>')
        text = text.replace("[VERIFIED]", '<span class="badge badge-verified">VERIFIED</span>')
        
        # Final safety close for sections
        if '<div class="section-accent' in text and '</div>' not in text[-10:]:
            text += "</div>"
        return text

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
            report_placeholder.markdown(f'<div class="output-card">{display_response}▌</div>', unsafe_allow_html=True)
        
        render_modal(100, "Audit Complete")
        modal_placeholder.empty() 
        # Final Clean Render with Badges and Success Glow
        final_response = apply_badges(full_response)
        report_placeholder.markdown(f'<div class="output-card report-ready">{final_response}</div>', unsafe_allow_html=True)
    except Exception as e:
        modal_placeholder.empty()
        st.error(f"Error: {str(e)}")
        
    st.session_state.is_auditing = False
    
    _, reset_btn, _ = st.columns([1.5, 3, 1.5])
    with reset_btn:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("RESET / NEW UPLOAD", use_container_width=True):
            st.rerun()
