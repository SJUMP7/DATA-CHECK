"""
app.py — Streamlit UI layer สำหรับ Hotel Audit Desk

Pages:
  - DATA AUDITOR    : ตรวจสอบ Excel vs PDF Contract (utils.py)
  - CONTRACT COMPARE: เปรียบเทียบสัญญา 2 ปี (utils2.py / utils_compare.py)
  - EXCEL GENERATOR : สร้างไฟล์ Excel จาก PDF (utils.py)

State keys: ดู _KEY_* constants บรรทัด 132-146
CSS:        load_css() บรรทัด 149 — tested on streamlit==1.42.0
"""
import os
import re
import json
import copy
from datetime import datetime
import streamlit as st
from gemini_client import validate_api_key
from utils import stream_recheck_analysis
from utils_generator import extract_pdf_to_excel_json, create_upload_excel

from modules.helpers import (
    get_honest_milestone, _clean_json, _repair_json,
    render_audit_modal, render_compare_modal, apply_badges
)

def main():
    st.set_page_config(
        page_title="Recheck Excel Data",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    # ─── Cancellation Interrupt (UX Fix C2: ป้องกัน user ถูก lock ใน loading) ───
    if "cancel_requested" not in st.session_state:
        st.session_state.cancel_requested = False
    
    if st.session_state.get("cancel_btn_trigger") or st.session_state.get("cancel_requested"):
        st.session_state.is_auditing = False
        st.session_state.cancel_requested = False
        st.session_state.pop("audit_done", None)
        st.session_state.pop("cancel_btn_trigger", None)
        st.rerun()
    
    # ─── Reset Confirmation Modal (UX Fix C1: destructive action ต้องมี confirm) ──
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
                            "cached_pdf_bytes", "cached_pdf_name", "cached_excel_bytes", "cached_excel_name", "confirm_reset", "audit_history", "cc_history"]:
                    st.session_state.pop(key, None)
                st.rerun()
        with btn2:
            if st.button("ยกเลิก", use_container_width=True, key="confirm_reset_no"):
                st.session_state.confirm_reset = False
                st.rerun()
        st.stop()
    
    # ─── Initialization of persistent histories ────────────────────────────────
    if "audit_history" not in st.session_state:
        st.session_state.audit_history = []
    if "cc_history" not in st.session_state:
        st.session_state.cc_history = []
    
    # ─── Dynamic Module Loading for COMPARE CONTRACT ─────────────────────────────
    if "active_app" not in st.session_state:
        st.session_state.active_app = "DATA AUDITOR"
    
    if "app_selector_open" not in st.session_state:
        st.session_state.app_selector_open = False
    
    # ─── Load compare modules ─────────────────────────────────────────────────────
    # Cloud: utils_compare.py is at repo root — direct import works.
    # Local dev: utils_compare.py lives in ../COMPARE CONTRACT — added to sys.path.
    compare_utils = None
    compare_excel = None
    try:
        import utils_compare as compare_utils
        try:
            import excel_generator as compare_excel
        except ImportError:
            compare_excel = None
    except ImportError:
        try:
            import sys
            _cc_dir = os.path.abspath(
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "../COMPARE CONTRACT")
            )
            if _cc_dir not in sys.path:
                sys.path.insert(0, _cc_dir)
            import utils_compare as compare_utils
            try:
                import excel_generator as compare_excel
            except ImportError:
                compare_excel = None
        except Exception:
            pass  # CONTRACT COMPARE module not deployed — feature disabled


    
    # ─── Session State Key Constants (prevents silent typo bugs) ─────────────────
    _KEY_AUDIT_DONE   = "audit_done"
    _KEY_IS_AUDITING  = "is_auditing"
    _KEY_SHOW_UPLOAD  = "show_upload"
    _KEY_AUDIT_RESULT = "_audit_result"
    _KEY_FOCUS_LIST   = "focus_list"
    _KEY_PREV_FOCUS   = "prev_focus"
    
    

    # ─── CSS INJECTION ───────────────────────────────────────────────────────────
    def load_css():
        css_path = os.path.join(os.path.dirname(__file__), "assets", "style.css")
        if os.path.exists(css_path):
            with open(css_path, "r", encoding="utf-8") as f:
                css = f.read()
            st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)
        else:
            st.error("CSS file not found!")



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
                except Exception as e:
                    print(f"[DEBUG] {type(e).__name__}: {e}")
        except Exception as e:
            print(f"[DEBUG] save_key failed (readonly fs?): {e}")

    def is_cloud_key():
        return bool(_CLOUD_KEY)

    # apply_badges, get_honest_milestone imported from modules.helpers (canonical)

    # ─── SAFE DEFAULTS (prevent NameError if sidebar errors) ─────────────────────
    saved_key = load_key()
    api_key = saved_key
    selected_page = st.session_state.get("selected_page", "CONTRACT AUDITOR")

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
            st.markdown('<div class="marker-close-bg"></div>', unsafe_allow_html=True)
            if st.button("Close bg", key="close_selector_overlay", use_container_width=True):
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
        st.markdown('<div class="marker-toggle-app"></div>', unsafe_allow_html=True)
        if st.button("Toggle App", key="toggle_selector_btn", use_container_width=True):
            st.session_state.app_selector_open = not st.session_state.app_selector_open
            st.rerun()

        # 4. Floating Popup Card Menu
        if st.session_state.get("app_selector_open", False):
            da_active = "active-item" if active_app == "DATA AUDITOR" else ""
            cc_active = "active-item" if active_app == "CONTRACT COMPARE" else ""
            da_badge = '<span class="active-badge">Active</span>' if active_app == "DATA AUDITOR" else "<!-- inactive -->"
            cc_badge = '<span class="active-badge">Active</span>' if active_app == "CONTRACT COMPARE" else "<!-- inactive -->"

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
            st.markdown('<div class="marker-popup-da"></div>', unsafe_allow_html=True)
            if st.button("Select DA", key="select_da_app", use_container_width=True):
                st.session_state.active_app = "DATA AUDITOR"
                st.session_state.app_selector_open = False
                st.rerun()

            st.markdown('<div class="marker-popup-cc"></div>', unsafe_allow_html=True)
            if st.button("Select CC", key="select_cc_app", use_container_width=True):
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
            # Match professional Material Icons to replace Emojis
            mat_icon = ":material/document_scanner:" # Default (Contract Auditor)
            if opt == "AI EXCEL GENERATOR": mat_icon = ":material/table_view:"
            elif opt == "CONTRACT COMPARE": mat_icon = ":material/difference:"

            is_active = (st.session_state.selected_page == opt)
            btn_type = "primary" if is_active else "secondary"

            # Use native material icon instead of emoji
            if st.button(opt, key=f"nav_btn_{opt}", use_container_width=True, type=btn_type, icon=mat_icon):
                st.session_state.selected_page = opt
                st.rerun()

        selected_page = st.session_state.selected_page
        saved_key = load_key()
        api_key = saved_key

        @st.dialog("Settings")
        def settings_dialog():
            st.markdown("""
                <style>
                div[data-testid="stDialog"] { border-radius: 16px !important; padding: 8px !important; }
                div[data-testid="stDialog"] h2 { padding-bottom: 16px !important; font-family: 'Plus Jakarta Sans', sans-serif !important; font-weight: 700 !important; }
                div[data-testid="stDialog"] .stTextInput label { display: none !important; }
                div[data-testid="stDialog"] .api-label { font-size: 11px; font-weight: 800; color: #94a3b8; letter-spacing: 0.05em; margin-bottom: 8px; text-transform: uppercase; }
                div[data-testid="stDialog"] .caption-text { font-size: 12px; color: #94a3b8; margin-top: 4px; margin-bottom: 24px; font-weight: 500; }
                div[data-testid="stDialog"] button[kind="primary"] { background: #6366f1 !important; color: white !important; font-size: 14px !important; border-radius: 10px !important; padding: 12px !important; border: none !important; font-weight: 600 !important; }
                div[data-testid="stDialog"] button[kind="secondary"] { border: 1px solid #e2e8f0 !important; color: #64748b !important; font-size: 14px !important; border-radius: 10px !important; padding: 12px !important; font-weight: 600 !important; background: white !important; }
                div[data-testid="stDialog"] button[kind="secondary"]:hover { border-color: #94a3b8 !important; color: #334155 !important; }
                </style>
            """, unsafe_allow_html=True)

            st.markdown('<div class="api-label">GEMINI API KEY</div>', unsafe_allow_html=True)
            api_key_input = st.text_input("GEMINI API KEY", type="password", value=load_key() or "", label_visibility="collapsed")
            st.markdown('<div class="caption-text">Key ถูกบันทึกลงไฟล์ .env เฉพาะในเครื่อง (Local) เพื่อความปลอดภัย</div>', unsafe_allow_html=True)

            col1, col2 = st.columns(2)
            with col1:
                if st.button("บันทึก", type="primary", use_container_width=True, key="settings_save_btn"):
                    if api_key_input: save_key(api_key_input)
                    st.rerun()
            with col2:
                if st.button("ยกเลิก", type="secondary", use_container_width=True, key="settings_cancel_btn"):
                    st.rerun()

        # ─── Settings & Profile Area ───
        st.markdown("""
            <style>
            /* Pin Settings to Bottom */
            div[data-testid="stSidebarUserContent"] { position: relative !important; padding-bottom: 120px !important; }
            div[data-testid="stVerticalBlock"]:has(> div > .marker-settings-bottom) {
                position: absolute !important;
                bottom: 20px !important;
                left: 0.75rem !important;
                right: 0.75rem !important;
                width: auto !important;
                z-index: 50;
            }
            /* Force perfect left alignment on the settings button */
            div[data-testid="stElementContainer"]:has(.marker-settings-btn) + div[data-testid="stElementContainer"] button {
                justify-content: flex-start !important;
                padding-left: 14px !important;
            }
            </style>
        """, unsafe_allow_html=True)

        with st.container():
            st.markdown('<div class="marker-settings-bottom"></div>', unsafe_allow_html=True)
            st.markdown('<div class="marker-settings-btn"></div>', unsafe_allow_html=True)

            if st.button("Settings", key="settings_btn", icon=":material/settings:", use_container_width=True):
                settings_dialog()

            if saved_key:
                st.markdown('<div class="api-status api-connected"><div class="api-status-dot"></div>API Connected</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="api-status api-disconnected"><div class="api-status-dot"></div>API Not Connected</div>', unsafe_allow_html=True)

        # ─── Recent History Sidebar ────────────────────────────────────────────
        if active_app == "CONTRACT COMPARE":
            st.markdown("<div style='font-size:10px;font-weight:700;letter-spacing:0.1em;text-transform:uppercase;opacity:0.4;margin-bottom:8px;margin-top:32px;'>RECENT COMPARISONS</div>", unsafe_allow_html=True)
            cc_history = st.session_state.get("cc_history", [])
            if not cc_history:
                st.markdown("<div style='font-size:11px;opacity:0.5;padding-left:4px;'>No comparisons yet.</div>", unsafe_allow_html=True)
            else:
                for entry in cc_history:
                    label = entry.get("name", "Unknown")[:22].upper()
                    st.download_button(
                        label=label,
                        data=entry["data"],
                        file_name=f"{entry['name']}_{entry['timestamp']}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"cc_hist_{entry['timestamp']}",
                        use_container_width=True
                    )

        elif active_app == "DATA AUDITOR":
            st.markdown("<div style='font-size:10px;font-weight:700;letter-spacing:0.1em;text-transform:uppercase;opacity:0.4;margin-bottom:8px;margin-top:32px;'>RECENT AUDITS</div>", unsafe_allow_html=True)
            audit_history = st.session_state.get("audit_history", [])
            if not audit_history:
                st.markdown("<div style='font-size:11px;opacity:0.5;padding-left:4px;'>No audits this session.</div>", unsafe_allow_html=True)
            else:
                for entry in audit_history:
                    score = entry.get("score")
                    score_str = f" — {score:.0f}%" if score is not None else ""
                    color = "#10b981" if (score or 0) >= 90 else "#f59e0b" if (score or 0) >= 70 else "#ef4444"
                    label = entry.get("name", "Audit")[:18]
                    ts = entry.get("timestamp", "")
                    st.markdown(f"""
                        <div style="padding:8px 10px;border-radius:8px;border:1px solid rgba(130,130,130,0.12);
                                    background:var(--secondary-background-color);margin-bottom:6px;cursor:default;">
                            <div style="font-size:11px;font-weight:700;color:var(--text-color);opacity:0.9;">{label.upper()}</div>
                            <div style="font-size:10px;margin-top:2px;color:{color};font-weight:700;">{score_str if score_str else 'In Progress'}</div>
                            <div style="font-size:9px;opacity:0.4;margin-top:1px;">{ts}</div>
                        </div>
                    """, unsafe_allow_html=True)
                    if entry.get("markdown"):
                        st.download_button(
                            label="↓ .md",
                            data=entry["markdown"],
                            file_name=f"{entry['name']}_{ts}.md",
                            mime="text/markdown",
                            key=f"audit_hist_md_{ts}",
                            use_container_width=True
                        )




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

    # ─── THEME OVERRIDE FOR DATA AUDITOR ──────────────────────────────────────────
    if selected_page in ["CONTRACT AUDITOR", "AI EXCEL GENERATOR"]:
        st.markdown("""
        <style>
        /* Override indigo with sky blue theme */
        .c-eye { color: #0ea5e9 !important; }
        .c-ttl { border-left-color: #0ea5e9 !important; }
        .h1 { background: linear-gradient(135deg, #0ea5e9, #38bdf8) !important; -webkit-background-clip: text !important; background-clip: text !important; -webkit-text-fill-color: transparent !important; }
        [data-testid="stMainBlockContainer"] button[kind="primary"] { background: linear-gradient(135deg, #0ea5e9, #38bdf8) !important; box-shadow: 0 4px 14px rgba(14,165,233,0.25) !important; }
        [data-testid="stMainBlockContainer"] button[kind="primary"]:hover { box-shadow: 0 6px 20px rgba(14,165,233,0.35) !important; }
        .score-number { color: #0ea5e9 !important; }
        .score-card { border-top-color: #0ea5e9 !important; }
        .output-section h3 { background: rgba(14,165,233,0.04) !important; border-left-color: #0ea5e9 !important; }
        .spinner-loader { border-top-color: #0ea5e9 !important; }
        .progress-fill { background: linear-gradient(90deg, #0ea5e9, #38bdf8) !important; }
        .perc-text { color: #0ea5e9 !important; }
        details[open] { border-color: rgba(14,165,233,0.22) !important; box-shadow: 0 4px 16px rgba(14,165,233,0.06) !important; }
        summary { color: #0ea5e9 !important; background: rgba(14,165,233,0.03) !important; }
        summary:hover { background: rgba(14,165,233,0.06) !important; }
        .unified-card:hover { border-color: rgba(14,165,233,0.18) !important; box-shadow: 0 4px 16px rgba(14,165,233,0.06) !important; }
        .divider { background: linear-gradient(90deg, transparent, rgba(14,165,233,0.1), transparent) !important; }
        </style>
        """, unsafe_allow_html=True)

    # ─── PAGE ROUTING ──────────────────────────────────────────────────────────────
    if selected_page == "AI EXCEL GENERATOR":
        from modules.page_excel import render_page_excel
        render_page_excel(api_key, anim_class)
    elif selected_page == "CONTRACT COMPARE":
        from modules.page_compare import render_page_compare
        render_page_compare(api_key, compare_utils, compare_excel)
    elif selected_page == "CONTRACT AUDITOR":
        from modules.page_auditor import render_page_auditor
        render_page_auditor(api_key, anim_class)
if __name__ == '__main__':
    main()
