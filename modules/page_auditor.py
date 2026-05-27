import streamlit as st
import os
from datetime import datetime
import time
import re
from utils import stream_recheck_analysis
from modules.helpers import get_honest_milestone, apply_badges, render_audit_modal


def render_page_auditor(api_key, anim_class):
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

    # ─── Decoupled Upload Panel & Scope Selector ─────────────────────────────────
    with st.expander("UPLOAD & SCOPE SETTINGS", expanded=not _audit_done):
        col1, col2 = st.columns(2, gap="large")

        with col1:
            with st.container(border=True):
                st.markdown('<div class="custom-card-marker"></div>', unsafe_allow_html=True)
                st.markdown(f'<div class="{anim_class} anim-delay-1"><p class="c-eye" style="margin:0;">STEP 1</p><p class="c-ttl" style="margin-top:0; margin-bottom:4px;">Contract PDF</p></div>', unsafe_allow_html=True)
                pdf_file = st.file_uploader("Upload PDF", type=["pdf"], key="pdf", label_visibility="collapsed")
                if pdf_file:
                    st.session_state.cached_pdf_bytes = pdf_file.getvalue()
                    st.session_state.cached_pdf_name = pdf_file.name
                    st.markdown(f'<div style="font-size:12px;color:#10b981;font-weight:600;margin-top:8px;">{pdf_file.name}</div>', unsafe_allow_html=True)
                elif st.session_state.get("cached_pdf_name"):
                    st.markdown(f'<div style="font-size:12px;color:#3b82f6;font-weight:600;margin-top:8px;">CACHED: {st.session_state.cached_pdf_name}</div>', unsafe_allow_html=True)


        with col2:
            with st.container(border=True):
                st.markdown('<div class="custom-card-marker"></div>', unsafe_allow_html=True)
                st.markdown(f'<div class="{anim_class} anim-delay-2"><p class="c-eye" style="margin:0;">STEP 2</p><p class="c-ttl" style="margin-top:0; margin-bottom:4px;">Data Excel</p></div>', unsafe_allow_html=True)
                excel_file = st.file_uploader("Upload Excel", type=["xlsx", "xls"], key="excel", label_visibility="collapsed")
                if excel_file:
                    st.session_state.cached_excel_bytes = excel_file.getvalue()
                    st.session_state.cached_excel_name = excel_file.name
                    st.markdown(f'<div style="font-size:12px;color:#10b981;font-weight:600;margin-top:8px;">{excel_file.name}</div>', unsafe_allow_html=True)
                elif st.session_state.get("cached_excel_name"):
                    st.markdown(f'<div style="font-size:12px;color:#3b82f6;font-weight:600;margin-top:8px;">CACHED: {st.session_state.cached_excel_name}</div>', unsafe_allow_html=True)


        # ─── Audit Focus ──────────────────────────────────────────────────────────
        _, focus_col, _ = st.columns([0.5, 5, 0.5])
        with focus_col:
            with st.container(border=True):
                st.markdown('<div class="custom-card-marker"></div>', unsafe_allow_html=True)
                st.markdown(f'<div class="{anim_class} anim-delay-3"><div class="c-eye">STEP 3</div><div class="c-ttl" style="margin-bottom:8px;">Audit Scope</div></div>', unsafe_allow_html=True)
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


        st.markdown("<br>", unsafe_allow_html=True)

        # ─── Process Button ────────────────────────────────────────────────────────
        has_pdf = bool(pdf_file or st.session_state.get("cached_pdf_bytes"))
        has_excel = bool(excel_file or st.session_state.get("cached_excel_bytes"))
        ready = bool(has_pdf and has_excel and api_key)

        _, btn_col, _ = st.columns([1.5, 3, 1.5])
        with btn_col:
            if st.button("Start Audit", type="primary", use_container_width=True, disabled=not ready):
                # ─── Bug #4 Fix: ล้างสถานะเก่าทั้งหมดก่อนเริ่ม audit ใหม่ ─────────
                st.session_state.is_auditing = True
                st.session_state.audit_done = False
                st.session_state.focus_list = selected_focus
                st.session_state.pop("_audit_result", None)       # ล้าง report เก่า
                st.session_state.pop("cancel_requested", None)    # ล้าง cancel flag เก่า
                st.rerun()


        if not ready and not st.session_state.get("is_auditing"):
            hint = "Upload PDF and Excel files to continue" if not (has_pdf and has_excel) else "Enter API Key in Settings"
            st.markdown(f"<p style='text-align:center; opacity:0.9;font-size:12px;margin-top:6px;letter-spacing:0.02em'>{hint}</p>", unsafe_allow_html=True)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # ─── Analysis Output ──────────────────────────────────────────────────────────
    if st.session_state.get("is_auditing"):
        st.markdown('<div style="font-size:10px;font-weight:800;letter-spacing:0.15em;text-transform:uppercase; opacity:0.9;margin-bottom:20px;padding-bottom:12px;border-bottom:1px solid rgba(130, 130, 130, 0.12)">AUDIT REPORT</div>', unsafe_allow_html=True)

        # ─── Bug #3 Guard: ถ้าไม่มีไฟล์ cache (เกิดจาก Streamlit restart / zombie state) ─────
        _pdf_check = st.session_state.get("cached_pdf_bytes")
        _excel_check = st.session_state.get("cached_excel_bytes")
        if not _pdf_check or not _excel_check:
            st.session_state.is_auditing = False
            st.session_state.pop("audit_done", None)
            st.warning("⚠️ Session หมดอายุ — กรุณาอัปโหลดไฟล์ใหม่อีกครั้งครับ")
            st.rerun()

        report_placeholder = st.empty()
        modal_placeholder = st.empty()
        cancel_placeholder = st.empty()

        focus_text = ", ".join(st.session_state.get("focus_list", [])) or "Full Scan"

        # Initial loading state — using canonical helper from modules.helpers
        render_audit_modal(modal_placeholder, 5, "Initializing AI Engine...", focus_text)

        # Centered floating cancel button above overlay
        with cancel_placeholder.container():
            st.markdown('<div class="cancel-btn-container">', unsafe_allow_html=True)
            if st.button("CANCEL AUDIT", key="cancel_btn_trigger"):
                st.session_state.cancel_requested = True
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

        full_response = ""
        chunk_count = 0
        try:
            pdf_bytes_to_use = st.session_state.get("cached_pdf_bytes")
            excel_bytes_to_use = st.session_state.get("cached_excel_bytes")

            for chunk in stream_recheck_analysis(pdf_bytes_to_use, excel_bytes_to_use, api_key, st.session_state.focus_list):
                chunk_count += 1

                current_perc, current_phase = get_honest_milestone(full_response, chunk_count)
                render_audit_modal(modal_placeholder, current_perc, current_phase, focus_text)

                if chunk == "[RESET_STREAM]":
                    full_response = ""
                    chunk_count = 0  # ← Bug #2 fix: reset counter so milestone % recalculates correctly
                    continue
                full_response += chunk

                display_response = apply_badges(full_response)
                report_placeholder.markdown(f'<div class="output-card">\n\n{display_response}▌\n\n</div>', unsafe_allow_html=True)

            render_audit_modal(modal_placeholder, 100, "Audit Complete", focus_text)
            modal_placeholder.empty()
            cancel_placeholder.empty()

            _is_error = any(kw in full_response for kw in [
                "API Key ไม่ถูกต้อง", "โควต้า API", "ไม่สามารถเชื่อมต่อ",
                "RESOURCE_EXHAUSTED", "INVALID_ARGUMENT", "API_KEY_INVALID"
            ])

            if _is_error:
                report_placeholder.empty()
                st.session_state.is_auditing = False
                st.error(full_response.replace("**", "").strip())
            else:
                # Save to audit history (session-based, cloud-safe)
                _score_m = re.search(r'\u0e04\u0e30\u0e41\u0e19\u0e19\u0e04\u0e27\u0e32\u0e21\u0e16\u0e39\u0e01\u0e15\u0e49\u0e2d\u0e07.*(\d+(?:\.\d+)?)\s*%', full_response)
                _score_val = float(_score_m.group(1)) if _score_m else None
                _pdf_name = st.session_state.get("cached_pdf_name", "Audit")
                _hist_name = re.sub(r'\.(pdf|PDF)$', '', _pdf_name)
                _ts = datetime.now().strftime("%H:%M")
                _hist_entry = {
                    "name": _hist_name,
                    "score": _score_val,
                    "timestamp": _ts,
                    "markdown": full_response.encode("utf-8"),
                }
                if "audit_history" not in st.session_state:
                    st.session_state.audit_history = []
                st.session_state.audit_history.insert(0, _hist_entry)
                st.session_state.audit_history = st.session_state.audit_history[:8]

                st.session_state.audit_done = True
                st.session_state.is_auditing = False
                st.session_state._audit_result = full_response
                st.rerun()

        except Exception as e:
            modal_placeholder.empty()
            st.error(f"Error: {str(e)}")
            st.session_state.is_auditing = False
            # Retry button — cached bytes still in session, no re-upload needed
            if st.button("\U0001f504 Retry with same files", key="retry_audit_btn"):
                st.session_state.is_auditing = True
                st.session_state.audit_done = False
                st.rerun()

    # ─── Show Saved Report (after rerun with audit_done=True) ────────────────────
    if st.session_state.get("audit_done") and not st.session_state.get("is_auditing"):
        saved_result = st.session_state.get("_audit_result", "")
        if saved_result:
            st.markdown('<div style="font-size:10px;font-weight:800;letter-spacing:0.15em;text-transform:uppercase; opacity:0.9;margin-bottom:20px;padding-bottom:12px;border-bottom:1px solid rgba(130, 130, 130, 0.12)">AUDIT REPORT</div>', unsafe_allow_html=True)

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
                        {"<div style='font-size:14px; opacity:0.9;margin-top:12px;'>"+summary_txt+"</div>" if summary_txt else ""}
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
        _, export_btn, reaudit_btn, reset_btn = st.columns([1, 2, 2, 2])
        with export_btn:
            _result_md = st.session_state.get("_audit_result", "")
            _pdf_nm = re.sub(r'\.(pdf|PDF)$', '', st.session_state.get("cached_pdf_name", "Audit"))
            _ts_export = datetime.now().strftime("%Y%m%d_%H%M")
            st.download_button(
                "\u2913 Export .md",
                data=_result_md.encode("utf-8"),
                file_name=f"{_pdf_nm}_{_ts_export}.md",
                mime="text/markdown",
                use_container_width=True,
                key="export_md_btn"
            )
        with reaudit_btn:
            if st.button("EDIT SCOPE / RE-AUDIT", use_container_width=True, key="edit_scope_btm_btn"):
                for key in ["audit_done", "is_auditing", "_audit_result"]:
                    st.session_state.pop(key, None)
                st.rerun()
        with reset_btn:
            if st.button("RESET / NEW UPLOAD", use_container_width=True, key="reset_btm_btn"):
                st.session_state.confirm_reset = True
                st.rerun()
