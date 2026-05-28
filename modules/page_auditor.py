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

        modal_placeholder  = st.empty()
        cancel_placeholder = st.empty()
        report_placeholder = st.empty()

        focus_text = ", ".join(st.session_state.get("focus_list", [])) or "Full Scan"

        # Inline progress card
        render_audit_modal(modal_placeholder, 5, "Initializing AI Engine...", focus_text)

        # Inline cancel button — sits below progress card, above live output
        with cancel_placeholder.container():
            _, _c, _ = st.columns([4, 1.5, 4])
            with _c:
                if st.button("Stop audit", key="cancel_btn_trigger", use_container_width=True):
                    st.session_state.cancel_requested = True
                    st.rerun()

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
                report_placeholder.markdown(
                    f'<div style="font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;'
                    f'opacity:0.4;margin-bottom:8px;padding-bottom:8px;border-bottom:1px solid rgba(130,130,130,.1);">'
                    f'Live output</div>'
                    f'<div class="output-card">\n\n{display_response}▌\n\n</div>',
                    unsafe_allow_html=True
                )

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
        if not saved_result:
            st.stop()

        # ── CSS ────────────────────────────────────────────────────────────────
        st.markdown("""
        <style>
        .da-done-banner{display:flex;align-items:center;gap:10px;background:rgba(16,185,129,.06);
            border:0.5px solid rgba(16,185,129,.3);border-left:3px solid #10b981;
            border-radius:0 10px 10px 0;padding:10px 14px;margin-bottom:14px;}
        .da-done-dot{width:6px;height:6px;border-radius:50%;background:#10b981;flex-shrink:0;}
        .da-done-text{font-size:13px;font-weight:600;color:#10b981;font-family:'Plus Jakarta Sans',sans-serif;}
        .da-score-row{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:14px;}
        .da-score-main{background:var(--secondary-background-color);border:1px solid rgba(128,128,128,.12);
            border-radius:12px;padding:18px 20px;border-top-width:2px;border-top-style:solid;}
        .da-score-val{font-family:'JetBrains Mono',monospace;font-size:40px;font-weight:600;line-height:1;letter-spacing:-.03em;}
        .da-score-lbl{font-size:10px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;opacity:.4;margin-top:4px;font-family:'Plus Jakarta Sans',sans-serif;}
        .da-score-summary{font-size:12px;opacity:.7;margin-top:10px;line-height:1.5;
            border-top:0.5px solid rgba(128,128,128,.12);padding-top:10px;font-family:'Plus Jakarta Sans',sans-serif;}
        .da-score-meta{background:var(--secondary-background-color);border:1px solid rgba(128,128,128,.12);
            border-radius:12px;padding:18px 20px;display:flex;flex-direction:column;gap:10px;}
        .da-meta-row{display:flex;align-items:center;gap:8px;font-size:12px;font-family:'Plus Jakarta Sans',sans-serif;}
        .da-meta-row svg{flex-shrink:0;opacity:.4;}
        .da-meta-lbl{opacity:.5;flex:1;}
        .da-meta-val{font-weight:600;font-size:11px;text-align:right;max-width:160px;
            overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
        .da-report-hdr{font-size:10px;font-weight:800;letter-spacing:.15em;text-transform:uppercase;
            opacity:.4;margin-bottom:16px;padding-bottom:10px;
            border-bottom:1px solid rgba(130,130,130,.1);font-family:'Plus Jakarta Sans',sans-serif;}
        </style>
        """, unsafe_allow_html=True)

        # ── Success banner ─────────────────────────────────────────────────────
        st.markdown("""
        <div class="da-done-banner">
            <div class="da-done-dot"></div>
            <div class="da-done-text">Audit complete — report ready for review</div>
        </div>
        """, unsafe_allow_html=True)
        st.toast("Audit complete")

        # ── Score + Meta cards ─────────────────────────────────────────────────
        score_match        = re.search(r'คะแนนความถูกต้อง.*?(\d+(?:\.\d+)?)\s*%', saved_result)
        summary_match_text = re.search(r'บทสรุป.*?:\s*(.+)', saved_result)
        _pdf_nm   = re.sub(r'\.(pdf|PDF)$', '',  st.session_state.get("cached_pdf_name",   "—"))
        _excel_nm = re.sub(r'\.[Xx][Ll][Ss][Xx]?$', '', st.session_state.get("cached_excel_name", "—"))
        _scope_nm = ", ".join(st.session_state.get("focus_list", [])) or "Full Scan"
        _ts_now   = datetime.now().strftime("%H:%M")

        if score_match:
            score       = float(score_match.group(1))
            score_color = "#10b981" if score >= 90 else "#f59e0b" if score >= 70 else "#ef4444"
            summary_txt = summary_match_text.group(1).strip() if summary_match_text else ""

            def _mrow(icon_d, lbl, val):
                return (
                    f'<div class="da-meta-row">'
                    f'<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor"'
                    f' stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">{icon_d}</svg>'
                    f'<span class="da-meta-lbl">{lbl}</span>'
                    f'<span class="da-meta-val">{val}</span></div>'
                )

            _i_pdf = '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/>'
            _i_xls = '<rect x="2" y="3" width="20" height="14" rx="2"/><line x1="8" y1="21" x2="16" y2="21"/><line x1="12" y1="17" x2="12" y2="21"/>'
            _i_tgt = '<circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/>'
            _i_clk = '<circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>'

            st.markdown(f"""
            <div class="da-score-row">
                <div class="da-score-main" style="border-top-color:{score_color};">
                    <div class="da-score-val" style="color:{score_color};">{score:.0f}%</div>
                    <div class="da-score-lbl">Accuracy score</div>
                    {'<div class="da-score-summary">' + summary_txt + '</div>' if summary_txt else ''}
                </div>
                <div class="da-score-meta">
                    {_mrow(_i_pdf, 'Contract PDF', _pdf_nm)}
                    {_mrow(_i_xls, 'Excel file',   _excel_nm)}
                    {_mrow(_i_tgt, 'Scope',         _scope_nm)}
                    {_mrow(_i_clk, 'Completed',     _ts_now)}
                </div>
            </div>
            """, unsafe_allow_html=True)

        # ── Action bar ─────────────────────────────────────────────────────────
        _ts_export = datetime.now().strftime("%Y%m%d_%H%M")
        _col_exp, _col_reaudit, _col_reset = st.columns(3)
        with _col_exp:
            st.download_button(
                "⬇ Export .md",
                data=saved_result.encode("utf-8"),
                file_name=f"{_pdf_nm}_{_ts_export}.md",
                mime="text/markdown",
                use_container_width=True,
                key="export_md_btn",
                type="primary",
            )
        with _col_reaudit:
            if st.button("Edit scope / Re-audit", use_container_width=True, key="edit_scope_btm_btn"):
                for _k in ["audit_done", "is_auditing", "_audit_result"]:
                    st.session_state.pop(_k, None)
                st.rerun()
        with _col_reset:
            if st.button("Reset / New upload", use_container_width=True, key="reset_btm_btn"):
                st.session_state.confirm_reset = True
                st.rerun()

        # ── Report body ────────────────────────────────────────────────────────
        st.markdown("<div class='da-report-hdr' style='margin-top:20px;'>Audit Report</div>", unsafe_allow_html=True)

        processed = apply_badges(saved_result)
        parts = re.split(r'(?=<details>)|(?<=</details>)', processed)
        for part in parts:
            stripped = part.strip()
            if not stripped:
                continue
            if stripped.startswith('<details>'):
                summary_m  = re.search(r'<summary>([\s\S]*?)</summary>', part)
                label      = "HTML Code — คลิกเพื่อดู / Copy"
                if summary_m:
                    raw_label = summary_m.group(1)
                    label     = re.sub(r'<[^>]+>', '', raw_label).strip()
                code_match = re.search(r'```html\s*([\s\S]*?)```', part)
                html_code  = code_match.group(1).strip() if code_match else ""
                if not html_code:
                    fallback  = re.search(r'</summary>([\s\S]*?)</details>', part)
                    html_code = fallback.group(1).strip() if fallback else ""
                with st.expander(label, expanded=False):
                    if html_code:
                        st.code(html_code, language="html")
            else:
                st.markdown(f'<div class="output-section">{stripped}</div>', unsafe_allow_html=True)
