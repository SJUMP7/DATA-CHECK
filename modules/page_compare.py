import streamlit as st
import os
from datetime import datetime
import time
import re
import pandas as pd
import io

def render_page_compare(api_key, compare_utils, compare_excel, render_compare_modal):
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
                        hint = "Upload both contracts to continue" if not (up1 and up2) else "Add API Key in Settings"
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
    
            def render_compare_modal(pct):
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
    
            render_compare_modal(10)
    
            if not up1 or not up2:
                placeholder.empty()
                st.error("Upload files not found. Please upload both contracts and try again.")
                st.session_state.cc_started = False
                st.stop()
    
            if compare_utils is None:
                placeholder.empty()
                st.error("The CONTRACT COMPARE module (`utils_compare.py`) was not found in your cloud deployment. Please ensure you uploaded all files to your GitHub repository.")
                st.session_state.cc_started = False
                st.stop()
    
            pdf1_bytes = up1.getvalue()
            pdf2_bytes = up2.getvalue()
    
            chunks = []
            char_count = 0
            # Dynamic expected chars based on PDF size (heuristic)
            total_pdf_size = len(pdf1_bytes) + len(pdf2_bytes)
            EXPECTED_CHARS = max(4000, total_pdf_size // 150)
    
            cancel_compare_placeholder = st.empty()
            with cancel_compare_placeholder.container():
                st.markdown('<div class="cancel-btn-container">', unsafe_allow_html=True)
                if st.button("CANCEL COMPARISON", key="cc_cancel_btn_trigger"):
                    st.session_state.cc_cancel_requested = True
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)
    
            try:
                for chunk in compare_utils.stream_contract_comparison(pdf1_bytes, pdf2_bytes, api_key):
                    if st.session_state.get("cc_cancel_requested"):
                        break
    
                    if chunk == "[RESET_STREAM]":
                        chunks = []
                        char_count = 0
                        render_compare_modal(10)
                        continue
    
                    chunks.append(chunk)
                    char_count += len(chunk)
                    pct = min(15 + int(char_count / EXPECTED_CHARS * 80), 98)
                    render_compare_modal(pct)
    
                cancel_compare_placeholder.empty()
    
                if st.session_state.get("cc_cancel_requested"):
                    st.session_state.cc_started = False
                    st.session_state.cc_cancel_requested = False
                    placeholder.empty()
                    st.rerun()
    
                render_compare_modal(100)
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
                    print(f"[DEBUG] History write failed: {e}")
    
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
    
