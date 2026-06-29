import streamlit as st
import os
from datetime import datetime
import time
import re
import pandas as pd
import io
import json
import copy
from modules.helpers import _clean_json, _repair_json, render_compare_modal

def render_page_compare(api_key, compare_utils, compare_excel):
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
        up3 = None
    
        if not is_focus_mode:
            st.markdown("""
                <div class="first-run-anim" style="text-align: center; margin-bottom: 24px; margin-top: 10px;">
                    <p style="font-weight: 800; font-size: 14px; letter-spacing: 0.15em; text-transform: uppercase; color: #0d9488; margin: 0;">Upload Contracts</p>
                    <p style="font-size: 12px; opacity: 0.65; margin-top: 4px;">Choose the previous and new version of contracts to compare</p>
                </div>
            """, unsafe_allow_html=True)
            c1, c2, c3 = st.columns(3, gap="medium")
            with c1:
                with st.container(border=True):
                    st.markdown('<div class="custom-card-marker"></div>', unsafe_allow_html=True)
                    st.markdown('<div class="unified-card-header"><p class="c-eye" style="margin:0;">STEP 1</p><p class="c-ttl" style="margin-top:0; margin-bottom:4px;">Previous Contract</p></div>', unsafe_allow_html=True)
                    up1 = st.file_uploader("Contract 1", type=["pdf"], key="pdf1", label_visibility="collapsed")
                    if up1:
                        st.markdown(f'<div style="font-size:12px;color:#10b981;font-weight:600;margin-top:8px;">{up1.name}</div>', unsafe_allow_html=True)

            with c2:
                with st.container(border=True):
                    st.markdown('<div class="custom-card-marker"></div>', unsafe_allow_html=True)
                    st.markdown('<div class="unified-card-header"><p class="c-eye" style="margin:0;">STEP 2</p><p class="c-ttl" style="margin-top:0; margin-bottom:4px;">New Contract</p></div>', unsafe_allow_html=True)
                    up2 = st.file_uploader("Contract 2", type=["pdf"], key="pdf2", label_visibility="collapsed")
                    if up2:
                        st.markdown(f'<div style="font-size:12px;color:#10b981;font-weight:600;margin-top:8px;">{up2.name}</div>', unsafe_allow_html=True)

            with c3:
                with st.container(border=True):
                    st.markdown('<div class="custom-card-marker"></div>', unsafe_allow_html=True)
                    st.markdown('<div class="unified-card-header"><p class="c-eye" style="margin:0;opacity:0.5;">STEP 3 — OPTIONAL</p><p class="c-ttl" style="margin-top:0; margin-bottom:4px;">Revise Contract</p></div>', unsafe_allow_html=True)
                    up3 = st.file_uploader("Contract 3 (Revise)", type=["pdf"], key="pdf3", label_visibility="collapsed")
                    if up3:
                        st.markdown(f'<div style="font-size:12px;color:#10b981;font-weight:600;margin-top:8px;">{up3.name}</div>', unsafe_allow_html=True)
                    elif st.session_state.get("cc_pdf3_name"):
                        st.markdown(f'<div style="font-size:12px;color:#3b82f6;font-weight:600;margin-top:8px;">Cached: {st.session_state.cc_pdf3_name}</div>', unsafe_allow_html=True)

    
            st.markdown("<br>", unsafe_allow_html=True)
    
        cta_placeholder = st.empty()
    
        # ─── CTA ──────────────────────────────────────────────────────────────────
        ready = bool(up1 and up2 and api_key)  # up3 is optional
    
        if not is_focus_mode:
            up1_name = up1.name if up1 else ""
            up2_name = up2.name if up2 else ""
            up3_name = up3.name if up3 else ""
            if ("cc_last_up1" not in st.session_state
                    or st.session_state.cc_last_up1 != up1_name
                    or st.session_state.cc_last_up2 != up2_name
                    or st.session_state.get("cc_last_up3", "") != up3_name):
                st.session_state.cc_started = False
                st.session_state.cc_review_mode = False
                st.session_state.cc_report_ready = False
                st.session_state.cc_extracted_data = None
                st.session_state.cc_last_up1 = up1_name
                st.session_state.cc_last_up2 = up2_name
                st.session_state.cc_last_up3 = up3_name
    
            with cta_placeholder.container():
                _, btn_col, _ = st.columns([1.5, 3, 1.5])
                with btn_col:
                    if st.button("Compare Contracts  →", type="primary", use_container_width=True, disabled=not ready, key="start_compare_btn"):
                        st.session_state.cc_started = True
                        st.rerun()
    
                    if not ready:
                        hint = "Upload both contracts to continue" if not (up1 and up2) else "Add API Key in Settings"
                        st.markdown(f"<p style='text-align:center;opacity:0.75;font-size:13px;margin-top:6px'>{hint}</p>",
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
                    # Bug C Fix: ล้าง PDF cache เก่า ไม่งั้นจะใช้ไฟล์เก่าต่อ
                    st.session_state.pop("cc_pdf1_bytes", None)
                    st.session_state.pop("cc_pdf2_bytes", None)
                    st.session_state.pop("cc_pdf3_bytes", None)
                    st.session_state.pop("cc_pdf1_name", None)
                    st.session_state.pop("cc_pdf2_name", None)
                    st.session_state.pop("cc_pdf3_name", None)
                    st.session_state.pop("cc_excel_bytes", None)
                    st.rerun()
            st.markdown("<br>", unsafe_allow_html=True)
    
        # ─── Processing ───────────────────────────────────────────────────────────
        if st.session_state.get("cc_started"):
            placeholder = st.empty()
    
            render_compare_modal(placeholder, 10)
    
            # Bug #6 Fix: Cache bytes ทันทีที่เริ่ม หรือ fallback จาก session cache
            if up1:
                st.session_state.cc_pdf1_bytes = up1.getvalue()
                st.session_state.cc_pdf1_name = up1.name
            if up2:
                st.session_state.cc_pdf2_bytes = up2.getvalue()
                st.session_state.cc_pdf2_name = up2.name
            if up3:
                st.session_state.cc_pdf3_bytes = up3.getvalue()
                st.session_state.cc_pdf3_name = up3.name

            pdf1_bytes = st.session_state.get("cc_pdf1_bytes")
            pdf2_bytes = st.session_state.get("cc_pdf2_bytes")
            pdf3_bytes = st.session_state.get("cc_pdf3_bytes")  # None if not uploaded

            if not pdf1_bytes or not pdf2_bytes:
                placeholder.empty()
                st.error("ไม่พบไฟล์ PDF — กรุณาอัปโหลดใหม่อีกครั้งครับ")
                st.session_state.cc_started = False
                st.stop()
    
            if compare_utils is None:
                placeholder.empty()
                st.error("The CONTRACT COMPARE module (`utils_compare.py`) was not found in your cloud deployment. Please ensure you uploaded all files to your GitHub repository.")
                st.session_state.cc_started = False
                st.stop()
    
            # pdf1_bytes / pdf2_bytes already set from session cache above (Bug #6 fix)
    
            chunks = []
            char_count = 0
            # Dynamic expected chars based on PDF size (heuristic)
            total_pdf_size = len(pdf1_bytes) + len(pdf2_bytes) + (len(pdf3_bytes) if pdf3_bytes else 0)
            EXPECTED_CHARS = max(4000, total_pdf_size // 150)
    
            cancel_compare_placeholder = st.empty()
            with cancel_compare_placeholder.container():
                st.markdown('<div class="cancel-btn-container">', unsafe_allow_html=True)
                if st.button("CANCEL COMPARISON", key="cc_cancel_btn_trigger"):
                    st.session_state.cc_cancel_requested = True
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)
    
            try:
                for chunk in compare_utils.stream_contract_comparison(pdf1_bytes, pdf2_bytes, api_key, pdf3_bytes=pdf3_bytes):
                    if st.session_state.get("cc_cancel_requested"):
                        break
    
                    if chunk == "[RESET_STREAM]":
                        chunks = []
                        char_count = 0
                        render_compare_modal(placeholder, 10)
                        continue
    
                    chunks.append(chunk)
                    char_count += len(chunk)
                    pct = min(15 + int(char_count / EXPECTED_CHARS * 80), 98)
                    render_compare_modal(placeholder, pct)
    
                cancel_compare_placeholder.empty()
    
                if st.session_state.get("cc_cancel_requested"):
                    st.session_state.cc_started = False
                    st.session_state.cc_cancel_requested = False
                    placeholder.empty()
                    st.rerun()
    
                render_compare_modal(placeholder, 100)
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
                st.session_state.cc_started = False
                st.error(f"Comparison failed: {str(e)}")
                # Retry: files are still held as uploaded objects in session
                if st.button("\U0001f504 Retry with same files", key="retry_compare_btn"):
                    st.session_state.cc_started = True
                    st.rerun()
    
        # ─── Review & Edit Mode ───────────────────────────────────────────────────
        if st.session_state.get("cc_review_mode"):
            st.markdown("""
                <div style="display:flex; align-items:center; gap:12px; margin-bottom: 20px;">
                    <div style="background:linear-gradient(135deg, #3b82f6, #8b5cf6); border-radius:6px; padding:6px 12px; color:white; font-weight:700; font-size:12px; letter-spacing:1px; box-shadow:0 4px 6px -1px rgba(59, 130, 246, 0.3);">DATA VERIFICATION</div>
                    <div style="font-size:24px; font-weight:700;  letter-spacing:-0.03em;">Review & Edit Prices</div>
                </div>
                <div style="background:var(--secondary-background-color); border-left: 4px solid #3b82f6; border-radius:4px 8px 8px 4px; padding:16px 20px; margin-bottom: 32px;">
                    <p style="margin:0; font-size:14px;  opacity:0.9; line-height:1.6;">
                        AI extraction is complete. Please verify the extracted prices below. You can <b>click any cell to edit</b> the value before finalizing the Excel report.
                    </p>
                </div>
            """, unsafe_allow_html=True)
    
            # Bug A Fix: ใช้ cc_extracted_data จาก session state โดยตรง แล้ว deepcopy ทุกรอบ
            # เพื่อให้ data_editor reflect ค่าล่าสุดเสมอ
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
                        <div style="font-size:17px; font-weight:700; ">{s_name}</div>
                        <div style="font-size:13px; ">
                            <span style="opacity:0.9;">Prev:</span> <span style="font-weight:600; opacity:0.9;">{p1_display}</span> 
                            <span style="margin:0 8px;opacity:0.2;">|</span> 
                            <span style="opacity:0.9;">New:</span> <span style="font-weight:600; color:#3b82f6;">{p2_display}</span>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
    
                rooms = season.get("rooms", [])
                # Bug D Guard: ensure rooms is a list of dicts
                if not isinstance(rooms, list):
                    rooms = []
                if rooms:
                    edited_rooms = st.data_editor(
                        rooms,
                        column_config={
                            "room_name": st.column_config.TextColumn("Room Name", width="large"),
                            "price_1": st.column_config.TextColumn("Contract 1 Price"),
                            "price_2": st.column_config.TextColumn("Contract 2 Price"),
                            "price_3": st.column_config.TextColumn("Revise Price"),
                        },
                        hide_index=True,
                        key=f"cc_editor_season_{i}",
                        use_container_width=True
                    )
                    # Bug A Fix: save edit กลับ session ทันทีเพื่อ survive rerun
                    edited_data["seasons"][i]["rooms"] = edited_rooms
                    st.session_state.cc_extracted_data["seasons"][i]["rooms"] = edited_rooms
    
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
                # Bug B Fix: Cache Excel bytes — อย่า generate ซ้ำทุก rerun
                if "cc_excel_bytes" not in st.session_state or st.session_state.get("cc_excel_bytes") is None:
                    excel_bytes = compare_excel.generate_comparison_excel(data)
                    st.session_state.cc_excel_bytes = excel_bytes
                else:
                    excel_bytes = st.session_state.cc_excel_bytes
    
                timestamp = datetime.now().strftime("%H:%M")
                hotel_name_raw = data.get("hotel_name")
                if not hotel_name_raw or not str(hotel_name_raw).strip() or str(hotel_name_raw).upper() == "HOTEL NAME":
                    hotel_name = "Unknown_Hotel"
                else:
                    hotel_name = str(hotel_name_raw).strip()
                hotel_name_safe = re.sub(r'[\\/*?"<>|]', "", hotel_name)

                # Save to session_state cc_history (cloud-safe, no filesystem)
                if "cc_history" not in st.session_state:
                    st.session_state.cc_history = []
                st.session_state.cc_history.insert(0, {
                    "name": hotel_name_safe,
                    "data": excel_bytes,
                    "timestamp": timestamp,
                })
                st.session_state.cc_history = st.session_state.cc_history[:8]
    
            except Exception as ex:
                st.error(f"Excel generation failed: {ex}")
                st.stop()
    
            # ── Derive display metadata from AI response ──────────────────────────
            year_1 = data.get("year_1", "—")
            year_2 = data.get("year_2", "—")
            year_3 = data.get("year_3", "")  # empty if 2-contract mode

            seasons_list = data.get("seasons", [])
            num_seasons = len(seasons_list)
            all_rooms: set = set()
            for _s in seasons_list:
                for _r in _s.get("rooms", []):
                    _rn = str(_r.get("room_name", "")).strip()
                    if _rn:
                        all_rooms.add(_rn)
            num_rooms = len(all_rooms) if all_rooms else "—"

            _policy_map = [
                ("room_rates",       "Room Rates (all seasons)",    True),   # always present
                ("extra_bed",        "Extra Bed / Extra Person",    False),
                ("early_bird",       "Early Bird Offer",            False),
                ("bonus_night",      "Bonus Night Offer",           False),
                ("wellbeing",        "Wellbeing / Long Stay",       False),
                ("cancellation",     "Cancellation Policy",         False),
                ("other_promotions", "Other Promotions",            False),
            ]
            _sections_status = []
            for _key, _label, _force in _policy_map:
                if _force:
                    _sections_status.append((_label, bool(seasons_list)))
                else:
                    _items = data.get(_key, [])
                    _changed = any(
                        str(i.get("diff_summary", "SAME")).strip().upper() != "SAME"
                        or str(i.get("diff_summary_2", "SAME")).strip().upper() != "SAME"
                        for i in _items
                    ) if _items else False
                    _sections_status.append((_label, _changed))

            num_changed_sections = sum(1 for _, _c in _sections_status if _c)
            total_sections       = len(_sections_status)
            _file_name           = f"{hotel_name_safe}_Comparison.xlsx"
            _words               = hotel_name.split()
            _initials            = "".join(w[0].upper() for w in _words[:2]) if len(_words) >= 2 else hotel_name[:2].upper()

            # ── CSS ───────────────────────────────────────────────────────────────
            st.markdown("""
            <style>
            .cc-success {
                display:flex; align-items:center; gap:10px;
                background:rgba(16,185,129,0.06);
                border:0.5px solid rgba(16,185,129,0.3);
                border-left:3px solid #10b981;
                border-radius:0 10px 10px 0;
                padding:11px 16px; margin-bottom:4px;
            }
            .cc-success-dot { width:7px; height:7px; border-radius:50%; background:#10b981; flex-shrink:0; }
            .cc-success-text { font-size:13px; font-weight:600; color:#10b981; font-family:'Plus Jakarta Sans',sans-serif; }
            .cc-meta {
                background:var(--secondary-background-color);
                border:1px solid rgba(128,128,128,0.12); border-radius:14px;
                padding:18px 22px; display:flex; align-items:center; gap:18px;
                margin-top:12px;
            }
            .cc-avatar {
                width:44px; height:44px; border-radius:10px;
                background:linear-gradient(135deg,#0891b2,#22d3ee);
                display:flex; align-items:center; justify-content:center;
                font-size:14px; font-weight:800; color:#fff; flex-shrink:0; letter-spacing:-0.02em;
                font-family:'Plus Jakarta Sans',sans-serif;
            }
            .cc-meta-body { flex:1; min-width:0; }
            .cc-hotel-name { font-size:15px; font-weight:700; font-family:'Plus Jakarta Sans',sans-serif; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
            .cc-hotel-sub  { font-size:11px; opacity:0.5; margin-top:3px; font-weight:500; font-family:'Plus Jakarta Sans',sans-serif; }
            .cc-pills { display:flex; gap:6px; margin-top:10px; flex-wrap:wrap; }
            .cc-pill { padding:3px 10px; border-radius:20px; font-size:10px; font-weight:700; font-family:'Plus Jakarta Sans',sans-serif; letter-spacing:0.02em; }
            .cc-pill-teal { background:rgba(8,145,178,0.1); color:#0891b2; }
            .cc-pill-gray { background:rgba(128,128,128,0.08); border:0.5px solid rgba(128,128,128,0.2); opacity:0.8; }
            .cc-vs {
                display:flex; align-items:center; gap:12px;
                background:var(--background-color); border:1px solid rgba(128,128,128,0.1);
                border-radius:10px; padding:12px 18px; flex-shrink:0;
            }
            .cc-vs-year { text-align:center; }
            .cc-vs-lbl { font-size:9px; font-weight:700; letter-spacing:0.1em; text-transform:uppercase; opacity:0.4; font-family:'Plus Jakarta Sans',sans-serif; }
            .cc-vs-val { font-family:'JetBrains Mono',monospace; font-size:17px; font-weight:600; margin-top:2px; }
            .cc-vs-arr { font-size:13px; opacity:0.3; }
            .cc-stats { display:grid; grid-template-columns:repeat(3,1fr); gap:10px; margin-top:12px; }
            .cc-stat {
                background:var(--secondary-background-color);
                border:1px solid rgba(128,128,128,0.1); border-radius:10px; padding:14px 16px;
            }
            .cc-stat-lbl { font-size:10px; font-weight:700; letter-spacing:0.08em; text-transform:uppercase; opacity:0.4; font-family:'Plus Jakarta Sans',sans-serif; }
            .cc-stat-val { font-size:26px; font-weight:700; margin-top:4px; letter-spacing:-0.02em; font-family:'Plus Jakarta Sans',sans-serif; }
            .cc-stat-sub { font-size:10px; opacity:0.45; margin-top:2px; font-weight:500; font-family:'Plus Jakarta Sans',sans-serif; }
            .cc-dl-header {
                display:flex; align-items:center; gap:12px;
                background:var(--secondary-background-color);
                border:1px solid rgba(128,128,128,0.12); border-radius:14px 14px 0 0;
                padding:18px 22px; margin-top:12px; border-bottom:none;
            }
            .cc-dl-icon {
                width:38px; height:38px; border-radius:9px;
                background:rgba(16,185,129,0.1);
                display:flex; align-items:center; justify-content:center; flex-shrink:0;
            }
            .cc-dl-title { font-size:14px; font-weight:700; font-family:'Plus Jakarta Sans',sans-serif; }
            .cc-dl-sub   { font-size:11px; opacity:0.45; margin-top:2px; font-family:'Plus Jakarta Sans',sans-serif; }
            .cc-dl-buttons {
                background:var(--secondary-background-color);
                border:1px solid rgba(128,128,128,0.12); border-top:none; border-bottom:none;
                padding:0 22px 16px;
            }
            .cc-dl-hint {
                display:flex; align-items:center; gap:7px;
                background:var(--secondary-background-color);
                border:1px solid rgba(128,128,128,0.12); border-top:none; border-radius:0 0 14px 14px;
                padding:10px 22px; font-size:11px; opacity:0.5; font-weight:500;
                font-family:'Plus Jakarta Sans',sans-serif;
            }
            .cc-sections {
                background:var(--secondary-background-color);
                border:1px solid rgba(128,128,128,0.12); border-radius:14px;
                padding:18px 22px; margin-top:12px;
            }
            .cc-sec-hdr { font-size:10px; font-weight:700; letter-spacing:0.08em; text-transform:uppercase; opacity:0.4; margin-bottom:12px; font-family:'Plus Jakarta Sans',sans-serif; }
            .cc-sec-row { display:flex; align-items:center; gap:10px; padding:9px 0; border-bottom:0.5px solid rgba(128,128,128,0.08); }
            .cc-sec-row:last-child { border-bottom:none; }
            .cc-sec-dot { width:5px; height:5px; border-radius:50%; background:#0891b2; flex-shrink:0; }
            .cc-sec-name { font-size:13px; font-weight:500; flex:1; font-family:'Plus Jakarta Sans',sans-serif; }
            .cc-badge-changed { background:rgba(245,158,11,0.12); color:#b45309; padding:3px 9px; border-radius:20px; font-size:10px; font-weight:700; font-family:'Plus Jakarta Sans',sans-serif; }
            .cc-badge-same    { background:rgba(128,128,128,0.08); opacity:0.6; border:0.5px solid rgba(128,128,128,0.15); padding:3px 9px; border-radius:20px; font-size:10px; font-weight:700; font-family:'Plus Jakarta Sans',sans-serif; }
            </style>
            """, unsafe_allow_html=True)

            # ── Success banner ─────────────────────────────────────────────────────
            st.markdown("""
            <div class="cc-success">
                <div class="cc-success-dot"></div>
                <div class="cc-success-text">Analysis complete — comparison report ready to download</div>
            </div>
            """, unsafe_allow_html=True)

            # ── Meta card ──────────────────────────────────────────────────────────
            _pills_html = (
                f'<div class="cc-pill cc-pill-teal">{total_sections} sections analyzed</div>'
                f'<div class="cc-pill cc-pill-gray">'
                f'{"No changes" if num_changed_sections == 0 else f"{num_changed_sections} changes detected"}'
                f'</div>'
            )
            st.markdown(f"""
            <div class="cc-meta">
                <div class="cc-avatar">{_initials}</div>
                <div class="cc-meta-body">
                    <div class="cc-hotel-name">{hotel_name}</div>
                    <div class="cc-hotel-sub">Compared {"3" if year_3 else "2"} contracts &middot; {timestamp}</div>
                    <div class="cc-pills">{_pills_html}</div>
                </div>
                <div class="cc-vs">
                    <div class="cc-vs-year">
                        <div class="cc-vs-lbl">Previous</div>
                        <div class="cc-vs-val">{year_1}</div>
                    </div>
                    <div class="cc-vs-arr">&rarr;</div>
                    <div class="cc-vs-year">
                        <div class="cc-vs-lbl">New</div>
                        <div class="cc-vs-val">{year_2}</div>
                    </div>
                    {('<div class="cc-vs-arr">&rarr;</div><div class="cc-vs-year"><div class="cc-vs-lbl">Revise</div><div class="cc-vs-val">' + year_3 + '</div></div>') if year_3 else ''}
                </div>
            </div>
            """, unsafe_allow_html=True)

            # ── Stats row ──────────────────────────────────────────────────────────
            _changed_color = "#0891b2" if num_changed_sections > 0 else "inherit"
            st.markdown(f"""
            <div class="cc-stats">
                <div class="cc-stat">
                    <div class="cc-stat-lbl">Room types</div>
                    <div class="cc-stat-val">{num_rooms}</div>
                    <div class="cc-stat-sub">extracted from contract</div>
                </div>
                <div class="cc-stat">
                    <div class="cc-stat-lbl">Seasons</div>
                    <div class="cc-stat-val">{num_seasons}</div>
                    <div class="cc-stat-sub">period blocks compared</div>
                </div>
                <div class="cc-stat">
                    <div class="cc-stat-lbl">Sections changed</div>
                    <div class="cc-stat-val" style="color:{_changed_color};">{num_changed_sections}&thinsp;/&thinsp;{total_sections}</div>
                    <div class="cc-stat-sub">vs previous contract</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            # ── Download card (header + native buttons + hint) ─────────────────────
            st.markdown(f"""
            <div class="cc-dl-header">
                <div class="cc-dl-icon">
                    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24"
                         fill="none" stroke="#10b981" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                        <polyline points="14 2 14 8 20 8"/>
                        <line x1="12" y1="18" x2="12" y2="12"/>
                        <polyline points="9 15 12 18 15 15"/>
                    </svg>
                </div>
                <div>
                    <div class="cc-dl-title">{_file_name}</div>
                    <div class="cc-dl-sub">Excel workbook &middot; Comparison report &middot; {total_sections} sections</div>
                </div>
            </div>
            <div class="cc-dl-buttons">
            """, unsafe_allow_html=True)

            _col_dl, _col_ca = st.columns([3, 1.5])
            with _col_dl:
                st.download_button(
                    "Download Excel Report",
                    data=excel_bytes,
                    file_name=_file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                    key="download_compare_report_btn"
                )
            with _col_ca:
                if st.button("Compare another", use_container_width=True, key="compare_another_btn"):
                    st.session_state.cc_started = False
                    st.session_state.cc_review_mode = False
                    st.session_state.cc_report_ready = False
                    st.session_state.cc_extracted_data = None
                    # Fix: ล้าง PDF cache และ Excel cache ไม่งั้นจะใช้ข้อมูลเก่าต่อ
                    st.session_state.pop("cc_pdf1_bytes", None)
                    st.session_state.pop("cc_pdf2_bytes", None)
                    st.session_state.pop("cc_pdf3_bytes", None)
                    st.session_state.pop("cc_pdf1_name", None)
                    st.session_state.pop("cc_pdf2_name", None)
                    st.session_state.pop("cc_pdf3_name", None)
                    st.session_state.pop("cc_excel_bytes", None)
                    st.rerun()

            st.markdown("""</div>
            <div class="cc-dl-hint">
                <svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24"
                     fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                    <circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/>
                    <line x1="12" y1="8" x2="12.01" y2="8"/>
                </svg>
                To open in Google Sheets: File &rarr; Import &rarr; Upload .xlsx
            </div>
            """, unsafe_allow_html=True)

            # ── Report contents list ────────────────────────────────────────────────
            _rows_html = ""
            for _label, _changed in _sections_status:
                _badge = (
                    '<span class="cc-badge-changed">Updated</span>' if _changed
                    else '<span class="cc-badge-same">Same</span>'
                )
                _rows_html += (
                    f'<div class="cc-sec-row">'
                    f'<div class="cc-sec-dot"></div>'
                    f'<div class="cc-sec-name">{_label}</div>'
                    f'{_badge}'
                    f'</div>'
                )

            st.markdown(
                '<div class="cc-sections">'
                '<div class="cc-sec-hdr">Report contents</div>'
                + _rows_html +
                '</div><br>',
                unsafe_allow_html=True
            )
    
        st.stop()
    
