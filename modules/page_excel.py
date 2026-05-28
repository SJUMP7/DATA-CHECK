import streamlit as st
from datetime import datetime
from utils_generator import extract_pdf_to_excel_json, create_upload_excel


def render_page_excel(api_key, anim_class):
    # ── CSS ────────────────────────────────────────────────────────────────────
    st.markdown("""
    <style>
    .xg-step-eye{font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;
        color:#8b5cf6;margin-bottom:4px;font-family:'Plus Jakarta Sans',sans-serif;}
    .xg-step-eye.opt{color:rgba(128,128,128,.45);}
    .xg-step-title{font-size:14px;font-weight:700;margin-bottom:4px;font-family:'Plus Jakarta Sans',sans-serif;}
    .xg-step-desc{font-size:11px;opacity:.55;margin-bottom:14px;line-height:1.5;
        font-weight:500;font-family:'Plus Jakarta Sans',sans-serif;}
    .xg-file-ok{font-size:12px;color:#10b981;font-weight:600;margin-top:8px;
        display:flex;align-items:center;gap:5px;font-family:'Plus Jakarta Sans',sans-serif;}
    .xg-file-cached{font-size:12px;color:#3b82f6;font-weight:600;margin-top:8px;
        font-family:'Plus Jakarta Sans',sans-serif;}
    .xg-progress-card{background:var(--secondary-background-color);
        border:1px solid rgba(128,128,128,.12);border-radius:12px;
        padding:18px 20px;margin-bottom:12px;font-family:'Plus Jakarta Sans',sans-serif;}
    .xg-p-top{display:flex;align-items:center;gap:12px;margin-bottom:12px;}
    .xg-spinner{width:18px;height:18px;border-radius:50%;
        border:2px solid rgba(139,92,246,.15);border-top-color:#8b5cf6;
        animation:xg-spin .8s linear infinite;flex-shrink:0;}
    @keyframes xg-spin{to{transform:rotate(360deg);}}
    .xg-phase{font-size:13px;font-weight:600;flex:1;}
    .xg-track{height:3px;background:rgba(139,92,246,.1);border-radius:99px;
        overflow:hidden;margin-bottom:14px;}
    .xg-fill{height:100%;background:#8b5cf6;border-radius:99px;
        animation:xg-prog 2.5s ease-in-out infinite alternate;}
    @keyframes xg-prog{from{width:25%}to{width:80%}}
    .xg-steps{display:flex;flex-direction:column;gap:7px;}
    .xg-step-row{display:flex;align-items:center;gap:8px;font-size:12px;
        font-weight:500;font-family:'Plus Jakarta Sans',sans-serif;}
    .xg-step-row.done{color:#10b981;opacity:.7;}
    .xg-step-row.active{color:var(--text-color);}
    .xg-step-row.pending{opacity:.3;}
    .xg-done-banner{display:flex;align-items:center;gap:10px;
        background:rgba(16,185,129,.06);border:0.5px solid rgba(16,185,129,.3);
        border-left:3px solid #10b981;border-radius:0 10px 10px 0;
        padding:10px 14px;margin-bottom:14px;font-family:'Plus Jakarta Sans',sans-serif;}
    .xg-done-dot{width:6px;height:6px;border-radius:50%;background:#10b981;flex-shrink:0;}
    .xg-done-text{font-size:13px;font-weight:600;color:#10b981;}
    .xg-result-card{background:var(--secondary-background-color);
        border:1px solid rgba(128,128,128,.12);border-radius:14px;padding:20px 22px;
        font-family:'Plus Jakarta Sans',sans-serif;}
    .xg-result-hdr{display:flex;align-items:center;gap:12px;margin-bottom:16px;}
    .xg-result-icon{width:38px;height:38px;border-radius:9px;
        background:rgba(16,185,129,.1);display:flex;align-items:center;
        justify-content:center;flex-shrink:0;}
    .xg-result-name{font-size:14px;font-weight:700;}
    .xg-result-sub{font-size:11px;opacity:.45;margin-top:2px;}
    .xg-stats{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:16px;}
    .xg-stat{background:var(--background-color);
        border:1px solid rgba(128,128,128,.1);border-radius:9px;padding:12px 14px;}
    .xg-stat-lbl{font-size:10px;font-weight:700;letter-spacing:.08em;
        text-transform:uppercase;opacity:.4;font-family:'Plus Jakarta Sans',sans-serif;}
    .xg-stat-val{font-size:22px;font-weight:700;margin-top:3px;letter-spacing:-.02em;}
    </style>
    """, unsafe_allow_html=True)

    # ─── State: Result ─────────────────────────────────────────────────────────
    if st.session_state.get("gen_done"):
        _bytes  = st.session_state.get("gen_result_bytes")
        _fname  = st.session_state.get("gen_result_name", "Generated_Upload.xlsx")
        _rows   = st.session_state.get("gen_result_rows", 0)
        _rooms  = st.session_state.get("gen_result_rooms", 0)
        _ctypes = st.session_state.get("gen_result_ctypes", 0)

        st.markdown("""
        <div class="xg-done-banner">
            <div class="xg-done-dot"></div>
            <div class="xg-done-text">Excel generated successfully — ready to download</div>
        </div>
        """, unsafe_allow_html=True)

        _icon_svg = (
            '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#10b981"'
            ' stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">'
            '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>'
            '<polyline points="14 2 14 8 20 8"/>'
            '<line x1="12" y1="18" x2="12" y2="12"/>'
            '<polyline points="9 15 12 18 15 15"/></svg>'
        )
        st.markdown(f"""
        <div class="xg-result-card">
            <div class="xg-result-hdr">
                <div class="xg-result-icon">{_icon_svg}</div>
                <div>
                    <div class="xg-result-name">{_fname}</div>
                    <div class="xg-result-sub">Hotel upload file &middot; ready for system import</div>
                </div>
            </div>
            <div class="xg-stats">
                <div class="xg-stat">
                    <div class="xg-stat-lbl">Rows generated</div>
                    <div class="xg-stat-val">{_rows}</div>
                </div>
                <div class="xg-stat">
                    <div class="xg-stat-lbl">Room types</div>
                    <div class="xg-stat-val">{_rooms}</div>
                </div>
                <div class="xg-stat">
                    <div class="xg-stat-lbl">Contract types</div>
                    <div class="xg-stat-val">{_ctypes}</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        _dl_col, _another_col = st.columns([3, 1.5])
        with _dl_col:
            st.download_button(
                "Download Excel file",
                data=_bytes,
                file_name=_fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
                key="gen_download_btn",
            )
        with _another_col:
            if st.button("Generate another", use_container_width=True, key="gen_another_btn"):
                for _k in [
                    "gen_done", "gen_generating", "gen_error",
                    "gen_result_bytes", "gen_result_name",
                    "gen_result_rows", "gen_result_rooms", "gen_result_ctypes",
                    "cached_gen_pdf_bytes", "cached_gen_pdf_name",
                    "cached_gen_excel_bytes", "cached_gen_excel_name",
                ]:
                    st.session_state.pop(_k, None)
                st.rerun()

        st.stop()

    # ─── State: Generating ─────────────────────────────────────────────────────
    if st.session_state.get("gen_generating"):
        _check = (
            '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor"'
            ' stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">'
            '<polyline points="20 6 9 17 4 12"/></svg>'
        )
        _circle = (
            '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor"'
            ' stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">'
            '<circle cx="12" cy="12" r="10"/></svg>'
        )
        _loader = (
            '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#8b5cf6"'
            ' stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"'
            ' style="animation:xg-spin .8s linear infinite">'
            '<path d="M21 12a9 9 0 1 1-6.219-8.56"/></svg>'
        )
        st.markdown(f"""
        <div class="xg-progress-card">
            <div class="xg-p-top">
                <div class="xg-spinner" aria-hidden="true"></div>
                <div class="xg-phase">AI is reading and extracting contract data...</div>
            </div>
            <div class="xg-track"><div class="xg-fill"></div></div>
            <div class="xg-steps">
                <div class="xg-step-row done">{_check} Reading contract PDF</div>
                <div class="xg-step-row active">{_loader} Extracting room types &times; seasons &times; promotions</div>
                <div class="xg-step-row pending">{_circle} Building Excel rows</div>
                <div class="xg-step-row pending">{_circle} Formatting output file</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # Blocking call — progress card already visible in browser
        try:
            _pdf_b  = st.session_state.get("cached_gen_pdf_bytes")
            _xlsx_b = st.session_state.get("cached_gen_excel_bytes")

            _result     = extract_pdf_to_excel_json(_pdf_b, api_key, excel_bytes=_xlsx_b)
            _extracted  = _result[0] if isinstance(_result, tuple) else _result
            _err        = _result[1] if isinstance(_result, tuple) and len(_result) > 1 else None

            if _extracted:
                _excel_out = create_upload_excel(_extracted)
                _ts        = datetime.now().strftime("%Y%m%d_%H%M")
                _fname     = f"Generated_Upload_{_ts}.xlsx"
                _n_rooms   = len({
                    str(r.get("room_name", "")).strip()
                    for r in _extracted if r.get("room_name")
                })
                _n_ctypes  = len({
                    str(r.get("contract_type", "")).strip()
                    for r in _extracted if r.get("contract_type")
                })
                st.session_state.gen_result_bytes  = _excel_out
                st.session_state.gen_result_name   = _fname
                st.session_state.gen_result_rows   = len(_extracted)
                st.session_state.gen_result_rooms  = _n_rooms
                st.session_state.gen_result_ctypes = _n_ctypes
                st.session_state.gen_done          = True
                st.session_state.gen_generating    = False
            else:
                st.session_state.gen_error      = _err or "AI could not extract data. Please try again."
                st.session_state.gen_generating = False
        except Exception as _e:
            st.session_state.gen_error      = str(_e)
            st.session_state.gen_generating = False

        st.rerun()

    # ─── State: Upload form ────────────────────────────────────────────────────
    col1, col2 = st.columns(2, gap="large")

    with col1:
        with st.container(border=True):
            st.markdown('<div class="custom-card-marker"></div>', unsafe_allow_html=True)
            st.markdown(f"""
            <div class="{anim_class} anim-delay-1">
                <div class="xg-step-eye">Step 1 — Required</div>
                <div class="xg-step-title">Contract PDF</div>
                <div class="xg-step-desc">The hotel contract to extract rates, seasons, and policies from</div>
            </div>
            """, unsafe_allow_html=True)
            pdf_file_gen = st.file_uploader(
                "Upload PDF", type=["pdf"], key="pdf_gen", label_visibility="collapsed"
            )
            if pdf_file_gen:
                st.session_state.cached_gen_pdf_bytes = pdf_file_gen.getvalue()
                st.session_state.cached_gen_pdf_name  = pdf_file_gen.name
                st.markdown(
                    f'<div class="xg-file-ok">'
                    f'<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor"'
                    f' stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">'
                    f'<polyline points="20 6 9 17 4 12"/></svg>'
                    f'{pdf_file_gen.name}</div>',
                    unsafe_allow_html=True,
                )
            elif st.session_state.get("cached_gen_pdf_name"):
                st.markdown(
                    f'<div class="xg-file-cached">Cached: {st.session_state.cached_gen_pdf_name}</div>',
                    unsafe_allow_html=True,
                )

    with col2:
        with st.container(border=True):
            st.markdown('<div class="custom-card-marker"></div>', unsafe_allow_html=True)
            st.markdown(f"""
            <div class="{anim_class} anim-delay-2">
                <div class="xg-step-eye opt">Step 2 — Optional</div>
                <div class="xg-step-title">Reference Excel</div>
                <div class="xg-step-desc">An existing upload file — helps AI match your column format and style</div>
            </div>
            """, unsafe_allow_html=True)
            excel_ref_file = st.file_uploader(
                "Upload Reference Excel", type=["xlsx", "xls"], key="excel_gen", label_visibility="collapsed"
            )
            if excel_ref_file:
                st.session_state.cached_gen_excel_bytes = excel_ref_file.getvalue()
                st.session_state.cached_gen_excel_name  = excel_ref_file.name
                st.markdown(
                    f'<div class="xg-file-ok">'
                    f'<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor"'
                    f' stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">'
                    f'<polyline points="20 6 9 17 4 12"/></svg>'
                    f'{excel_ref_file.name}</div>',
                    unsafe_allow_html=True,
                )
            elif st.session_state.get("cached_gen_excel_name"):
                st.markdown(
                    f'<div class="xg-file-cached">Cached: {st.session_state.cached_gen_excel_name}</div>',
                    unsafe_allow_html=True,
                )

    st.markdown("<br>", unsafe_allow_html=True)

    has_pdf_gen = bool(pdf_file_gen or st.session_state.get("cached_gen_pdf_bytes"))
    gen_ready   = bool(has_pdf_gen and api_key)

    _, btn_col, _ = st.columns([1.5, 3, 1.5])
    with btn_col:
        if st.button(
            "Generate upload Excel from PDF",
            type="primary",
            use_container_width=True,
            disabled=not gen_ready,
            key="gen_start_btn",
        ):
            st.session_state.gen_generating = True
            st.session_state.pop("gen_error", None)
            st.rerun()

    if st.session_state.get("gen_error"):
        st.error(st.session_state.gen_error)

    if not gen_ready:
        _hint = "Upload a PDF contract to continue" if not has_pdf_gen else "Enter API Key in Settings"
        st.markdown(
            f"<p style='text-align:center;opacity:.5;font-size:12px;"
            f"margin-top:6px;font-weight:500;'>{_hint}</p>",
            unsafe_allow_html=True,
        )

    st.stop()
