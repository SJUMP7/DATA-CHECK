import streamlit as st
import os
from datetime import datetime
from utils_generator import extract_pdf_to_excel_json, create_upload_excel

def render_page_excel(api_key, anim_class):
        if True:
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
                            # imported at top
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
                hint = "Upload PDF file to continue" if not has_pdf_gen else "Enter API Key in Settings"
                st.markdown(f"<p style='text-align:center;color:rgba(130, 130, 130, 0.5);font-size:12px;margin-top:6px;letter-spacing:0.02em'>{hint}</p>", unsafe_allow_html=True)
    
            st.stop()
