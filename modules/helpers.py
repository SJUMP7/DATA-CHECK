import streamlit as st
import re

# ─── Dynamic Honest Milestone Parser ───
def get_honest_milestone(full_text: str, chunk_count: int) -> tuple[int, str]:
    if not full_text:
        return 5, "Initializing AI Engine..."
    if "[SECTION_FAIL]" not in full_text:
        perc = min(10 + (chunk_count * 1.5), 30)
        return perc, "Reading Contract PDF..."
    if "[SECTION_REVIEW]" not in full_text:
        perc = min(30 + (chunk_count * 0.8), 65)
        return perc, "Cross-Referencing Excel Rates..."
    if "สรุปผลการตรวจสอบ" not in full_text:
        perc = min(65 + (chunk_count * 0.5), 90)
        return perc, "Validating Booking Policies..."
    perc = min(90 + (chunk_count * 0.3), 98)
    return perc, "Assembling Final Audit Score..."

# ─── Clean & Repair JSON helper ───
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
    return text

def render_audit_modal(perc: int, text: str):
    st.markdown(f'''
    <div class="first-run-anim">
        <div class="perc-text">{perc}%</div>
        <div class="progress-container"><div class="progress-fill" style="width: {perc}%"></div></div>
        <div class="modal-title" style="margin-top: 16px;">{text}</div>
    </div>
    ''', unsafe_allow_html=True)

def render_compare_modal(perc: int):
    st.markdown(f'''
    <div class="first-run-anim">
        <div class="perc-text">{perc}%</div>
        <div class="progress-container"><div class="progress-fill" style="width: {perc}%"></div></div>
        <div class="modal-title" style="margin-top: 16px;">Analyzing document structures...</div>
    </div>
    ''', unsafe_allow_html=True)

def apply_badges(text: str) -> str:
    text = re.sub(r'\[SECTION_PASS\]', '<div class="section-badge badge-pass">✅ <b>SECTION PASS</b></div>', text)
    text = re.sub(r'\[SECTION_FAIL\]', '<div class="section-badge badge-fail">❌ <b>SECTION FAIL</b></div>', text)
    text = re.sub(r'\[SECTION_REVIEW\]', '<div class="section-badge badge-review">⚠️ <b>NEEDS REVIEW</b></div>', text)
    return text
