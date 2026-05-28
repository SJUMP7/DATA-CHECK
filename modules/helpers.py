"""
helpers.py — Shared utilities for all page modules.
Single source of truth. Do NOT redefine these in app.py or any page module.
"""
import re
import html as _html
import streamlit as st


# ─── Dynamic Honest Milestone Parser ──────────────────────────────────────────
def get_honest_milestone(full_text: str, chunk_count: int) -> tuple[int, str]:
    """
    แปลง streaming text จาก Gemini เป็น phase label สำหรับ progress bar
    Input:  full_text = text ที่สะสมจาก stream, chunk_count = จำนวน chunks
    Output: tuple (เปอร์เซ็นต์, ข้อความ)
    """
    if not full_text:
        return 5, "Initializing AI Engine..."
    if "[SECTION_FAIL]" not in full_text:
        perc = min(10 + (chunk_count * 1.5), 30)
        return perc, "Reading Contract PDF..."
    if "[SECTION_REVIEW]" not in full_text:
        perc = min(30 + (chunk_count * 0.8), 65)
        return perc, "Cross-Referencing Excel Rates..."
    if "คะแนนความถูกต้อง" not in full_text:
        perc = min(65 + (chunk_count * 0.5), 90)
        return perc, "Validating Booking Policies..."
    perc = min(90 + (chunk_count * 0.3), 98)
    return perc, "Assembling Final Audit Score..."


# ─── JSON Helpers ──────────────────────────────────────────────────────────────
def _clean_json(raw: str) -> str:
    text = re.sub(r"```(?:json)?", "", raw).strip()
    s, e = text.find("{"), text.rfind("}")
    if s != -1 and e > s:
        text = text[s: e + 1]
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", text)
    return text


def _repair_json(text: str) -> str:
    """Best version: cleans trailing commas, incomplete keys, and closes brackets."""
    if not text:
        return text
    text = re.sub(r',\s*([\]\}])', r'\1', text)
    text = re.sub(r',?\s*"[^"]*"\s*:\s*[^,}\]]*$', '', text)
    text = re.sub(r',?\s*"[^"]*"\s*:\s*$', '', text)
    text = re.sub(r',?\s*$', '', text)
    stack = []
    for char in text:
        if char == '{':
            stack.append('}')
        elif char == '[':
            stack.append(']')
        elif char in ('}', ']'):
            if stack and stack[-1] == char:
                stack.pop()
    if stack:
        text += "".join(reversed(stack))
    return text


# ─── Badge Renderer — canonical version with html.unescape + div wrappers ─────
def apply_badges(text: str) -> str:
    """
    แปลง section markers จาก Gemini เป็น styled HTML badges + section wrappers
    Markers: [SECTION_FAIL], [SECTION_REVIEW], [SECTION_VERIFIED], [FAIL], [REVIEW], [VERIFIED]
    Pure function — no side effects.
    """
    text = _html.unescape(text)
    text = text.replace("[SECTION_FAIL]",     '\n\n<div class="section-accent accent-fail">\n\n')
    text = text.replace("[SECTION_REVIEW]",   '\n\n</div>\n\n<div class="section-accent accent-review">\n\n')
    text = text.replace("[SECTION_VERIFIED]", '\n\n</div>\n\n<div class="section-accent accent-verified">\n\n')
    text = text.replace("[FAIL]",     '<span class="badge badge-fail">FAIL</span>')
    text = text.replace("[REVIEW]",   '<span class="badge badge-review">REVIEW</span>')
    text = text.replace("[VERIFIED]", '<span class="badge badge-verified">VERIFIED</span>')
    if '<div class="section-accent' in text and not text.rstrip().endswith('</div>'):
        text += "\n\n</div>\n\n"
    # Ensure blank line before bullets
    text = re.sub(r'(?<!\n)\n(• |\* |- (?=\S))', r'\n\n\1', text)
    return text


# ─── Modal Renderers (self-contained, use st.empty() internally) ──────────────
def render_audit_modal(placeholder, perc: int, phase: str, focus_text: str = ""):
    """Renders an inline progress card (replaces the old fixed overlay modal)."""
    _focus_html = (
        f'<div class="da-focus-row">'
        f'<svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor"'
        f' stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">'
        f'<circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="3"/></svg>'
        f'Scope: {focus_text}'
        f'</div>'
    ) if focus_text else ""
    placeholder.markdown(f"""
    <style>
    .da-progress-card{{background:var(--secondary-background-color);border:1px solid rgba(128,128,128,.12);
        border-radius:12px;padding:16px 20px;margin-bottom:10px;font-family:'Plus Jakarta Sans',sans-serif;}}
    .da-p-top{{display:flex;align-items:center;gap:12px;margin-bottom:12px;}}
    .da-spinner{{width:18px;height:18px;border-radius:50%;
        border:2px solid rgba(139,92,246,.15);border-top-color:#8b5cf6;
        animation:da-spin .8s linear infinite;flex-shrink:0;}}
    @keyframes da-spin{{to{{transform:rotate(360deg);}}}}
    .da-phase{{font-size:13px;font-weight:600;flex:1;}}
    .da-pct{{font-family:'JetBrains Mono',monospace;font-size:13px;font-weight:600;color:#8b5cf6;}}
    .da-track{{height:3px;background:rgba(139,92,246,.1);border-radius:99px;overflow:hidden;margin-bottom:10px;}}
    .da-fill{{height:100%;background:#8b5cf6;border-radius:99px;transition:width .4s;}}
    .da-focus-row{{display:flex;align-items:center;gap:5px;font-size:11px;opacity:0.5;font-weight:500;}}
    </style>
    <div class="da-progress-card">
        <div class="da-p-top">
            <div class="da-spinner" aria-hidden="true"></div>
            <div class="da-phase">{phase}</div>
            <div class="da-pct">{int(perc)}%</div>
        </div>
        <div class="da-track"><div class="da-fill" style="width:{perc}%"></div></div>
        {_focus_html}
    </div>
    """, unsafe_allow_html=True)


def render_compare_modal(placeholder, perc: int):
    """Renders the compare progress modal into the given st.empty() placeholder."""
    placeholder.markdown(f'''
    <div class="first-run-anim">
        <div class="perc-text">{perc}%</div>
        <div class="progress-container"><div class="progress-fill" style="width: {perc}%"></div></div>
        <div class="modal-title" style="margin-top: 16px;">Analyzing document structures...</div>
    </div>
    ''', unsafe_allow_html=True)
