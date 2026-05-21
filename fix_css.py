import re

with open('app.py', 'r', encoding='utf-8') as f:
    text = f.read()

def repl_text(m):
    percent = int(m.group(1))
    alpha = percent / 100.0
    return f'rgba(130, 130, 130, {alpha})'

text = re.sub(r'color-mix\(\s*in srgb,\s*var\(--text-color\)\s+(\d+)%,\s*transparent\s*\)', repl_text, text)

def repl_hex(m):
    hex_color = m.group(1)
    percent = int(m.group(2))
    alpha = percent / 100.0
    hex_color = hex_color.lstrip('#')
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return f'rgba({r}, {g}, {b}, {alpha})'

text = re.sub(r'color-mix\(\s*in srgb,\s*(#[0-9a-fA-F]{6})\s+(\d+)%,\s*transparent\s*\)', repl_hex, text)

text = text.replace('div[data-testid="stSidebar"]', '[data-testid="stSidebar"]')

old_radio = """    [data-testid="stSidebar"] div[role="radiogroup"] [data-testid="stRadioButtonCircle"],
    [data-testid="stSidebar"] div[role="radiogroup"] .st-bs,
    [data-testid="stSidebar"] div[role="radiogroup"] label > div:first-child:empty,
    [data-testid="stSidebar"] div[role="radiogroup"] label > div:first-child,
    [data-testid="stSidebar"] div[role="radiogroup"] label[data-baseweb="radio"] > div:first-child {"""

new_radio = """    [data-testid="stSidebar"] div[role="radiogroup"] label > div:first-of-type,
    [data-testid="stSidebar"] div[role="radiogroup"] [data-testid="stRadioButtonCircle"],
    [data-testid="stSidebar"] div[role="radiogroup"] .st-bs {"""

text = text.replace(old_radio, new_radio)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(text)

print("Regex replacements complete!")
