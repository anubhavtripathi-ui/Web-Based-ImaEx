import streamlit as st
import anthropic
import base64
import json
import re
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="ImaEx",
    page_icon="⚡",
    layout="centered",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>


*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

html, body, [class*="css"], .stApp {
    font-family: 'Plus Jakarta Sans', sans-serif;
    background: #f8f7f4;
    color: #1a1a2e;
}

.stApp { background: #f8f7f4; }

.block-container {
    max-width: 780px !important;
    padding: 2.5rem 2rem 4rem !important;
}

/* Header */
.imaex-header {
    text-align: center;
    margin-bottom: 2.5rem;
}
.imaex-logo {
    display: inline-flex;
    align-items: center;
    gap: 10px;
    margin-bottom: 0.5rem;
}
.logo-icon {
    width: 40px; height: 40px;
    background: #2d6a4f;
    border-radius: 10px;
    display: flex; align-items: center; justify-content: center;
    font-size: 20px;
}
.logo-text {
    font-size: 1.9rem;
    font-weight: 700;
    color: #1a1a2e;
    letter-spacing: -0.5px;
}
.logo-text span { color: #2d6a4f; }
.imaex-sub {
    font-size: 0.9rem;
    color: #6b7280;
    font-weight: 400;
}

/* Card */
.card {
    background: #ffffff;
    border: 1px solid #e8e4df;
    border-radius: 16px;
    padding: 1.5rem;
    margin-bottom: 1.2rem;
}
.card-label {
    font-size: 0.78rem;
    font-weight: 600;
    color: #6b7280;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    margin-bottom: 0.8rem;
}

/* API input */
.stTextInput input {
    font-family: 'DM Mono', monospace !important;
    font-size: 0.85rem !important;
    background: #f8f7f4 !important;
    border: 1px solid #e8e4df !important;
    border-radius: 10px !important;
    padding: 0.65rem 1rem !important;
    color: #1a1a2e !important;
}
.stTextInput input:focus {
    border-color: #2d6a4f !important;
    box-shadow: 0 0 0 3px rgba(45,106,79,0.1) !important;
}
label[data-testid="stWidgetLabel"] {
    font-size: 0.82rem !important;
    font-weight: 600 !important;
    color: #374151 !important;
    margin-bottom: 6px !important;
}

/* File uploader */
[data-testid="stFileUploader"] {
    background: #f8f7f4;
    border: 1.5px dashed #c9c4be;
    border-radius: 14px;
    padding: 1.5rem;
    transition: border-color 0.2s;
}
[data-testid="stFileUploader"]:hover { border-color: #2d6a4f; }
[data-testid="stFileUploaderDropzone"] { background: transparent !important; }

/* Buttons */
.stButton > button {
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.95rem !important;
    background: #2d6a4f !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 12px !important;
    padding: 0.75rem 2rem !important;
    width: 100% !important;
    transition: background 0.2s, transform 0.1s !important;
    letter-spacing: 0.01em !important;
}
.stButton > button:hover { background: #1e4d38 !important; transform: translateY(-1px) !important; }
.stButton > button:active { transform: translateY(0) !important; }
.stButton > button:disabled { background: #9ca3af !important; }

/* Download button */
.stDownloadButton > button {
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.95rem !important;
    background: #ffffff !important;
    color: #2d6a4f !important;
    border: 2px solid #2d6a4f !important;
    border-radius: 12px !important;
    padding: 0.75rem 2rem !important;
    width: 100% !important;
    transition: all 0.2s !important;
}
.stDownloadButton > button:hover {
    background: #2d6a4f !important;
    color: #ffffff !important;
}

/* Progress */
.stProgress > div > div { background: #2d6a4f !important; border-radius: 99px; }
.stProgress > div { background: #e8e4df !important; border-radius: 99px; height: 6px !important; }

/* Stats */
.stats-row {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 12px;
    margin: 1.2rem 0;
}
.stat-box {
    background: #f8f7f4;
    border: 1px solid #e8e4df;
    border-radius: 12px;
    padding: 1rem;
    text-align: center;
}
.stat-num { font-size: 1.8rem; font-weight: 700; color: #1a1a2e; }
.stat-num.green { color: #2d6a4f; }
.stat-num.red { color: #b91c1c; }
.stat-label { font-size: 0.75rem; color: #6b7280; margin-top: 2px; font-weight: 500; }

/* Preview table */
.preview-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.82rem;
    margin-top: 0.5rem;
}
.preview-table th {
    background: #f8f7f4;
    color: #6b7280;
    font-weight: 600;
    font-size: 0.72rem;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    padding: 8px 12px;
    text-align: left;
    border-bottom: 1px solid #e8e4df;
}
.preview-table td {
    padding: 9px 12px;
    border-bottom: 0.5px solid #f0ece8;
    color: #1a1a2e;
    vertical-align: middle;
}
.preview-table tr:last-child td { border-bottom: none; }
.mono { font-family: 'DM Mono', monospace; font-size: 0.82rem; letter-spacing: 0.02em; }
.badge-sum {
    display: inline-block;
    background: #d1fae5;
    color: #065f46;
    border-radius: 6px;
    padding: 2px 10px;
    font-family: 'DM Mono', monospace;
    font-size: 0.8rem;
    font-weight: 500;
}
.badge-any {
    display: inline-block;
    background: #fee2e2;
    color: #991b1b;
    border-radius: 6px;
    padding: 2px 10px;
    font-size: 0.78rem;
    font-weight: 500;
}
.num-pill {
    display: inline-block;
    background: #f0f7f4;
    border: 1px solid #c6ddd6;
    border-radius: 6px;
    padding: 3px 10px;
    font-family: 'DM Mono', monospace;
    font-size: 0.82rem;
    color: #1a3d2e;
    letter-spacing: 0.05em;
}

/* Status message */
.status-msg {
    background: #f0f7f4;
    border-left: 3px solid #2d6a4f;
    border-radius: 0 8px 8px 0;
    padding: 0.7rem 1rem;
    font-size: 0.85rem;
    color: #374151;
    margin: 0.5rem 0;
    font-family: 'Plus Jakarta Sans', sans-serif;
}

/* Hide streamlit default elements */
#MainMenu, footer, header { visibility: hidden; }
.stDeployButton { display: none; }
[data-testid="stToolbar"] { display: none; }
</style>
""", unsafe_allow_html=True)


# ── Helpers ────────────────────────────────────────────────────────────────────

def digit_sum_single(number_str: str) -> str:
    """Sum all digits repeatedly until single digit. Returns 'ANY' if '?' present."""
    if '?' in number_str:
        return "ANY"
    digits = re.sub(r'\D', '', number_str)
    if not digits:
        return "ANY"
    total = sum(int(d) for d in digits)
    while total >= 10:
        total = sum(int(d) for d in str(total))
    return str(total)


def to_b64(img_bytes: bytes) -> str:
    return base64.standard_b64encode(img_bytes).decode("utf-8")


def extract_numbers(client: anthropic.Anthropic, img_bytes: bytes) -> list[str]:
    """Send image to Claude Vision, extract 10-digit numbers."""
    resp = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=2048,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": "image/jpeg",
                        "data": to_b64(img_bytes)
                    }
                },
                {
                    "type": "text",
                    "text": (
                        "Examine this image carefully.\n\n"
                        "TASK: Extract every number that is EXACTLY 10 digits long.\n\n"
                        "Rules:\n"
                        "- Only 10-digit numbers. Ignore 3-4 digit serial/row numbers.\n"
                        "- Ignore any text (e.g. 'FREE POOL', labels, headings).\n"
                        "- Read top to bottom.\n"
                        "- If a digit is genuinely unreadable, use '?' for that digit only.\n"
                        "- Do NOT guess uncertain digits — use '?' if unsure.\n\n"
                        "Return ONLY a valid JSON array of strings. No explanation, no markdown.\n"
                        "Example: [\"8375052028\", \"8375052042\"]\n"
                        "If none found: []"
                    )
                }
            ]
        }]
    )
    raw = resp.content[0].text.strip()
    match = re.search(r'\[.*?\]', raw, re.DOTALL)
    if match:
        try:
            nums = json.loads(match.group())
            # Keep only strings that look like 10-char numbers (digits + possible ?)
            return [n for n in nums if isinstance(n, str) and re.match(r'^[\d?]{10}$', n)]
        except json.JSONDecodeError:
            return []
    return []


def build_excel(all_numbers: list[str]) -> io.BytesIO:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ImaEx Output"

    thin = Side(style='thin', color='E0DBD5')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal='center', vertical='center')
    left_align = Alignment(horizontal='left', vertical='center')

    # Header
    headers = ["Sr. No.", "10-Digit Number", "Digit Sum (Single)"]
    col_widths = [12, 26, 22]
    hdr_font = Font(name='Calibri', bold=True, color='FFFFFF', size=11)
    hdr_fill = PatternFill('solid', start_color='2D6A4F')

    for col, (h, w) in enumerate(zip(headers, col_widths), 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = hdr_font
        c.fill = hdr_fill
        c.alignment = center
        c.border = border
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.row_dimensions[1].height = 30

    alt_fill = PatternFill('solid', start_color='F0F7F4')
    any_fill = PatternFill('solid', start_color='FEF2F2')
    any_font_color = 'B91C1C'
    normal_font_color = '111827'

    for i, num in enumerate(all_numbers, 1):
        row = i + 1
        ds = digit_sum_single(num)
        is_any = ds == "ANY"

        row_fill = any_fill if is_any else (alt_fill if i % 2 == 0 else None)
        fcolor = any_font_color if is_any else normal_font_color

        vals = [i, num, ds]
        aligns = [center, left_align, center]

        for col, (val, aln) in enumerate(zip(vals, aligns), 1):
            c = ws.cell(row=row, column=col, value=val)
            c.font = Font(name='Calibri', size=10, color=fcolor)
            c.alignment = aln
            c.border = border
            if row_fill:
                c.fill = row_fill

        ws.row_dimensions[row].height = 22

    ws.freeze_panes = 'A2'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ── UI ─────────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="imaex-header">
  <div class="imaex-logo">
    <div class="logo-icon">⚡</div>
    <div class="logo-text">Ima<span>Ex</span></div>
  </div>
  <div class="imaex-sub">Image → Extract 10-digit numbers → Excel</div>
</div>
""", unsafe_allow_html=True)

# API Key card
st.markdown('<div class="card"><div class="card-label">Anthropic API Key</div>', unsafe_allow_html=True)
api_key = st.text_input(
    "key",
    type="password",
    placeholder="sk-ant-api03-...",
    label_visibility="collapsed"
)
st.markdown('<div style="font-size:0.75rem;color:#9ca3af;margin-top:6px;">Your key is never stored. Get one at <a href="https://console.anthropic.com" style="color:#2d6a4f;">console.anthropic.com</a></div>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# Upload card
st.markdown('<div class="card"><div class="card-label">Upload Images &nbsp;·&nbsp; max 30</div>', unsafe_allow_html=True)
uploaded = st.file_uploader(
    "Drop images here or click to browse",
    type=["jpg", "jpeg", "png", "webp"],
    accept_multiple_files=True,
    label_visibility="collapsed"
)
st.markdown('<div style="font-size:0.75rem;color:#9ca3af;margin-top:8px;">💡 To upload a full folder: open the folder, Ctrl+A to select all, then drag here.</div>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

if uploaded:
    if len(uploaded) > 30:
        st.error("⚠️ Maximum 30 images allowed. Please remove extras.")
        st.stop()

    st.markdown(f'<div class="status-msg">📂 {len(uploaded)} image(s) selected and ready to process.</div>', unsafe_allow_html=True)

    go = st.button("⚡ Extract Numbers")

    if go:
        if not api_key.strip():
            st.error("Please enter your Anthropic API key.")
            st.stop()

        client = anthropic.Anthropic(api_key=api_key.strip())
        all_numbers: list[str] = []
        progress_bar = st.progress(0)
        status_slot = st.empty()

        for idx, file in enumerate(uploaded):
            status_slot.markdown(
                f'<div class="status-msg">🔍 Processing <b>{file.name}</b> ({idx+1}/{len(uploaded)})...</div>',
                unsafe_allow_html=True
            )
            try:
                nums = extract_numbers(client, file.read())
                all_numbers.extend(nums)
                status_slot.markdown(
                    f'<div class="status-msg">✅ <b>{file.name}</b> — {len(nums)} number(s) found</div>',
                    unsafe_allow_html=True
                )
            except Exception as e:
                status_slot.markdown(
                    f'<div class="status-msg" style="border-color:#b91c1c;">❌ <b>{file.name}</b> — Error: {e}</div>',
                    unsafe_allow_html=True
                )
            progress_bar.progress((idx + 1) / len(uploaded))

        progress_bar.empty()
        status_slot.empty()

        if not all_numbers:
            st.warning("No 10-digit numbers were found in any of the uploaded images.")
            st.stop()

        any_count = sum(1 for n in all_numbers if '?' in n)
        clean_count = len(all_numbers) - any_count

        # Stats
        st.markdown(f"""
        <div class="stats-row">
          <div class="stat-box">
            <div class="stat-num">{len(all_numbers)}</div>
            <div class="stat-label">Total Numbers</div>
          </div>
          <div class="stat-box">
            <div class="stat-num green">{clean_count}</div>
            <div class="stat-label">Clean Reads</div>
          </div>
          <div class="stat-box">
            <div class="stat-num red">{any_count}</div>
            <div class="stat-label">Unclear (ANY)</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        # Preview
        st.markdown('<div class="card"><div class="card-label">Preview — first 10 rows</div>', unsafe_allow_html=True)
        rows_html = ""
        for i, num in enumerate(all_numbers[:10], 1):
            ds = digit_sum_single(num)
            badge = f'<span class="badge-any">ANY</span>' if ds == "ANY" else f'<span class="badge-sum">{ds}</span>'
            rows_html += f"""
            <tr>
              <td class="mono" style="color:#6b7280;text-align:center;">{i}</td>
              <td><span class="num-pill">{num}</span></td>
              <td style="text-align:center;">{badge}</td>
            </tr>"""

        st.markdown(f"""
        <table class="preview-table">
          <thead>
            <tr>
              <th style="width:70px;text-align:center;">Sr. No.</th>
              <th>10-Digit Number</th>
              <th style="text-align:center;">Digit Sum</th>
            </tr>
          </thead>
          <tbody>{rows_html}</tbody>
        </table>
        """, unsafe_allow_html=True)

        if len(all_numbers) > 10:
            st.markdown(f'<div style="font-size:0.78rem;color:#9ca3af;margin-top:8px;padding:0 4px;">... and {len(all_numbers)-10} more rows in the Excel file.</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # Build & download
        excel_buf = build_excel(all_numbers)
        st.download_button(
            label="📥 Download Excel",
            data=excel_buf,
            file_name="imaex_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
