import streamlit as st
import anthropic
import base64
import json
import re
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="ImaEx", page_icon="⚡", layout="centered")

st.markdown("""
<style>
.block-container { max-width: 750px !important; padding: 2rem 1.5rem 4rem !important; }
h1 { text-align: center; font-size: 2rem; }
.sub { text-align: center; color: #6b7280; font-size: 0.9rem; margin-bottom: 2rem; }
.stButton > button {
    background: #2d6a4f !important; color: white !important;
    border: none !important; border-radius: 10px !important;
    font-weight: 600 !important; width: 100% !important;
    padding: 0.7rem !important; font-size: 1rem !important;
}
.stDownloadButton > button {
    background: white !important; color: #2d6a4f !important;
    border: 2px solid #2d6a4f !important; border-radius: 10px !important;
    font-weight: 600 !important; width: 100% !important;
    padding: 0.7rem !important; font-size: 1rem !important;
}
.stTextInput input { font-family: monospace !important; font-size: 0.9rem !important; }
table { width: 100%; border-collapse: collapse; margin-top: 1rem; font-size: 0.85rem; }
th { background: #2d6a4f; color: white; padding: 10px; text-align: center; }
td { padding: 8px 12px; border-bottom: 1px solid #e5e7eb; }
tr:nth-child(even) { background: #f9fafb; }
.badge-ok { background: #d1fae5; color: #065f46; padding: 2px 8px; border-radius: 4px; font-weight: 600; }
.badge-any { background: #fee2e2; color: #991b1b; padding: 2px 8px; border-radius: 4px; font-weight: 600; }
#MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)


def digit_sum_single(s):
    if '?' in s:
        return "ANY"
    total = sum(int(d) for d in s if d.isdigit())
    while total >= 10:
        total = sum(int(d) for d in str(total))
    return str(total)


def extract_numbers(api_key, img_bytes, filename):
    ext = filename.lower().rsplit('.', 1)[-1]
    mtype = {"jpg": "image/jpeg", "jpeg": "image/jpeg",
             "png": "image/png", "webp": "image/webp"}.get(ext, "image/jpeg")

    client = anthropic.Anthropic(api_key=api_key)
    b64 = base64.standard_b64encode(img_bytes).decode()

    msg = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=2048,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": mtype, "data": b64}},
                {"type": "text", "text": (
                    "This image contains a table/list with multiple columns.\n"
                    "One column has 10-digit mobile/ID numbers like 8375052028, 8375048954.\n"
                    "Other columns have short 3-4 digit numbers (901, 902, 976) — IGNORE THOSE.\n"
                    "Also ignore text like FREE POOL.\n\n"
                    "Extract ONLY the 10-digit numbers, top to bottom.\n"
                    "If any digit unclear, use ? in that position.\n\n"
                    "Reply with ONLY a JSON array, no other text:\n"
                    '["8375052028","8375052042"]'
                )}
            ]
        }]
    )

    raw = msg.content[0].text.strip()
    raw = re.sub(r'```\w*', '', raw).strip('`').strip()
    m = re.search(r'\[.*\]', raw, re.DOTALL)
    if m:
        try:
            nums = json.loads(m.group())
            return [str(n) for n in nums if re.match(r'^[\d?]{10}$', str(n))]
        except Exception:
            return []
    return []


def build_excel(numbers):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ImaEx Output"

    thin = Side(style='thin', color='CCCCCC')
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)
    ctr = Alignment(horizontal='center', vertical='center')
    lft = Alignment(horizontal='left', vertical='center')

    hfont = Font(bold=True, color='FFFFFF', size=11)
    hfill = PatternFill('solid', start_color='2D6A4F')
    alt = PatternFill('solid', start_color='F0FDF4')
    anyfill = PatternFill('solid', start_color='FEF2F2')

    for col, (h, w) in enumerate(zip(["Sr. No.", "10-Digit Number", "Digit Sum (Single)"], [10, 24, 20]), 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = hfont; c.fill = hfill; c.alignment = ctr; c.border = bdr
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[1].height = 28

    for i, num in enumerate(numbers, 1):
        ds = digit_sum_single(num)
        is_any = ds == "ANY"
        fill = anyfill if is_any else (alt if i % 2 == 0 else None)
        fc = 'B91C1C' if is_any else '111827'
        for col, (val, aln) in enumerate(zip([i, num, ds], [ctr, lft, ctr]), 1):
            c = ws.cell(row=i+1, column=col, value=val)
            c.font = Font(size=10, color=fc)
            c.alignment = aln; c.border = bdr
            if fill: c.fill = fill
        ws.row_dimensions[i+1].height = 20

    ws.freeze_panes = 'A2'
    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf


# ── UI ───────────────────────────────────────────────────────────────────────

st.markdown("# ⚡ ImaEx")
st.markdown('<div class="sub">Image → Extract 10-digit numbers → Excel</div>', unsafe_allow_html=True)

# API Key
api_key = st.text_input(
    "🔑 Anthropic API Key",
    type="password",
    placeholder="sk-ant-api03-...",
    help="Get your key from console.anthropic.com"
)

# Upload
uploaded = st.file_uploader(
    "📁 Upload Images (max 30) — JPG, PNG, WEBP",
    type=["jpg", "jpeg", "png", "webp"],
    accept_multiple_files=True
)

if uploaded and len(uploaded) > 30:
    st.error("Max 30 images allowed.")
    st.stop()

if uploaded:
    st.info(f"{len(uploaded)} image(s) selected")

    if st.button("⚡ Extract Numbers"):
        if not api_key.strip():
            st.error("Please enter your Anthropic API key above.")
            st.stop()

        all_numbers = []
        progress = st.progress(0)

        for idx, f in enumerate(uploaded):
            with st.spinner(f"Processing {f.name} ({idx+1}/{len(uploaded)})..."):
                try:
                    nums = extract_numbers(api_key.strip(), f.read(), f.name)
                    all_numbers.extend(nums)
                    st.success(f"✅ {f.name} → {len(nums)} numbers found")
                except Exception as e:
                    st.error(f"❌ {f.name} → Error: {e}")
            progress.progress((idx + 1) / len(uploaded))

        progress.empty()

        if not all_numbers:
            st.warning("No 10-digit numbers found. Check image quality or API key.")
            st.stop()

        any_c = sum(1 for n in all_numbers if '?' in n)
        c1, c2, c3 = st.columns(3)
        c1.metric("Total", len(all_numbers))
        c2.metric("Clean", len(all_numbers) - any_c)
        c3.metric("Unclear", any_c)

        st.markdown("### Preview (first 15 rows)")
        rows = ""
        for i, num in enumerate(all_numbers[:15], 1):
            ds = digit_sum_single(num)
            badge = f'<span class="badge-any">ANY</span>' if ds == "ANY" else f'<span class="badge-ok">{ds}</span>'
            rows += f"<tr><td style='text-align:center'>{i}</td><td style='font-family:monospace'>{num}</td><td style='text-align:center'>{badge}</td></tr>"

        st.markdown(f"""
        <table>
          <thead><tr><th>Sr. No.</th><th>10-Digit Number</th><th>Digit Sum</th></tr></thead>
          <tbody>{rows}</tbody>
        </table>
        """, unsafe_allow_html=True)

        if len(all_numbers) > 15:
            st.caption(f"... and {len(all_numbers)-15} more rows in the Excel file.")

        st.markdown("---")
        st.download_button(
            "📥 Download Excel",
            data=build_excel(all_numbers),
            file_name="imaex_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
