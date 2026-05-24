import streamlit as st
import pandas as pd
import anthropic
import base64
import json
import re
from PIL import Image
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="ImaEx", page_icon="⚡", layout="wide")

st.markdown("""
<style>
header {visibility:hidden;}
#MainMenu {visibility:hidden;}
footer {visibility:hidden;}
.stApp{background:#eef2f7;}
.stDownloadButton > button{
width:100%;
background:linear-gradient(90deg,#ff9966,#ff5e62);
color:white;
border:none;
border-radius:12px;
padding:14px;
font-size:18px;
font-weight:700;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div style='background:white;padding:35px;border-radius:24px;margin-bottom:20px;'>
<div style='display:flex;justify-content:center;align-items:center;gap:12px;'>
<div style='font-size:50px;'>⚡</div>
<div style='font-size:56px;font-weight:800;'>ImaEx</div>
</div>
<div style='text-align:center;color:#6b7280;font-size:18px;margin-top:10px;'>
Advanced 10-digit Number Extraction · AI-Powered OCR
</div>
</div>
""", unsafe_allow_html=True)


# ─── Digit-root sum ───────────────────────────────────────────────────────────

def reduce_to_single(num: int) -> int:
    """Repeatedly sum digits until single digit."""
    while num > 9:
        num = sum(int(d) for d in str(num))
    return num


def calculate_sum(number_str: str) -> str:
    """Return single-digit root sum for a 10-digit number string."""
    try:
        total = sum(int(d) for d in number_str)
        return str(reduce_to_single(total))
    except Exception:
        return "N"


# ─── Validation ───────────────────────────────────────────────────────────────

def validate_mobile(num: str) -> bool:
    """Must be exactly 10 digits starting with 6/7/8/9."""
    return len(num) == 10 and num[0] in "6789" and num.isdigit()


# ─── Claude Vision OCR ────────────────────────────────────────────────────────

OCR_PROMPT = """This image shows a table/list of 10-digit Indian mobile phone numbers on a screen.
Each row has a serial number, a 10-digit phone number (starting with 6, 7, 8, or 9), and "FREE POOL".

Extract EVERY 10-digit phone number visible in the image, in order from top to bottom.

Rules:
- Only extract the 10-digit numbers (not serial numbers, not "FREE POOL")
- If a number is clearly visible, include it exactly as shown
- If a number row exists but the number is unreadable/unclear/cut off, include "N" for that row
- Do NOT skip any rows

Return ONLY a valid JSON array of strings. No explanation, no markdown, no extra text.
Example format: ["8375052028", "8375052042", "N", "8375052068"]"""


def extract_numbers_from_image(client: anthropic.Anthropic, image_file) -> list[dict]:
    """Use Claude Vision to extract 10-digit numbers from an image."""
    # Read and encode image
    image_bytes = image_file.read()
    image_file.seek(0)  # reset for any re-reads

    # Detect media type
    suffix = image_file.name.lower().rsplit(".", 1)[-1]
    media_type_map = {"jpg": "image/jpeg", "jpeg": "image/jpeg",
                      "png": "image/png", "webp": "image/webp"}
    media_type = media_type_map.get(suffix, "image/jpeg")

    b64_data = base64.standard_b64encode(image_bytes).decode("utf-8")

    try:
        message = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=2000,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": media_type,
                                "data": b64_data,
                            },
                        },
                        {"type": "text", "text": OCR_PROMPT},
                    ],
                }
            ],
        )

        raw_text = message.content[0].text.strip()

        # Strip any accidental markdown fences
        raw_text = re.sub(r"^```[a-z]*\n?", "", raw_text)
        raw_text = re.sub(r"\n?```$", "", raw_text)

        numbers_list = json.loads(raw_text)

        rows = []
        for item in numbers_list:
            item = str(item).strip()
            if item.upper() == "N" or not item:
                rows.append({"number": "N", "valid": False})
            else:
                # Clean non-digit characters that Claude may occasionally include
                cleaned = re.sub(r"\D", "", item)
                if validate_mobile(cleaned):
                    rows.append({"number": cleaned, "valid": True})
                elif len(cleaned) > 10:
                    # Try to salvage a 10-digit mobile number from a longer string
                    match = re.search(r"[6-9]\d{9}", cleaned)
                    if match:
                        rows.append({"number": match.group(), "valid": True})
                    else:
                        rows.append({"number": "N", "valid": False})
                elif len(cleaned) >= 8:
                    # Partially visible — mark N
                    rows.append({"number": "N", "valid": False})
                else:
                    # Too short — skip serials/junk
                    pass

        return rows

    except json.JSONDecodeError:
        st.warning(f"Could not parse Claude's response for {image_file.name}. "
                   "Falling back to regex extraction.")
        # Fallback: extract any 10-digit sequences from raw text
        found = re.findall(r"[6-9]\d{9}", raw_text)
        return [{"number": n, "valid": True} for n in dict.fromkeys(found)]

    except anthropic.APIError as e:
        st.error(f"API error while processing {image_file.name}: {e}")
        return []


# ─── Excel builder ────────────────────────────────────────────────────────────

def build_excel(df: pd.DataFrame) -> bytes:
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="ImaEx_Output")
        ws = writer.book["ImaEx_Output"]

        # Header style
        hdr_fill = PatternFill(start_color="1f2937", end_color="1f2937", fill_type="solid")
        hdr_font = Font(bold=True, color="FFFFFF", size=12)
        hdr_align = Alignment(horizontal="center", vertical="center")

        for cell in ws[1]:
            cell.font = hdr_font
            cell.fill = hdr_fill
            cell.alignment = hdr_align

        # Alternate row shading + center alignment
        fill_light = PatternFill(start_color="F3F4F6", end_color="F3F4F6", fill_type="solid")
        fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        n_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")  # light red for N

        thin = Side(style="thin", color="D1D5DB")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
            is_n = str(ws.cell(row=row_idx, column=2).value).upper() == "N"
            bg = n_fill if is_n else (fill_light if row_idx % 2 == 0 else fill_white)
            for cell in row:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.fill = bg
                cell.border = border

        # Column widths
        ws.column_dimensions["A"].width = 10   # S.No
        ws.column_dimensions["B"].width = 22   # Extracted Number
        ws.column_dimensions["C"].width = 16   # Number Sum
        ws.row_dimensions[1].height = 28

    return output.getvalue()


# ─── Sidebar: API key ─────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("### ⚙️ Settings")
    api_key_input = st.text_input(
        "Anthropic API Key",
        type="password",
        placeholder="sk-ant-...",
        help="Get your key from console.anthropic.com"
    )
    st.markdown("""
    **How it works:**
    1. Upload screen-photo images
    2. Claude AI reads every number
    3. Download clean Excel output
    
    **Number Sum:** Digit root (e.g. 8+3+7+5+0+4+8+9+5+4 = 53 → 5+3 = 8)
    """)


# ─── Main UI ──────────────────────────────────────────────────────────────────

import os
api_key = api_key_input or os.environ.get("ANTHROPIC_API_KEY", "")

if not api_key:
    st.info("👈 Enter your Anthropic API key in the sidebar to get started.")
    st.stop()

uploaded_files = st.file_uploader(
    "Upload Images (photos of number lists)",
    type=["jpg", "jpeg", "png", "webp"],
    accept_multiple_files=True,
    help="You can upload up to 30 images at once"
)

if uploaded_files:
    if len(uploaded_files) > 30:
        st.error("Maximum 30 images allowed at once.")
        st.stop()

    client = anthropic.Anthropic(api_key=api_key)

    progress_bar = st.progress(0)
    status_box = st.empty()

    all_rows = []
    serial = 1

    for idx, file in enumerate(uploaded_files):
        status_box.info(f"🔍 Processing image {idx + 1} of {len(uploaded_files)}: **{file.name}**")

        rows = extract_numbers_from_image(client, file)

        for row in rows:
            number = row["number"]
            num_sum = calculate_sum(number) if row["valid"] else "N"
            all_rows.append({
                "S.No": serial,
                "Extracted Number": number,
                "Number Sum": num_sum,
            })
            serial += 1

        progress_bar.progress((idx + 1) / len(uploaded_files))

    status_box.success(f"✅ Processing complete! {len(all_rows)} records extracted from {len(uploaded_files)} image(s).")

    if not all_rows:
        st.warning("No numbers were detected. Please check your images.")
        st.stop()

    df = pd.DataFrame(all_rows)

    # Stats
    valid_count = sum(1 for r in all_rows if r["Extracted Number"] != "N")
    n_count = len(all_rows) - valid_count

    col1, col2, col3 = st.columns(3)
    col1.metric("Total Records", len(all_rows))
    col2.metric("Valid Numbers", valid_count)
    col3.metric("Unreadable (N)", n_count)

    st.dataframe(df, use_container_width=True, hide_index=True)

    excel_bytes = build_excel(df)

    st.download_button(
        label="⬇ Download Excel File",
        data=excel_bytes,
        file_name="ImaEx_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
