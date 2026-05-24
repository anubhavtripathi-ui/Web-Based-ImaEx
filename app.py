import streamlit as st
import pandas as pd
import cv2
import numpy as np
import re
from PIL import Image
from io import BytesIO
from paddleocr import PaddleOCR
from openpyxl.styles import Font, PatternFill

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
<div style='background:white;padding:35px;border-radius:24px;'>
<div style='display:flex;justify-content:center;align-items:center;gap:12px;'>
<div style='font-size:50px;'>⚡</div>
<div style='font-size:56px;font-weight:800;'>ImaEx</div>
</div>
<div style='text-align:center;color:#6b7280;font-size:18px;margin-top:10px;'>
Advanced 10-digit Number Extraction
</div>
</div>
""", unsafe_allow_html=True)

ocr = PaddleOCR(use_angle_cls=True, lang='en', show_log=False)

def reduce_single(num):
    while num > 9:
        num = sum(int(x) for x in str(num))
    return num

def calculate_sum(number):
    try:
        total = sum(int(x) for x in number)
        return reduce_single(total)
    except:
        return "NA"

def validate_number(num):
    if len(num) != 10:
        return False
    if not num.startswith(("6","7","8","9")):
        return False
    return True

def preprocess(img):
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
    gray = cv2.resize(gray, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
    blur = cv2.GaussianBlur(gray, (3,3), 0)

    thresh = cv2.adaptiveThreshold(
        blur,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        15,
        8
    )

    return thresh

def extract_numbers(file):

    image = Image.open(file).convert("RGB")
    img_np = np.array(image)

    processed = preprocess(img_np)

    result = ocr.ocr(processed, cls=True)

    final_rows = []

    if result:

        full_text = ""

        for line in result:

            if line:

                for item in line:

                    try:
                        txt = item[1][0]
                        full_text += " " + txt
                    except:
                        pass

        matches = re.findall(r'\d+', full_text)

        seen = set()

        for raw in matches:

            cleaned = re.sub(r'\D', '', raw)

            if len(cleaned) > 10:
                chunks = re.findall(r'\d{10}', cleaned)
            else:
                chunks = [cleaned]

            for number in chunks:

                if number in seen:
                    continue

                seen.add(number)

                if validate_number(number):

                    number_sum = calculate_sum(number)

                    final_rows.append({
                        "Extracted Number": number,
                        "Number Sum": number_sum
                    })

                else:

                    if len(number) >= 8:

                        final_rows.append({
                            "Extracted Number": number,
                            "Number Sum": "NA"
                        })

    return final_rows

uploaded_files = st.file_uploader(
    "Upload Images",
    type=["jpg","jpeg","png","webp"],
    accept_multiple_files=True
)

if uploaded_files:

    if len(uploaded_files) > 30:
        st.error("Maximum 30 images allowed")

    else:

        progress = st.progress(0)

        status = st.empty()

        all_rows = []

        serial = 1

        for idx, file in enumerate(uploaded_files):

            status.info(f"Processing image {idx+1} of {len(uploaded_files)}")

            rows = extract_numbers(file)

            for row in rows:

                all_rows.append({
                    "S.No": serial,
                    "Extracted Number": row["Extracted Number"],
                    "Number Sum": row["Number Sum"]
                })

                serial += 1

            progress.progress((idx+1)/len(uploaded_files))

        status.success("Processing Completed")

        if len(all_rows) == 0:

            st.warning("No valid records detected")

        else:

            df = pd.DataFrame(all_rows)

            st.success(f"Successfully extracted {len(df)} records")

            st.dataframe(df, use_container_width=True)

            output = BytesIO()

            with pd.ExcelWriter(output, engine='openpyxl') as writer:

                df.to_excel(
                    writer,
                    index=False,
                    sheet_name='ImaEx_Output'
                )

                ws = writer.book['ImaEx_Output']

                fill = PatternFill(
                    start_color="1f2937",
                    end_color="1f2937",
                    fill_type="solid"
                )

                for cell in ws[1]:

                    cell.font = Font(
                        bold=True,
                        color="FFFFFF"
                    )

                    cell.fill = fill

            st.download_button(
                label="⬇ Download Excel File",
                data=output.getvalue(),
                file_name="ImaEx_Output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
