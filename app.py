import streamlit as st
import pandas as pd
import cv2
import numpy as np
import re
from PIL import Image
from io import BytesIO
from paddleocr import PaddleOCR
from openpyxl.styles import Font, PatternFill

# ---------------- PAGE CONFIG ---------------- #

st.set_page_config(
    page_title="ImaEx",
    page_icon="⚡",
    layout="wide"
)

# ---------------- UI ---------------- #

st.markdown("""
<style>

header {visibility:hidden;}
#MainMenu {visibility:hidden;}
footer {visibility:hidden;}

.stApp{
    background:#eef2f7;
}

.main-box{
    background:white;
    border-radius:28px;
    padding:40px;
    box-shadow:0 6px 24px rgba(0,0,0,0.06);
}

.title-row{
    display:flex;
    justify-content:center;
    align-items:center;
    gap:14px;
}

.title{
    font-size:62px;
    font-weight:800;
    color:#1f2937;
}

.subtitle{
    text-align:center;
    color:#6b7280;
    font-size:22px;
    margin-top:10px;
    margin-bottom:30px;
}

.stDownloadButton > button{
    width:100%;
    background:linear-gradient(90deg,#ff9966,#ff5e62);
    color:white;
    border:none;
    border-radius:14px;
    padding:15px;
    font-size:18px;
    font-weight:700;
}

</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-box">', unsafe_allow_html=True)

st.markdown("""
<div class="title-row">
<div style="font-size:52px;">⚡</div>
<div class="title">ImaEx</div>
</div>
""", unsafe_allow_html=True)

st.markdown(
    '<div class="subtitle">Advanced 10-digit Number Extraction</div>',
    unsafe_allow_html=True
)

# ---------------- OCR ENGINE ---------------- #

ocr = PaddleOCR(lang='en')

# ---------------- HELPERS ---------------- #

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

    if not num.startswith(("6", "7", "8", "9")):
        return False

    return True


def preprocess(img):

    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)

    gray = cv2.resize(
        gray,
        None,
        fx=2,
        fy=2,
        interpolation=cv2.INTER_CUBIC
    )

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


# ---------------- EXTRACTION ---------------- #

def extract_numbers(file):

    image = Image.open(file).convert("RGB")

    img_np = np.array(image)

    processed = preprocess(img_np)

    result = ocr.ocr(processed)

    final_rows = []

    full_text = ""

    try:

        if result:

            for page in result:

                if not page:
                    continue

                for line in page:

                    try:

                        txt = line[1][0]

                        full_text += " " + str(txt)

                    except:
                        continue

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

    except Exception as e:

        final_rows.append({
            "Extracted Number": "OCR_ERROR",
            "Number Sum": "NA"
        })

    return final_rows


# ---------------- FILE UPLOAD ---------------- #

uploaded_files = st.file_uploader(
    "Upload Images",
    type=["jpg", "jpeg", "png", "webp"],
    accept_multiple_files=True
)

# ---------------- MAIN PROCESS ---------------- #

if uploaded_files:

    if len(uploaded_files) > 30:

        st.error("Maximum 30 images allowed")

    else:

        progress = st.progress(0)

        status = st.empty()

        all_rows = []

        serial = 1

        for idx, file in enumerate(uploaded_files):

            status.info(
                f"Processing image {idx+1} of {len(uploaded_files)}"
            )

            rows = extract_numbers(file)

            for row in rows:

                all_rows.append({
                    "S.No": serial,
                    "Extracted Number": row["Extracted Number"],
                    "Number Sum": row["Number Sum"]
                })

                serial += 1

            progress.progress((idx + 1) / len(uploaded_files))

        status.success("Processing Completed")

        if len(all_rows) == 0:

            st.warning("No valid numbers detected")

        else:

            df = pd.DataFrame(all_rows)

            st.success(
                f"Successfully extracted {len(df)} records"
            )

            st.dataframe(df, use_container_width=True)

            # ---------------- SUMMARY ---------------- #

            st.markdown("### 📊 Summary")

            st.write(f"Total Records: {len(df)}")

            st.write(
                f"NA Records: {len(df[df['Number Sum'] == 'NA'])}"
            )

            # ---------------- EXCEL EXPORT ---------------- #

            output = BytesIO()

            with pd.ExcelWriter(
                output,
                engine='openpyxl'
            ) as writer:

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

                # AUTO WIDTH

                for column_cells in ws.columns:

                    length = max(
                        len(str(cell.value)) if cell.value else 0
                        for cell in column_cells
                    )

                    ws.column_dimensions[
                        column_cells[0].column_letter
                    ].width = length + 6

            st.download_button(
                label="⬇ Download Excel File",
                data=output.getvalue(),
                file_name="ImaEx_Output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

st.markdown('</div>', unsafe_allow_html=True)
