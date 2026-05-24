
import streamlit as st
import pandas as pd
import cv2
import numpy as np
import re
from PIL import Image
from io import BytesIO
import pytesseract
from openpyxl.styles import Font, PatternFill

st.set_page_config(page_title="ImaEx", page_icon="⚡", layout="centered")

# ---------------- UI ---------------- #
st.markdown("""
<style>

.stApp {
    background: #eef2f7;
}

.block-container {
    padding-top: 2rem;
}

.main-box {
    background: white;
    border-radius: 24px;
    padding: 40px;
    box-shadow: 0 6px 24px rgba(0,0,0,0.05);
}

.logo {
    text-align:center;
    font-size:70px;
}

.title {
    text-align:center;
    font-size:56px;
    font-weight:800;
    color:#1f2937;
}

.subtitle {
    text-align:center;
    color:#6b7280;
    font-size:20px;
    margin-bottom:35px;
}

.stFileUploader {
    background:#f8fafc;
    border-radius:18px;
    padding:18px;
    border:2px dashed #cbd5e1;
}

[data-testid="stFileUploaderDropzone"] {
    background:#f8fafc;
    border:none;
    padding:25px;
}

.processing-text {
    color:#111827;
    font-size:18px;
    font-weight:600;
}

.summary-box {
    background:#f8fafc;
    border-radius:16px;
    padding:20px;
    border:1px solid #e5e7eb;
}

.stDownloadButton > button {
    width:100%;
    background: linear-gradient(90deg,#ff9966,#ff5e62);
    color:white;
    border:none;
    border-radius:14px;
    padding:16px;
    font-size:18px;
    font-weight:700;
}

</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-box">', unsafe_allow_html=True)

st.markdown('<div class="logo">⚡</div>', unsafe_allow_html=True)
st.markdown('<div class="title">ImaEx</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Image → Extract 10-digit numbers → Excel</div>', unsafe_allow_html=True)

st.markdown("### 📁 Upload Images")
st.caption("Upload up to 30 JPG / PNG / WEBP images")

# ---------------- FUNCTIONS ---------------- #

def reduce_to_single_digit(num):
    while num > 9:
        num = sum(int(d) for d in str(num))
    return num

def calculate_sum(number):
    try:
        total = sum(int(x) for x in number)
        final = reduce_to_single_digit(total)

        # double verification
        total2 = sum(int(x) for x in number)
        final2 = reduce_to_single_digit(total2)

        if final == final2:
            return final
        return "NA"

    except:
        return "NA"

def preprocess_image(img_np):

    gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)

    gray = cv2.resize(gray, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)

    gray = cv2.fastNlMeansDenoising(gray)

    kernel = np.array([[0,-1,0],[-1,5,-1],[0,-1,0]])
    sharp = cv2.filter2D(gray, -1, kernel)

    thresh = cv2.adaptiveThreshold(
        sharp,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        11,
        2
    )

    return thresh

def extract_numbers(image_file):

    image = Image.open(image_file).convert("RGB")
    img_np = np.array(image)

    processed = preprocess_image(img_np)

    text = pytesseract.image_to_string(
        processed,
        config='--psm 6'
    )

    found = re.findall(r'\b\d{10}\b', text)

    extracted = []

    for number in found:

        if len(number) == 10:

            number_sum = calculate_sum(number)

            extracted.append({
                "Extracted Number": number,
                "Number Sum": number_sum
            })

    return extracted

# ---------------- UPLOAD ---------------- #

uploaded_files = st.file_uploader(
    " ",
    type=["jpg","jpeg","png","webp"],
    accept_multiple_files=True
)

if uploaded_files:

    if len(uploaded_files) > 30:

        st.error("Maximum 30 images allowed")

    else:

        all_rows = []
        serial = 1

        progress = st.progress(0)

        status_text = st.empty()

        for idx, file in enumerate(uploaded_files):

            status_text.markdown(
                f'<div class="processing-text">⚙ Processing Image {idx+1} of {len(uploaded_files)}...</div>',
                unsafe_allow_html=True
            )

            data = extract_numbers(file)

            for row in data:

                all_rows.append({
                    "S.No": serial,
                    "Extracted Number": row["Extracted Number"],
                    "Number Sum": row["Number Sum"]
                })

                serial += 1

            progress.progress((idx + 1) / len(uploaded_files))

        status_text.markdown(
            '<div class="processing-text">✅ Processing Completed</div>',
            unsafe_allow_html=True
        )

        if len(all_rows) == 0:

            st.warning("No valid 10-digit numbers detected")

        else:

            df = pd.DataFrame(all_rows)

            st.success(f"Successfully extracted {len(df)} numbers")

            st.dataframe(df, use_container_width=True)

            total_records = len(df)
            na_records = len(df[df["Number Sum"] == "NA"])
            valid_records = total_records - na_records

            st.markdown('<div class="summary-box">', unsafe_allow_html=True)

            st.markdown("### 📊 Summary")
            st.write(f"• Total Extracted Numbers: {total_records}")
            st.write(f"• Successfully Processed: {valid_records}")
            st.write(f"• NA Records: {na_records}")

            st.markdown('</div>', unsafe_allow_html=True)

            output = BytesIO()

            with pd.ExcelWriter(output, engine='openpyxl') as writer:

                df.to_excel(writer, index=False, sheet_name='ImaEx_Output')

                ws = writer.book['ImaEx_Output']

                header_fill = PatternFill(
                    start_color="1f2937",
                    end_color="1f2937",
                    fill_type="solid"
                )

                for cell in ws[1]:
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = header_fill

                for column_cells in ws.columns:
                    length = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)
                    ws.column_dimensions[column_cells[0].column_letter].width = length + 8

            st.download_button(
                label="⬇ Download Excel File",
                data=output.getvalue(),
                file_name="ImaEx_Output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

st.markdown('</div>', unsafe_allow_html=True)
