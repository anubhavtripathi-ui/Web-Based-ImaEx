
import streamlit as st
import pandas as pd
import cv2
import numpy as np
import re
from PIL import Image
from io import BytesIO
import easyocr
from openpyxl.styles import Font, PatternFill

st.set_page_config(
    page_title="ImaEx",
    page_icon="⚡",
    layout="centered"
)

# ---------------- UI ---------------- #
st.markdown("""
<style>

html, body, [class*="css"] {
    font-family: 'Segoe UI', sans-serif;
}

.stApp {
    background: #eef2f7;
}

.main-container {
    background: white;
    border-radius: 28px;
    padding: 45px;
    margin-top: 20px;
    box-shadow: 0 8px 30px rgba(0,0,0,0.06);
}

.logo {
    text-align:center;
    font-size:60px;
    margin-bottom:0px;
}

.title {
    text-align:center;
    font-size:58px;
    font-weight:800;
    color:#1f2937;
    margin-top:-10px;
}

.subtitle {
    text-align:center;
    color:#6b7280;
    font-size:22px;
    margin-top:5px;
    margin-bottom:35px;
}

.upload-box {
    background:#f8fafc;
    border:2px dashed #d1d5db;
    border-radius:20px;
    padding:25px;
}

.summary-card {
    background:#f8fafc;
    border-radius:18px;
    padding:20px;
    margin-top:20px;
    border:1px solid #e5e7eb;
}

.stDownloadButton > button {
    width:100%;
    background: linear-gradient(90deg, #ff9966, #ff5e62);
    color:white;
    border:none;
    padding:16px;
    border-radius:14px;
    font-size:18px;
    font-weight:700;
}

.stFileUploader {
    background:#f8fafc;
    padding:20px;
    border-radius:20px;
    border:2px dashed #cbd5e1;
}

[data-testid="stFileUploaderDropzone"] {
    background:#f8fafc;
    border:none;
    padding:25px;
}

</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-container">', unsafe_allow_html=True)

st.markdown('<div class="logo">⚡</div>', unsafe_allow_html=True)
st.markdown('<div class="title">ImaEx</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="subtitle">Image → Extract 10-digit numbers → Excel</div>',
    unsafe_allow_html=True
)

st.markdown("### 📁 Upload Images")
st.caption("Upload up to 30 JPG / PNG / WEBP images")

reader = easyocr.Reader(['en'], gpu=False)

# ---------------- FUNCTIONS ---------------- #

def reduce_to_single_digit(num):
    while num > 9:
        num = sum(int(d) for d in str(num))
    return num

def calculate_sum(number):

    try:
        total_1 = sum(int(x) for x in number)
        final_1 = reduce_to_single_digit(total_1)

        total_2 = sum(int(x) for x in number)
        final_2 = reduce_to_single_digit(total_2)

        if final_1 == final_2:
            return final_1

        return "NA"

    except:
        return "NA"

def preprocess_image(img_np):

    gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)

    # upscale
    gray = cv2.resize(gray, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)

    # denoise
    gray = cv2.fastNlMeansDenoising(gray)

    # sharpen
    kernel = np.array([[0,-1,0],[-1,5,-1],[0,-1,0]])
    gray = cv2.filter2D(gray, -1, kernel)

    # adaptive threshold
    thresh = cv2.adaptiveThreshold(
        gray,
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

    results = reader.readtext(
        processed,
        detail=1,
        paragraph=False
    )

    extracted = []

    for item in results:

        text = item[1].strip()
        confidence = item[2]

        cleaned = re.sub(r'\D', '', text)

        # only 10 digit focus
        if len(cleaned) == 10:

            if confidence < 0.58:
                number_sum = "NA"
            else:
                number_sum = calculate_sum(cleaned)

            extracted.append({
                "Extracted Number": cleaned,
                "Number Sum": number_sum
            })

    return extracted

# ---------------- UPLOADER ---------------- #

uploaded_files = st.file_uploader(
    " ",
    type=["jpg", "jpeg", "png", "webp"],
    accept_multiple_files=True
)

if uploaded_files:

    if len(uploaded_files) > 30:

        st.error("Maximum 30 images allowed.")

    else:

        all_rows = []
        serial = 1

        progress = st.progress(0)

        for idx, file in enumerate(uploaded_files):

            data = extract_numbers(file)

            for row in data:

                all_rows.append({
                    "S.No": serial,
                    "Extracted Number": row["Extracted Number"],
                    "Number Sum": row["Number Sum"]
                })

                serial += 1

            progress.progress((idx + 1) / len(uploaded_files))

        if len(all_rows) == 0:

            st.warning("No valid 10-digit numbers detected.")

        else:

            df = pd.DataFrame(all_rows)

            st.success(f"✅ Extraction completed successfully — {len(df)} numbers found")

            st.dataframe(df, use_container_width=True)

            # summary
            total_records = len(df)
            valid_records = len(df[df["Number Sum"] != "NA"])
            na_records = len(df[df["Number Sum"] == "NA"])

            st.markdown('<div class="summary-card">', unsafe_allow_html=True)

            st.markdown("### 📊 Summary")
            st.write(f"• Total Numbers Extracted: {total_records}")
            st.write(f"• Successfully Processed: {valid_records}")
            st.write(f"• NA / Low Confidence: {na_records}")

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
