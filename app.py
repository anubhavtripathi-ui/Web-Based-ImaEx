import streamlit as st
import pandas as pd
import cv2
import numpy as np
import re
from PIL import Image
from io import BytesIO
from rapidocr_onnxruntime import RapidOCR
from openpyxl.styles import Font, PatternFill

st.set_page_config(
    page_title="ImaEx",
    page_icon="⚡",
    layout="centered",
    initial_sidebar_state="collapsed"
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

.block-container{
    padding-top:1rem;
}

.main-box{
    background:white;
    border-radius:24px;
    padding:40px;
    box-shadow:0 6px 24px rgba(0,0,0,0.05);
}

.logo-row{
    display:flex;
    justify-content:center;
    align-items:center;
    gap:12px;
}

.logo{
    font-size:48px;
}

.title{
    font-size:58px;
    font-weight:800;
    color:#1f2937;
}

.subtitle{
    text-align:center;
    color:#6b7280;
    font-size:20px;
    margin-top:10px;
    margin-bottom:35px;
}

.stFileUploader{
    background:#f8fafc;
    border:2px dashed #cbd5e1;
    border-radius:18px;
    padding:18px;
}

.processing{
    color:#111827;
    font-size:18px;
    font-weight:700;
}

.summary{
    background:#f8fafc;
    border:1px solid #e5e7eb;
    border-radius:16px;
    padding:18px;
    margin-top:20px;
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
<div class="logo-row">
<div class="logo">⚡</div>
<div class="title">ImaEx</div>
</div>
""", unsafe_allow_html=True)

st.markdown(
    '<div class="subtitle">Image → Extract 10-digit numbers → Excel</div>',
    unsafe_allow_html=True
)

st.markdown("### 📁 Upload Images")
st.caption("Upload up to 30 JPG / PNG / WEBP images")

# ---------------- OCR ---------------- #

reader = RapidOCR()

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

def preprocess(img_np):

    gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)

    gray = cv2.resize(
        gray,
        None,
        fx=2,
        fy=2,
        interpolation=cv2.INTER_CUBIC
    )

    gray = cv2.fastNlMeansDenoising(gray)

    kernel = np.array([
        [0,-1,0],
        [-1,5,-1],
        [0,-1,0]
    ])

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

def extract_numbers(file):

    image = Image.open(file).convert("RGB")

    img_np = np.array(image)

    processed = preprocess(img_np)

    results, _ = reader(processed)

    extracted = []

    if results:

        for item in results:

            try:
                text = item[1]

                cleaned = re.sub(r'\D', '', text)

                if len(cleaned) == 10:

                    number_sum = calculate_sum(cleaned)

                    extracted.append({
                        "Extracted Number": cleaned,
                        "Number Sum": number_sum
                    })

            except:
                pass

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

        progress = st.progress(0)

        status = st.empty()

        all_rows = []

        serial = 1

        for idx, file in enumerate(uploaded_files):

            status.markdown(
                f'<div class="processing">⚙ Processing image {idx+1} of {len(uploaded_files)}...</div>',
                unsafe_allow_html=True
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

        status.markdown(
            '<div class="processing">✅ Processing Completed</div>',
            unsafe_allow_html=True
        )

        if len(all_rows) == 0:

            st.warning("No valid 10-digit numbers detected")

        else:

            df = pd.DataFrame(all_rows)

            st.success(f"Successfully extracted {len(df)} numbers")

            st.dataframe(df, use_container_width=True)

            st.markdown('<div class="summary">', unsafe_allow_html=True)

            st.markdown("### 📊 Summary")

            st.write(f"• Total Numbers: {len(df)}")
            st.write(f"• Successfully Processed: {len(df[df['Number Sum'] != 'NA'])}")
            st.write(f"• NA Records: {len(df[df['Number Sum'] == 'NA'])}")

            st.markdown('</div>', unsafe_allow_html=True)

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

st.markdown('</div>', unsafe_allow_html=True)
