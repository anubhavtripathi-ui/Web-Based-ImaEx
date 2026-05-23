
import streamlit as st
import pandas as pd
import cv2
import numpy as np
import re
from PIL import Image
from io import BytesIO
import easyocr
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from tempfile import NamedTemporaryFile

st.set_page_config(
    page_title="Web-based ImaEx",
    page_icon="📷",
    layout="wide"
)

# ---------------- UI ---------------- #
st.markdown("""
<style>
.main {
    background-color: #0f172a;
}
.stApp {
    background: linear-gradient(135deg, #0f172a, #111827);
    color: white;
}
.title {
    font-size: 42px;
    font-weight: 800;
    color: #f8fafc;
}
.subtitle {
    color: #cbd5e1;
    font-size: 18px;
    margin-bottom: 20px;
}
.card {
    background-color: #1e293b;
    padding: 20px;
    border-radius: 18px;
    margin-bottom: 20px;
}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="title">📷 Web-based ImaEx</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="subtitle">Upload up to 30 images and extract only 10-digit numbers into Excel.</div>',
    unsafe_allow_html=True
)

reader = easyocr.Reader(['en'], gpu=False)

# ---------------- Functions ---------------- #
def reduce_to_single_digit(num):
    while num > 9:
        num = sum(int(d) for d in str(num))
    return num

def process_image(image_file):
    image = Image.open(image_file).convert("RGB")
    img_np = np.array(image)

    gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)

    # Enhance image
    gray = cv2.GaussianBlur(gray, (3,3), 0)
    gray = cv2.adaptiveThreshold(
        gray, 255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        11,
        2
    )

    results = reader.readtext(gray, detail=1)

    extracted = []

    for result in results:
        text = result[1].strip()
        confidence = result[2]

        cleaned = re.sub(r'\D', '', text)

        if len(cleaned) == 10:
            if confidence < 0.70:
                extracted.append({
                    "Extracted Number": cleaned,
                    "Number Sum": "NA"
                })
            else:
                digit_sum = sum(int(x) for x in cleaned)
                final_sum = reduce_to_single_digit(digit_sum)

                # Double verification
                digit_sum_2 = sum(int(x) for x in cleaned)
                final_sum_2 = reduce_to_single_digit(digit_sum_2)

                if final_sum == final_sum_2:
                    extracted.append({
                        "Extracted Number": cleaned,
                        "Number Sum": final_sum
                    })
                else:
                    extracted.append({
                        "Extracted Number": cleaned,
                        "Number Sum": "NA"
                    })

    return extracted

# ---------------- Upload ---------------- #
uploaded_files = st.file_uploader(
    "Upload Images (Max 30)",
    type=["png", "jpg", "jpeg"],
    accept_multiple_files=True
)

if uploaded_files:

    if len(uploaded_files) > 30:
        st.error("❌ Maximum 30 images allowed.")
    else:

        all_data = []
        serial = 1

        with st.spinner("Processing images..."):
            for file in uploaded_files:
                data = process_image(file)

                for row in data:
                    all_data.append({
                        "S.No": serial,
                        "Extracted Number": row["Extracted Number"],
                        "Number Sum": row["Number Sum"]
                    })
                    serial += 1

        if len(all_data) == 0:
            st.warning("No valid 10-digit numbers detected.")
        else:

            df = pd.DataFrame(all_data)

            st.success(f"✅ Extraction completed. {len(df)} records found.")

            st.dataframe(df, use_container_width=True)

            # Excel creation
            output = BytesIO()

            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='ImaEx_Output')

                ws = writer.book['ImaEx_Output']

                for cell in ws[1]:
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill(
                        start_color="1E293B",
                        end_color="1E293B",
                        fill_type="solid"
                    )

                for column_cells in ws.columns:
                    length = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)
                    ws.column_dimensions[column_cells[0].column_letter].width = length + 5

            st.download_button(
                label="⬇ Download Excel",
                data=output.getvalue(),
                file_name="ImaEx_Output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
