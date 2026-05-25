import streamlit as st
import pandas as pd
import easyocr
import numpy as np
from PIL import Image
import io

# Page Configuration
st.set_page_config(page_title="AI Mobile Number Extractor", page_icon="📱", layout="wide")

# App Title
st.title("📱 AI Vertical Column Mobile Number Extractor")
st.write("Upload your images (maximum 30). This tool automatically detects vertical columns and extracts 10-digit mobile numbers without requiring a fixed grid or layout pattern.")

# Cache the OCR Reader so it doesn't reload on every interaction
@st.cache_resource
def load_ocr_reader():
    # 'en' language works best for standard digital fonts
    return easyocr.Reader(['en'], gpu=False)

reader = load_ocr_reader()

def calculate_digital_root(number_str):
    """Calculates the single-digit sum (Digital Root) of a 10-digit number."""
    if len(number_str) != 10 or not number_str.isdigit():
        return "NA"
    
    total_sum = sum(int(digit) for digit in number_str)
    while total_sum > 9:
        total_sum = sum(int(digit) for digit in str(total_sum))
    return total_sum

def process_image(image_file):
    """Processes the image, groups numbers into vertical columns, and extracts data."""
    # Convert uploaded file to PIL Image and then to a numpy array
    image = Image.open(image_file).convert('RGB')
    image_np = np.array(image)
    img_width, _ = image.size

    # Run EasyOCR
    results = reader.readtext(image_np)
    
    blocks = []
    for res in results:
        bbox, text, prob = res
        # Calculate center X and Y coordinates of the text block
        cx = sum([p[0] for p in bbox]) / 4
        cy = sum([p[1] for p in bbox]) / 4
        
        # Clean text to retain only digits
        clean_digits = "".join([c for c in text if c.isdigit()])
        
        if len(clean_digits) >= 5:  # Filter out small irrelevant noise/numbers
            blocks.append({'cx': cx, 'cy': cy, 'digits': clean_digits})
            
    if not blocks:
        return []

    # Dynamic Column Clustering Heuristic
    # Sort blocks primarily by X coordinate to map out columns
    blocks.sort(key=lambda b: b['cx'])
    columns = []
    current_col = [blocks[0]]
    
    # 10% of image width is used as a threshold to group elements into the same column
    col_threshold = img_width * 0.10 

    for b in blocks[1:]:
        avg_cx = sum([x['cx'] for x in current_col]) / len(current_col)
        if abs(b['cx'] - avg_cx) < col_threshold:
            current_col.append(b)
        else:
            columns.append(current_col)
            current_col = [b]
    columns.append(current_col)

    # Sort each individual column from top to bottom (Y coordinate)
    extracted_data = []
    for col in columns:
        col.sort(key=lambda b: b['cy'])
        for b in col:
            num = b['digits']
            if len(num) == 10:
                digit_sum = calculate_digital_root(num)
                extracted_data.append({"Mobile Number": num, "Single Digit Sum": digit_sum})
            else:
                # Logic for unclear, incomplete, or broken numbers
                extracted_data.append({"Mobile Number": num, "Single Digit Sum": "NA"})
                
    return extracted_data

# File Uploader (Restricted to 30 files max)
uploaded_files = st.file_uploader("Upload Images (Max 30)", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

if uploaded_files:
    if len(uploaded_files) > 30:
        st.error("Please upload a maximum of 30 images at a time.")
    else:
        st.success(f"Successfully uploaded {len(uploaded_files)} image(s)!")
        
        if st.button("Start AI Extraction 🚀"):
            all_records = []
            
            # Progress bar setup
            progress_bar = st.progress(0)
            
            for index, file in enumerate(uploaded_files):
                st.write(f"Processing: **{file.name}**...")
                file_data = process_image(file)
                all_records.extend(file_data)
                
                # Update progress bar
                progress_bar.progress((index + 1) / len(uploaded_files))
                
            if all_records:
                # Create DataFrame
                df = pd.DataFrame(all_records)
                
                st.write("---")
                st.subheader("📊 Extracted Data Preview")
                st.dataframe(df, use_container_width=True)
                
                # Convert DataFrame to Excel format in-memory
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer_excel:
                    df.to_excel(writer_excel, index=False, sheet_name='Mobile Numbers')
                processed_data = output.getvalue()
                
                # Download Button
                st.write("---")
                st.download_button(
                    label="📥 Download Excel File",
                    data=processed_data,
                    file_name="Extracted_Mobile_Numbers.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No valid numbers could be extracted from the uploaded images.")
