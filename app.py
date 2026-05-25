import streamlit as st
import pandas as pd
import pytesseract
import numpy as np
from PIL import Image
import io

# Page Configuration
st.set_page_config(page_title="AI Mobile Number Extractor", page_icon="📱", layout="wide")

# App Title
st.title("📱 AI Vertical Column Mobile Number Extractor")
st.write("Upload your images (maximum 30). This tool automatically detects vertical columns and extracts mobile numbers without requiring a fixed grid or layout pattern.")

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
    # Convert uploaded file to PIL Image
    image = Image.open(image_file).convert('RGB')
    img_width, _ = image.size

    # Run Tesseract OCR with detailed structural data layout
    try:
        data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DICT)
    except Exception as e:
        st.error(f"OCR Engine Error: Please ensure 'packages.txt' is added to GitHub. Details: {e}")
        return []
    
    blocks = []
    n_boxes = len(data['text'])
    
    for i in range(n_boxes):
        text = data['text'][i].strip()
        # Clean text to retain only digits
        clean_digits = "".join([c for c in text if c.isdigit()])
        
        # Filter out tiny standalone noise numbers
        if len(clean_digits) >= 5:  
            # Extract coordinates to find centers
            x = data['left'][i]
            y = data['top'][i]
            w = data['width'][i]
            h = data['height'][i]
            
            cx = x + w / 2
            cy = y + h / 2
            blocks.append({'cx': cx, 'cy': cy, 'digits': clean_digits})
            
    if not blocks:
        return []

    # Dynamic Column Clustering Heuristic
    # Sort blocks primarily by X coordinate to map out separate columns
    blocks.sort(key=lambda b: b['cx'])
    columns = []
    current_col = [blocks[0]]
    
    # 10% of image width used as a threshold boundary to distinguish columns
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
                extracted_data.append({"Mobile Number": num, "Sum": digit_sum})
            else:
                # Logic for unclear, incomplete, or broken numbers
                extracted_data.append({"Mobile Number": num, "Sum": "NA"})
                
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
                # Build final sequential data format matching user template
                final_structured_data = []
                for idx, record in enumerate(all_records, start=1):
                    final_structured_data.append({
                        "S No": idx,
                        "Mobile Number": record["Mobile Number"],
                        "Sum": record["Sum"]
                    })
                
                df = pd.DataFrame(final_structured_data)
                
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
