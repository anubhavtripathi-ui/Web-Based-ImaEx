import streamlit as st
import pandas as pd
from PIL import Image
import pytesseract
import re
from io import BytesIO

# ===================== CONFIG =====================
st.set_page_config(
    page_title="ImaEx - Image to Excel",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📸 Web-based ImaEx")
st.markdown("**10-Digit Numbers Extractor + Digital Root Calculator**")
st.markdown("---")

# Custom CSS for modern look
st.markdown("""
<style>
    .main {background-color: #f8f9fa;}
    .stButton>button {background-color: #4f46e5; color: white; border-radius: 8px; height: 3em;}
    .stButton>button:hover {background-color: #4338ca;}
    .uploadedFile {background-color: #f1f5f9;}
</style>
""", unsafe_allow_html=True)

# ===================== HELPER FUNCTIONS =====================
def digital_root(n):
    """Calculate digital root (repeated sum till single digit)"""
    if not str(n).isdigit():
        return "ANY"
    n = int(n)
    if n == 0:
        return 0
    return 1 + (n - 1) % 9

def extract_ten_digit_numbers(image):
    """Extract 10-digit numbers from image"""
    try:
        # Preprocessing for better OCR
        img = image.convert('L')  # Grayscale
        custom_config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789'
        
        text = pytesseract.image_to_string(img, config=custom_config)
        
        # Find all 10-digit numbers
        numbers = re.findall(r'\b\d{10}\b', text)
        return [num for num in numbers if len(num) == 10]
        
    except Exception as e:
        st.error(f"Error processing image: {str(e)}")
        return []

# ===================== SIDEBAR =====================
st.sidebar.header("⚙️ Settings")
st.sidebar.info("Maximum 30 images supported")

# ===================== MAIN UI =====================
uploaded_files = st.file_uploader(
    "📤 Upload Images (Multiple allowed)",
    type=["png", "jpg", "jpeg"],
    accept_multiple_files=True,
    help="Maximum 30 images"
)

if uploaded_files:
    if len(uploaded_files) > 30:
        st.error("Maximum 30 images allowed!")
    else:
        st.success(f"{len(uploaded_files)} images uploaded successfully!")

        if st.button("🚀 Process All Images", type="primary", use_container_width=True):
            with st.spinner("Processing images... This may take a minute"):
                all_numbers = []
                
                progress_bar = st.progress(0)
                
                for idx, uploaded_file in enumerate(uploaded_files):
                    image = Image.open(uploaded_file)
                    numbers = extract_ten_digit_numbers(image)
                    
                    for num in numbers:
                        all_numbers.append({
                            "Sr. No.": len(all_numbers) + 1,
                            "Number": num,
                            "Digital Root": digital_root(num)
                        })
                    
                    progress_bar.progress((idx + 1) / len(uploaded_files))
                
                # Create DataFrame
                df = pd.DataFrame(all_numbers)
                
                # Display results
                st.success(f"✅ Processing Complete! Extracted {len(df)} numbers")
                
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.dataframe(df, use_container_width=True, height=600)
                
                with col2:
                    st.metric("Total Numbers", len(df))
                
                # Download Button
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Extracted_Numbers')
                
                output.seek(0)
                
                st.download_button(
                    label="📥 Download Excel File",
                    data=output,
                    file_name="ImaEx_Extracted_Numbers.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )

st.caption("Made for Anubhav • Youth-friendly & Clean UI")
