# ⚡ ImaEx — Web-based Image to Excel

Upload images containing 10-digit numbers → auto-extract → download Excel.

## Features
- Upload up to 30 images at once (JPG, PNG, WEBP)
- Claude Vision AI detects 10-digit numbers accurately
- Auto-calculates digit sum (reduces to single digit)
- Marks unclear digits as **ANY**
- Clean Excel download with Sr. No., Number, Digit Sum columns

## Deploy on Streamlit Cloud

### 1. Push to GitHub
```bash
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/YOUR_USERNAME/imaex.git
git push -u origin main
```

### 2. Deploy
- Go to [share.streamlit.io](https://share.streamlit.io)
- Connect your GitHub repo
- Set **Main file path**: `app.py`

### 3. Add API Key (Secrets)
In Streamlit Cloud → App settings → **Secrets**, add:
```toml
# Optional: if you want to pre-fill the API key
# ANTHROPIC_API_KEY = "sk-ant-..."
```
Or simply enter the key in the UI when using the app.

## Local Run
```bash
pip install -r requirements.txt
streamlit run app.py
```
