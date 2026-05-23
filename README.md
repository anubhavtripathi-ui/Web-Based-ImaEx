# ⚡ ImaEx — Web-based Image to Excel

Upload images containing 10-digit numbers → auto-extract → download Excel.

## Deploy on Streamlit Cloud

### 1. Push to GitHub
```bash
git init
git add .
git commit -m "ImaEx v1"
git remote add origin https://github.com/YOUR_USERNAME/imaex.git
git push -u origin main
```

### 2. Deploy
- Go to [share.streamlit.io](https://share.streamlit.io)
- Connect your GitHub repo
- Set **Main file path**: `app.py`

### 3. Add API Key (Required)
In Streamlit Cloud → Your App → **Settings → Secrets**, add this:
```toml
ANTHROPIC_API_KEY = "sk-ant-api03-xxxxxxxxxxxx"
```
Save → App will restart automatically. Users won't see any API key field.

## Features
- Upload up to 30 images at once (JPG, PNG, WEBP)
- Claude Vision AI detects 10-digit numbers accurately
- Auto-calculates digit sum (reduces to single digit)
- Marks unclear digits as ANY
- Clean Excel download
