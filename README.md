# SegmentPPTCharts
# PPTX Chart Segmenter

This Streamlit app segments bar charts in a PowerPoint file into multiple series.

## How to Use

1. Upload your `.pptx` file.
2. Specify how many segments (bars) per chart.
3. Enter segment names.
4. Click "Process PPT" and download the updated file.

## Supported Chart Types

Currently supports **clustered bar charts**. Other chart types can be added upon request.

## Deployment

To deploy on Streamlit Cloud:

1. Push this repo to GitHub.
2. Go to [Streamlit Cloud](https://streamlit.io/cloud) and sign in with GitHub.
3. Click “New app” → select your repo → choose `main` branch and `app.py`.
4. Click **Deploy**.

Your live app link appears instantly. Share it with your team!

## Optional: Password Protection

To restrict access:
- Create `.streamlit/secrets.toml`:
  ```toml
  [admin]
  password = "YourPasswordHere"
