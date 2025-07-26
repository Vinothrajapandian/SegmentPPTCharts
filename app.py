import streamlit as st
from utils import process_pptx
import tempfile
import os

st.set_page_config(page_title="PPTX Segmenter", layout="centered")
st.title("📊 PPTX Chart Segment Segmenter")

uploaded = st.file_uploader("Upload your PPTX file", type=["pptx"])
seg_count = st.number_input("Number of segments per chart", min_value=1, max_value=10, value=3)

segment_names = []
if seg_count:
    st.markdown("### Enter segment names:")
    for i in range(int(seg_count)):
        segment_names.append(st.text_input(f"Segment {i+1}", value=f"Segment {i+1}"))

if uploaded and st.button("Process PPT"):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        tmp.write(uploaded.read())
        tmp.flush()
        output_path = process_pptx(tmp.name, int(seg_count), segment_names)

    with open(output_path, "rb") as f:
        st.download_button("Download Updated PPTX", f, file_name="segmented_output.pptx")
    os.remove(output_path)
