"""
Charging Session Report Generator - Web App
--------------------------------------------
A simple browser-based version of failed_session_analyzer.py.
Upload a CSV/Excel file, pick a failure cutoff, and download the
generated Excel + PDF reports - no command line needed.

This file does NOT change failed_session_analyzer.py at all - it just
imports and reuses the analyze() function from it.
"""

import os
import tempfile
from pathlib import Path

import streamlit as st

from failed_session_analyzer import analyze

st.set_page_config(page_title="Charging Session Report Generator", page_icon="⚡", layout="centered")

st.title("⚡ Charging Session Report Generator")
st.write(
    "Upload your charging sessions file (CSV or Excel). "
    "The tool will automatically detect failed sessions and generate "
    "an Excel + PDF report you can download."
)

uploaded_file = st.file_uploader("Upload Charging Sessions file", type=["csv", "xlsx", "xls"])
minutes = st.number_input("Failure cutoff (minutes)", min_value=0.5, value=2.0, step=0.5)
output_name = st.text_input("Report file name (without extension)", value="Charging Session Report")

if uploaded_file is not None:
    if st.button("Generate Report", type="primary"):
        with st.spinner("Analyzing sessions and building your report..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                # Save the uploaded file to a temp path
                input_path = os.path.join(tmpdir, uploaded_file.name)
                with open(input_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                output_xlsx = os.path.join(tmpdir, f"{output_name}.xlsx")

                try:
                    analyze(input_path, output_path=output_xlsx, cutoff_minutes=minutes)
                except Exception as e:
                    st.error(f"Something went wrong while processing the file: {e}")
                    st.stop()

                pdf_path = str(Path(output_xlsx).with_suffix(".pdf"))

                st.success("Report generated successfully!")

                with open(output_xlsx, "rb") as f:
                    st.download_button(
                        "📊 Download Excel Report",
                        f.read(),
                        file_name=f"{output_name}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                with open(pdf_path, "rb") as f:
                    st.download_button(
                        "📄 Download PDF Report",
                        f.read(),
                        file_name=f"{output_name}.pdf",
                        mime="application/pdf",
                    )

st.caption("This is a system generated report - Charging Session Analyzer.")
