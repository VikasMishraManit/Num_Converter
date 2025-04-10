import streamlit as st
import pandas as pd
import tempfile
import os

# Function to strip leading zeros and convert to int
def strip_leading_zeros(val):
    if isinstance(val, str) and val.strip().lstrip("0").isdigit():
        return int(val.lstrip("0") or "0")
    return val

# Process Excel with pandas and save using xlsxwriter
def process_excel(file_path):
    writer_path = file_path.replace(".xlsx", "_cleaned.xlsx")
    xls = pd.read_excel(file_path, sheet_name=None)

    with pd.ExcelWriter(writer_path, engine='xlsxwriter') as writer:
        for sheet_name, df in xls.items():
            cleaned_df = df.applymap(strip_leading_zeros)
            cleaned_df.to_excel(writer, sheet_name=sheet_name, index=False)

    return writer_path

# Streamlit app
st.set_page_config(page_title="Excel Numeric Cleaner", layout="centered")
st.title("📊 Excel Numeric Auto-Cleaner")
st.markdown("Upload an `.xlsx` file and this app will convert all numeric-looking strings (even with leading zeros) into actual numbers, skipping date-like values.")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
        tmp_file.write(uploaded_file.read())
        temp_path = tmp_file.name

    st.success("File uploaded! Processing...")

    try:
        result_path = process_excel(temp_path)

        with open(result_path, "rb") as f:
            st.download_button(
                label="⬇️ Download Cleaned Excel File",
                data=f,
                file_name="cleaned_excel.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"❌ Error: {e}")

    # Clean up
    if os.path.exists(temp_path):
        os.remove(temp_path)
