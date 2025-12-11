import streamlit as st
import pandas as pd
import io

def extract_columns_from_file(uploaded_file):
    # Read the file, skipping the first 5 rows so row 6 is the first row
    df = pd.read_excel(uploaded_file, header=None, skiprows=5)

    # Extract columns B, G, and H (1, 6, 7 in zero-based indexing)
    extracted = df.iloc[:, [1, 6, 7, 18, 25]]

    # Drop rows where all values are empty
    extracted = extracted.dropna(how="all")

    # Rename columns
    extracted.columns = ["Item Name", "MFR SKU", "QTY", "Tracking to Vivo", "Tracking to Site"]

    return extracted

st.title("VIVO BOM Extractor")

uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    st.success("File uploaded successfully!")

    df = extract_columns_from_file(uploaded_file)
    st.dataframe(df)

    # Create downloadable Excel in memory
    output = io.BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    st.download_button(
        label="Download Extracted Excel",
        data=output,
        file_name="extracted_BOM.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


