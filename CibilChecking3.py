# Streamlit Duplicate Checker App

import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Duplicate Checker", layout="wide")


st.set_page_config(layout="wide")

hide_st_style = """
<style>

/* Hide all floating buttons */
button[data-testid="manage-app-button"] {
    visibility: hidden !important;
    display: none !important;
}

/* Hide parent container */
._terminalButton_rix23_138 {
    display: none !important;
}

/* Hide toolbar/header/footer */
header, footer, #MainMenu {
    visibility: hidden !important;
}

[data-testid="stToolbar"] {
    display: none !important;
}

/* Remove spacing */
.block-container {
    padding-top: 0rem;
}

</style>
"""
st.markdown(hide_st_style, unsafe_allow_html=True)

st.title("📊 Universal ID & PAN Duplicate Checker")
st.write("Upload Excel file and detect duplicate Universal IDs / PAN with different names.")

# =========================
# FILE UPLOAD
# =========================

uploaded_file = st.file_uploader(
    "Upload Excel File",
    type=["xlsx", "xls"]
)
if uploaded_file is not None:
    # =========================
    # READ FILE
    # =========================
    df = pd.read_excel(uploaded_file)
    st.success("File Uploaded Successfully ✅")
    st.subheader("Raw Data Preview")
    st.dataframe(df.head())
    # =========================
    # ACTUAL FILE COLUMNS
    # =========================
    uploaded_columns = list(df.columns)

# =========================
# REQUIRED COLUMN NAMES
# =========================

    required_columns = [
        "Universal ID Number",
        "Income Tax Id Number",
        "Consumer Name"
    ]

# Show Sample Format
    st.info("""
        ✅ Correct Header Format Should Be:

        1. Universal ID Number
        2. Income Tax Id Number
        3. Consumer Name
        """)
# =========================
# CHECK MISSING COLUMNS
# =========================

    missing_columns = [
        col for col in required_columns
        if col not in uploaded_columns
    ]

    # =========================
    # SHOW ERROR IF MISSING
    # =========================

    if missing_columns:

        st.error("❌ Required Column Names Not Found")

        st.warning(
            "Please check your Excel headers carefully."
        )

        # Show Missing Columns
        st.write("### Missing Columns:")
        st.write(missing_columns)

        # Show Uploaded Columns
        st.write("### Your Uploaded File Headers:")
        st.write(uploaded_columns)

        

        st.stop()

    # =========================
    # COLUMN NAMES
    # =========================

    uid_col = "Universal ID Number"
    pan_col = "Income Tax Id Number"
    name_col = "Consumer Name"

    # =========================
    # CLEAN DATA
    # =========================

    df[uid_col] = df[uid_col].astype("string").str.strip()
    df[pan_col] = df[pan_col].astype("string").str.strip().str.upper()
    df[name_col] = df[name_col].astype("string").str.strip()

    # Remove blank UID values
    uid_df = df[
        df[uid_col].notna() &
        (df[uid_col] != "") &
        (df[uid_col].str.lower() != "nan")
    ].copy()

    # Remove blank PAN values
    pan_df = df[
        df[pan_col].notna() &
        (df[pan_col] != "") &
        (df[pan_col].str.lower() != "nan")
    ].copy()

    # =========================
    # UNIVERSAL ID CHECK
    # =========================

    uid_name_count = (
        uid_df.groupby(uid_col)[name_col]
        .nunique()
        .reset_index(name="Duplicate Count")
    )

    uid_multiple_names = uid_name_count[
        uid_name_count["Duplicate Count"] > 1
    ][uid_col]

    uid_duplicate = uid_df[
        uid_df[uid_col].isin(uid_multiple_names)
    ].copy()

    uid_duplicate = uid_duplicate.merge(
        uid_name_count,
        on=uid_col,
        how="left"
    )

    uid_duplicate = uid_duplicate[
        [uid_col, name_col, "Duplicate Count"]
    ].drop_duplicates().sort_values(by=uid_col)

    # =========================
    # PAN CHECK
    # =========================

    pan_name_count = (
        pan_df.groupby(pan_col)[name_col]
        .nunique()
        .reset_index(name="Duplicate Count")
    )

    pan_multiple_names = pan_name_count[
        pan_name_count["Duplicate Count"] > 1
    ][pan_col]

    pan_duplicate = pan_df[
        pan_df[pan_col].isin(pan_multiple_names)
    ].copy()

    pan_duplicate = pan_duplicate.merge(
        pan_name_count,
        on=pan_col,
        how="left"
    )

    pan_duplicate = pan_duplicate[
        [pan_col, name_col, "Duplicate Count"]
    ].drop_duplicates().sort_values(by=pan_col)

    # =========================
    # PREVIEW RESULTS
    # =========================

    st.subheader("🔍 Universal ID Issues")
    st.write(f"Total Records: {len(uid_duplicate)}")
    st.dataframe(uid_duplicate)

    st.subheader("🔍 PAN Issues")
    st.write(f"Total Records: {len(pan_duplicate)}")
    st.dataframe(pan_duplicate)

    # =========================
    # CREATE EXCEL DOWNLOAD
    # =========================

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        uid_duplicate.to_excel(
            writer,
            sheet_name="Universal ID Issues",
            index=False
        )

        pan_duplicate.to_excel(
            writer,
            sheet_name="PAN Issues",
            index=False
        )

    output.seek(0)

    # =========================
    # DOWNLOAD BUTTON
    # =========================

    st.download_button(
        label="⬇ Download Duplicate Report",
        data=output,
        file_name="Actual_Duplicate_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
