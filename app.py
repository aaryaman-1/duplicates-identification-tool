import streamlit as st
import pandas as pd
import io
from contextlib import redirect_stdout

# Import backend functions (NO LOGIC CHANGES)
from backend_logic import (
    load_excel_master_dataframe,
    extract_filtered_excel_inputs,
    find_duplicates_one_to_many
)

st.set_page_config(
    page_title="Duplicates Identification Tool",
    layout="centered"
)

st.title("Duplicates Identification Tool")


# ---------------------------------------------------
# CACHE EXCEL LOADER (Load once per day)
# ---------------------------------------------------
@st.cache_data(ttl=86400)
def cached_load_excel(uploaded_file):
    return load_excel_master_dataframe(uploaded_file)


# ---------------------------------------------------
# MODE SELECTION
# ---------------------------------------------------
mode = st.radio(
    "Select Input Method",
    [
        "Manual User Input",
        "Excel Extraction Input"
    ]
)


# ===================================================
# MANUAL INPUT MODE
# ===================================================
if mode == "Manual User Input":

    st.subheader("Manual Input Mode")

    col1, col2 = st.columns(2)

    with col1:
        new_product_number = st.text_input(
            "New Product Number"
        )

    with col2:
        new_ecdv = st.text_input(
            "New Product ECDV"
        )

    st.markdown("### Existing Parts")

    col3, col4 = st.columns(2)

    with col3:
        existing_products_text = st.text_area(
            "Existing Product Numbers (one per line)",
            height=200
        )

    with col4:
        existing_ecdvs_text = st.text_area(
            "Existing Product ECDVs (one per line)",
            height=200
        )

    st.info(
        """
Manual Mode Notes:
- Date filtering must be done manually.
- Only paste parts belonging to the same code function.
"""
    )

    if st.button("Check Duplicate"):

        other_product_numbers = [
            x.strip()
            for x in existing_products_text.split("\n")
            if x.strip()
        ]

        other_ecdvs = [
            x.strip()
            for x in existing_ecdvs_text.split("\n")
            if x.strip()
        ]

        buffer = io.StringIO()

        with redirect_stdout(buffer):

            find_duplicates_one_to_many(
                new_ecdv,
                other_ecdvs,
                new_product_number,
                other_product_numbers
            )

        output = buffer.getvalue()

        st.markdown("### Output")

        st.code(output)



# ===================================================
# EXCEL INPUT MODE
# ===================================================
elif mode == "Excel Extraction Input":

    st.subheader("Excel Extraction Mode")

    new_product_number = st.text_input(
        "New Product Number"
    )

    new_ecdv = st.text_input(
        "New Product ECDV"
    )

    code_function = st.text_input(
        "Code Function"
    )

    new_product_NFCdate = st.text_input(
        "New Product NFC Date (YYYY-MM-DD)"
    )

    uploaded_file = st.file_uploader(
        "Upload MBOM Excel File",
        type=["xlsx"]
    )

    if st.button("Check Duplicate"):

        if uploaded_file is None:
            st.error("Please upload Excel file.")
        else:

            df_master = cached_load_excel(uploaded_file)

            other_product_numbers, other_ecdvs = extract_filtered_excel_inputs(
                df_master=df_master,
                code_function=code_function,
                new_product_NFCdate=new_product_NFCdate
            )

            buffer = io.StringIO()

            with redirect_stdout(buffer):

                find_duplicates_one_to_many(
                    new_ecdv,
                    other_ecdvs,
                    new_product_number,
                    other_product_numbers
                )

            output = buffer.getvalue()

            st.markdown("### Output")

            st.code(output)
