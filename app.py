import streamlit as st
import io
import sys

from backend_logic import (
    find_duplicates_one_to_many,
    load_excel_master_dataframe,
    extract_filtered_excel_inputs
)

st.set_page_config(
    page_title="Duplicates Identification Tool",
    layout="wide"
)

st.title("Duplicates Identification Tool")

st.markdown("---")

# -----------------------------------------------------------
# MODE SELECTION
# -----------------------------------------------------------

mode = st.radio(
    "Choose Input Method",
    [
        "Manual User Input",
        "Excel File Extraction"
    ]
)

st.markdown("---")

# ===========================================================
# MANUAL INPUT MODE
# ===========================================================

if mode == "Manual User Input":

    st.subheader("Manual User Input")

    col1, col2 = st.columns(2)

    with col1:
        new_product_number = st.text_input(
            "New Product Number (e.g 9808812280)"
        )

    with col2:
        new_ecdv = st.text_input(
            "New Part ECDV (e.g EN.1GUO.XD03.DC02/XD03.DC03/DC02*)"
        )

    st.markdown("### Existing Parts")

    col3, col4 = st.columns(2)

    with col3:
        existing_products_text = st.text_area(
            "Existing Product Numbers\n(One per line)",
            height=200
        )

    with col4:
        existing_ecdv_text = st.text_area(
            "Existing Part ECDVs\n(One per line)",
            height=200
        )

    st.info(
        "Note:\n"
        "- Ensure NFC date filtering is done manually.\n"
        "- Ensure all parts belong to same code function."
    )

    if st.button("Check Duplicate"):

        other_product_numbers = [
            x.strip()
            for x in existing_products_text.split("\n")
            if x.strip()
        ]

        other_ecdvs = [
            x.strip()
            for x in existing_ecdv_text.split("\n")
            if x.strip()
        ]

        buffer = io.StringIO()
        sys.stdout = buffer

        try:
            find_duplicates_one_to_many(
                new_ecdv,
                other_ecdvs,
                new_product_number,
                other_product_numbers
            )
            output = buffer.getvalue()

        except Exception as e:
            output = str(e)

        finally:
            sys.stdout = sys.__stdout__

        st.markdown("### Output")

        st.code(output, language="text")


# ===========================================================
# EXCEL INPUT MODE
# ===========================================================

if mode == "Excel File Extraction":

    st.subheader("Excel File Extraction Input")

    col1, col2 = st.columns(2)

    with col1:
        new_product_number = st.text_input("New Product Number")

    with col2:
        new_ecdv = st.text_input("New Part ECDV")

    col3, col4 = st.columns(2)

    with col3:
        code_function = st.text_input("Code Function (e.g E1101001)")

    with col4:
        new_product_NFCdate = st.text_input(
            "New Product NFC Date (YYYY-MM-DD)"
        )

    file_path = st.text_input("Excel File Path")

    if st.button("Check Duplicate"):

        buffer = io.StringIO()
        sys.stdout = buffer

        try:
            df_master = load_excel_master_dataframe(file_path)

            other_product_numbers, other_ecdvs = extract_filtered_excel_inputs(
                df_master=df_master,
                code_function=code_function,
                new_product_NFCdate=new_product_NFCdate
            )

            find_duplicates_one_to_many(
                new_ecdv,
                other_ecdvs,
                new_product_number,
                other_product_numbers
            )

            output = buffer.getvalue()

        except Exception as e:
            output = str(e)

        finally:
            sys.stdout = sys.__stdout__

        st.markdown("### Output")

        st.code(output, language="text")