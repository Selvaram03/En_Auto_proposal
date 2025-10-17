# app_epc.py
import streamlit as st
import os
from shared_app import load_excel, generate_word, download_docx, download_pdf

def run_epc():
    st.header("📄 EPC Proposal Generator")
    TEMPLATE_PATH = "EPC_Template.docx"
    TEMPLATE_EXCEL_PATH = "Input_EPC_Proposal.xlsx"

    # Initialize session state for EPC
    if 'epc_uploaded' not in st.session_state:
        st.session_state.epc_uploaded = None
    if 'epc_df' not in st.session_state:
        st.session_state.epc_df = None
    if 'epc_doc' not in st.session_state:
        st.session_state.epc_doc = None

    # Clear previous EPC data if BESS was previously used
    if st.session_state.get('current_page') != 'epc':
        st.session_state.epc_uploaded = None
        st.session_state.epc_df = None
        st.session_state.epc_doc = None
        st.session_state.current_page = 'epc'

    # Download Excel template
    try:
        with open(TEMPLATE_EXCEL_PATH, "rb") as f:
            st.download_button("📥 Download EPC Excel Template",
                               f, os.path.basename(TEMPLATE_EXCEL_PATH),
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except FileNotFoundError:
        st.warning("⚠️ EPC Excel template not found.")

    # Upload Excel
    uploaded = st.file_uploader("📤 Upload EPC Excel", type=["xlsx"], key="epc_uploader")
    if uploaded:
        st.session_state.epc_uploaded = uploaded
        df = load_excel(uploaded)
        if df is not None:
            st.session_state.epc_df = df

    # Show table
    if st.session_state.epc_df is not None:
        st.dataframe(st.session_state.epc_df)

        if st.button("🚀 Generate EPC Proposal"):
            doc = generate_word(TEMPLATE_PATH, st.session_state.epc_df)
            st.session_state.epc_doc = doc
            download_docx(doc, "EPC_Proposal")
            download_pdf(doc, "EPC_Proposal")
