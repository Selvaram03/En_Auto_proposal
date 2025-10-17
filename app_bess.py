# app_bess.py
import streamlit as st
import os
from shared_app import load_excel, generate_word, download_docx, download_pdf

def run_bess():
    st.header("🔋 BESS Proposal Generator")
    TEMPLATE_PATH = "BESS_Template.docx"
    TEMPLATE_EXCEL_PATH = "Input_BESS_Proposal.xlsx"

    # Initialize session state for BESS
    if 'bess_uploaded' not in st.session_state:
        st.session_state.bess_uploaded = None
    if 'bess_df' not in st.session_state:
        st.session_state.bess_df = None
    if 'bess_doc' not in st.session_state:
        st.session_state.bess_doc = None

    # Clear previous BESS data if EPC was previously used
    if st.session_state.get('current_page') != 'bess':
        st.session_state.bess_uploaded = None
        st.session_state.bess_df = None
        st.session_state.bess_doc = None
        st.session_state.current_page = 'bess'

    # Download Excel template
    try:
        with open(TEMPLATE_EXCEL_PATH, "rb") as f:
            st.download_button("📥 Download BESS Excel Template",
                               f, os.path.basename(TEMPLATE_EXCEL_PATH),
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except FileNotFoundError:
        st.warning("⚠️ BESS Excel template not found.")

    # Upload Excel
    uploaded = st.file_uploader("📤 Upload BESS Excel", type=["xlsx"], key="bess_uploader")
    if uploaded:
        st.session_state.bess_uploaded = uploaded
        df = load_excel(uploaded)
        if df is not None:
            st.session_state.bess_df = df

    # Show table
    if st.session_state.bess_df is not None:
        st.dataframe(st.session_state.bess_df)

        if st.button("🚀 Generate BESS Proposal"):
            doc = generate_word(TEMPLATE_PATH, st.session_state.bess_df)
            st.session_state.bess_doc = doc
            download_docx(doc, "BESS_Proposal")
            download_pdf(doc, "BESS_Proposal")
