# main.py
import streamlit as st
from app_epc import run_epc
from app_bess import run_bess

st.set_page_config(page_title="Proposal Generator", layout="wide")

# Sidebar logo
logo_path = "enrich_logo.png"
try:
    st.sidebar.image(logo_path, width=150)
except:
    st.sidebar.warning("⚠️ Logo not found.")

# Sidebar page selection
st.sidebar.header("📑 Select Template Page")
page = st.sidebar.radio("Choose Template:", ["EPC Proposal", "BESS Proposal"])

# Run the corresponding page
if page == "EPC Proposal":
    run_epc()
else:
    run_bess()
