import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import os, tempfile, re

# ========== Streamlit Config ==========
st.set_page_config(page_title="Proposal Auto Generator", layout="wide")
st.title("📄 Techno-Commercial Proposal Auto Generator")

# ========== Sidebar Logo ==========
enrich_logo_path = r"enrich_logo.png"
try:
    st.sidebar.image(enrich_logo_path, width=150)
except Exception:
    st.sidebar.warning("⚠️ Logo not found or could not be loaded.")

# ========== Page Selection ==========
st.sidebar.header("📑 Select Template Page")
page = st.sidebar.radio("Go to:", ["EPC Proposal", "BESS Proposal"])

# ========== Template Settings ==========
if page == "EPC Proposal":
    TEMPLATE_PATH = "EPC_Template.docx"
    TEMPLATE_EXCEL_PATH = "Input_EPC_Proposal.xlsx"
    session_prefix = "epc_"
else:
    TEMPLATE_PATH = "BESS_Template.docx"
    TEMPLATE_EXCEL_PATH = "Input_BESS_Proposal.xlsx"
    session_prefix = "bess_"

# Initialize session state for this page
for key in ["uploaded", "df", "generated_doc"]:
    session_key = session_prefix + key
    if session_key not in st.session_state:
        st.session_state[session_key] = None

# ========== Active Template Banner ==========
st.markdown(
    f"""
    <div style="background-color:#f0f8ff;padding:12px;border-radius:10px;margin-bottom:10px;">
        <b>🧩 Currently Selected Template:</b> 
        <span style="color:#0056b3;">{page}</span>
    </div>
    """,
    unsafe_allow_html=True
)

# ========== Instructions ==========
st.markdown("""
Upload your Excel sheet with **Parameters** and **Value** columns.  
The app will replace placeholders like `{{Parameter Name}}` in the Word template,  
including those inside **text boxes**, **headers**, and **footers**.
""")

# ========== Download Template ==========
try:
    with open(TEMPLATE_EXCEL_PATH, "rb") as f:
        st.download_button(
            label=f"📥 Download {page} Excel Template",
            data=f,
            file_name=os.path.basename(TEMPLATE_EXCEL_PATH),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
except FileNotFoundError:
    st.warning(f"⚠️ {page} Excel template not found.")

# ========== Upload Excel ==========
uploaded_excel = st.file_uploader(f"📤 Upload Excel for {page}", type=["xlsx"])
if uploaded_excel is not None:
    st.session_state[session_prefix + "uploaded"] = uploaded_excel

# ========== Read Excel ==========
if st.session_state[session_prefix + "uploaded"]:
    try:
        df = pd.read_excel(st.session_state[session_prefix + "uploaded"], engine="openpyxl")
        df.columns = df.columns.str.strip()
        if 'Parameters' not in df.columns or 'Value' not in df.columns:
            st.error("❌ Excel must have 'Parameters' and 'Value' columns.")
        else:
            df["Parameters"] = df["Parameters"].astype(str).str.strip()
            df["Value"] = df["Value"].astype(str)
            st.session_state[session_prefix + "df"] = df
    except Exception as e:
        st.error(f"❌ Error reading Excel: {e}")

# ========== Show Excel Table ==========
if st.session_state[session_prefix + "df"] is not None:
    st.success("✅ Excel loaded successfully!")
    st.dataframe(st.session_state[session_prefix + "df"])

# ========== Helper Functions ==========
def replace_placeholders(doc, param_dict):
    pattern_template = r"\{\{\s*{}\s*\}\}"
    for p in doc.paragraphs:
        for key, val in param_dict.items():
            pattern = pattern_template.format(re.escape(key))
            p.text = re.sub(pattern, val, p.text, flags=re.IGNORECASE)
    # Tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, val in param_dict.items():
                    pattern = pattern_template.format(re.escape(key))
                    for para in cell.paragraphs:
                        para.text = re.sub(pattern, val, para.text, flags=re.IGNORECASE)

# ========== Generate Proposal ==========
if st.session_state[session_prefix + "df"] is not None and st.button(f"🚀 Generate {page} Proposal"):
    try:
        doc = Document(TEMPLATE_PATH)
        param_dict = {p: v for p, v in zip(st.session_state[session_prefix + "df"]["Parameters"],
                                          st.session_state[session_prefix + "df"]["Value"])}
        replace_placeholders(doc, param_dict)
        st.session_state[session_prefix + "generated_doc"] = doc

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.download_button(
            label=f"⬇️ Download {page} Word File",
            data=buffer,
            file_name=f"Generated_{page.replace(' ','_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.error(f"❌ Error generating proposal: {e}")

# ========== Info if no upload ==========
if st.session_state[session_prefix + "df"] is None:
    st.info(f"📥 Please upload your Excel file for {page} to begin.")
