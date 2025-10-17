import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from io import BytesIO
import re, tempfile, os, importlib.util, subprocess
from lxml import etree  # Needed for XML manipulation

# ========== Auto-install dependencies if missing ==========
required_libs = ["openpyxl", "lxml"]
for lib in required_libs:
    if importlib.util.find_spec(lib) is None:
        try:
            subprocess.run(["pip", "install", lib], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        except subprocess.CalledProcessError:
            st.warning(f"Could not auto-install {lib}. Please install it manually.")

# ========== Streamlit Config ==========
st.set_page_config(page_title="Proposal Auto Generator", layout="wide")
st.title("📄 Techno-Commercial Proposal Auto Generator")

enrich_logo_path = r"enrich_logo.png"
try:
    st.sidebar.image(enrich_logo_path, width=150)
except Exception:
    st.sidebar.warning("Logo not found at provided path or could not be loaded.")

st.markdown("""
Upload your Excel sheet with **Parameters** and **Value** columns. The app will replace all placeholders in the Word template like `{{Parameter Name}}` automatically, including those in **text boxes**, **headers**, and **footers**.
""")

# ========== Default Excel Template Download ==========
TEMPLATE_EXCEL_PATH = "Input_EPC_Proposal.xlsx"
try:
    with open(TEMPLATE_EXCEL_PATH, "rb") as f:
        st.download_button(
            label="📥 Download Excel Template",
            data=f,
            file_name="Proposal_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    st.markdown("**ℹ️ Fill in this template with your inputs and upload it below to generate the Word document.**")
except FileNotFoundError:
    st.warning("⚠️ Default Excel template not found at the specified path.")

# Upload Excel
uploaded_excel = st.file_uploader("📤 Upload Excel File", type=["xlsx"])

# Template path
TEMPLATE_PATH = "EPC_template.docx"

# ========== Core Function Helpers ==========

def replace_in_xml(doc_part, param_dict):
    """Replaces text in a document part (body, header, or footer) using XML."""
    try:
        if doc_part.element is None:
            return
        root = doc_part.element.getroottree()
    except AttributeError:
        return

    namespaces = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'v': 'urn:schemas-microsoft-com:vml', 
    }

    for key, value in param_dict.items():
        placeholder = "{{" + key + "}}"
        for elem in root.xpath('//w:t|//v:t', namespaces=namespaces):
            if elem.text and placeholder.lower() in elem.text.lower():
                original_text = elem.text
                pattern = r"\{\{\s*" + re.escape(key) + r"\s*\}\}"
                new_text = re.sub(pattern, value, original_text, flags=re.IGNORECASE)
                elem.text = new_text

def process_paragraphs(paragraphs, param_dict):
    """Processes paragraphs using the standard docx API for formatting."""
    def replace_placeholders(text):
        for key, value in param_dict.items():
            pattern = r"\{\{\s*" + re.escape(key) + r"\s*\}\}"
            text = re.sub(pattern, value, text, flags=re.IGNORECASE)
        return text

    for para in paragraphs:
        if "{{" in para.text:
            full_text = "".join(run.text for run in para.runs)
            new_text = replace_placeholders(full_text)
            if new_text != full_text:
                first_run = para.runs[0] if para.runs else None
                first_run_style = first_run.style if first_run else None
                first_run_font_size = first_run.font.size if first_run and first_run.font.size else None
                for r in para.runs:
                    r.text = ""
                new_run = para.add_run(new_text)
                if first_run_style:
                    new_run.style = first_run_style
                new_run.font.name = 'Calibri'
                if first_run_font_size:
                     new_run.font.size = first_run_font_size

def process_cell(cell, param_dict):
    """Processes paragraphs in a cell, including nested tables, for formatting fix."""
    process_paragraphs(cell.paragraphs, param_dict)
    for nested_table in cell.tables:
        for row in nested_table.rows:
            for cell in row.cells:
                process_cell(cell, param_dict)

def fill_template(df, template_path):
    param_dict = {p.lower(): v for p, v in zip(df["Parameters"], df["Value"])}
    doc = Document(template_path)

    replace_in_xml(doc, param_dict)
    process_paragraphs(doc.paragraphs, param_dict)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                process_cell(cell, param_dict)

    for section in doc.sections:
        replace_in_xml(section.header, param_dict)
        replace_in_xml(section.footer, param_dict)
        process_paragraphs(section.header.paragraphs, param_dict)
        process_paragraphs(section.footer.paragraphs, param_dict)
        try:
            for table in section.header.tables + section.footer.tables:
                for row in table.rows:
                    for cell in row.cells:
                        process_cell(cell, param_dict)
        except AttributeError:
            pass

    return doc

# ========== App Flow ==========
if uploaded_excel is not None:
    try:
        df = pd.read_excel(uploaded_excel, engine="openpyxl")
        df.columns = df.columns.str.strip()
        if 'Parameters' not in df.columns or 'Value' not in df.columns:
            st.error("❌ Error: The Excel sheet must contain columns named 'Parameters' and 'Value'.")
            st.stop()
        df["Parameters"] = df["Parameters"].astype(str).str.strip()
        df["Value"] = df["Value"].astype(str)

        st.success("✅ Excel loaded successfully!")
        st.dataframe(df)

        if st.button("🚀 Generate Word Proposal"):
            try:
                filled_doc = fill_template(df, TEMPLATE_PATH)
                buffer = BytesIO()
                filled_doc.save(buffer)
                buffer.seek(0)

                st.download_button(
                    label="⬇️ Download Word File",
                    data=buffer,
                    file_name="Generated_Proposal.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

                try:
                    from docx2pdf import convert
                    with tempfile.TemporaryDirectory() as tmpdir:
                        docx_path = os.path.join(tmpdir, "temp.docx")
                        pdf_path = os.path.join(tmpdir, "temp.pdf")
                        filled_doc.save(docx_path)
                        convert(docx_path, pdf_path)
                        with open(pdf_path, "rb") as pdf_file:
                            st.download_button(
                                label="⬇️ Download PDF File",
                                data=pdf_file,
                                file_name="Generated_Proposal.pdf",
                                mime="application/pdf",
                            )
                except Exception:
                    st.warning("⚠️ PDF conversion skipped (requires MS Word on Windows/Linux with LibreOffice).")

            except Exception as e:
                st.error(f"❌ Error generating proposal: {e}")

    except Exception as e:
        st.error(f"❌ Error reading Excel: {e}. Please ensure the first sheet is a valid Excel format and has the required columns.")
else:
    st.info("📥 Please upload your Excel file to begin.")
