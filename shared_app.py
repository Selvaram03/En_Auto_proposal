# shared_app.py
import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import re, os, tempfile

def load_excel(uploaded_file):
    """Read Excel and return dataframe with proper column checks."""
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        df.columns = df.columns.str.strip()
        if 'Parameters' not in df.columns or 'Value' not in df.columns:
            st.error("❌ Excel must have 'Parameters' and 'Value' columns.")
            return None
        df["Parameters"] = df["Parameters"].astype(str).str.strip()
        df["Value"] = df["Value"].astype(str)
        return df
    except Exception as e:
        st.error(f"❌ Error reading Excel: {e}")
        return None

def replace_placeholders(doc, param_dict):
    """Replace placeholders {{Parameter}} in paragraphs and tables."""
    pattern_template = r"\{\{\s*{}\s*\}\}"
    for p in doc.paragraphs:
        for key, val in param_dict.items():
            pattern = pattern_template.format(re.escape(key))
            p.text = re.sub(pattern, val, p.text, flags=re.IGNORECASE)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, val in param_dict.items():
                    pattern = pattern_template.format(re.escape(key))
                    for para in cell.paragraphs:
                        para.text = re.sub(pattern, val, para.text, flags=re.IGNORECASE)

def generate_word(doc_template_path, df):
    """Generate Word document from template and dataframe."""
    doc = Document(doc_template_path)
    param_dict = {p: v for p, v in zip(df["Parameters"], df["Value"])}
    replace_placeholders(doc, param_dict)
    return doc

def download_docx(doc, filename):
    """Show download button for Word document."""
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    st.download_button(
        label=f"⬇️ Download {filename}.docx",
        data=buffer,
        file_name=f"{filename}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

def download_pdf(doc, filename):
    """Convert to PDF if docx2pdf available and show download."""
    try:
        from docx2pdf import convert
        with tempfile.TemporaryDirectory() as tmpdir:
            docx_path = os.path.join(tmpdir, "temp.docx")
            pdf_path = os.path.join(tmpdir, "temp.pdf")
            doc.save(docx_path)
            convert(docx_path, pdf_path)
            with open(pdf_path, "rb") as f:
                st.download_button(
                    label=f"⬇️ Download {filename}.pdf",
                    data=f,
                    file_name=f"{filename}.pdf",
                    mime="application/pdf"
                )
    except Exception:
        st.warning("⚠️ PDF conversion skipped (requires MS Word or LibreOffice).")
