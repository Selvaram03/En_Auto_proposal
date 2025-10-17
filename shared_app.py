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
    """
    Replace placeholders {{Parameter}} in:
      - Paragraphs
      - Tables
      - Headers
      - Footers
      - Text boxes (shapes)
    Handles special characters and empty keys/values.
    """
    pattern_template = r"\{\{\s*{}\s*\}\}"

    # Helper to safely replace text in runs
    def replace_in_paragraphs(paragraphs):
        for p in paragraphs:
            full_text = "".join(run.text for run in p.runs)
            new_text = full_text
            for key, val in param_dict.items():
                key_str = str(key) if key is not None else ""
                val_str = str(val) if val is not None else ""
                pattern = pattern_template.format(re.escape(key_str))
                new_text = re.sub(pattern, val_str, new_text, flags=re.IGNORECASE)
            if new_text != full_text:
                # Clear existing runs
                for run in p.runs:
                    run.text = ""
                p.add_run(new_text)

    # 1️⃣ Replace in main paragraphs
    replace_in_paragraphs(doc.paragraphs)

    # 2️⃣ Replace in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_in_paragraphs(cell.paragraphs)
                # Recursively replace in nested tables
                for nested_table in cell.tables:
                    for n_row in nested_table.rows:
                        for n_cell in n_row.cells:
                            replace_in_paragraphs(n_cell.paragraphs)

    # 3️⃣ Replace in headers and footers
    for section in doc.sections:
        replace_in_paragraphs(section.header.paragraphs)
        replace_in_paragraphs(section.footer.paragraphs)
        for table in section.header.tables + section.footer.tables:
            for row in table.rows:
                for cell in row.cells:
                    replace_in_paragraphs(cell.paragraphs)

    # 4️⃣ Replace in text boxes (shapes)
    for shape in doc.inline_shapes:
        try:
            if shape.text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    replace_in_paragraphs([paragraph])
        except AttributeError:
            pass

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
