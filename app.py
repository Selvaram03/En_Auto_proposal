import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt  # NEW IMPORT: Added for explicit font size handling
from io import BytesIO
import re, tempfile, os, importlib.util, subprocess

# ========== Auto-install openpyxl if missing ==========
if importlib.util.find_spec("openpyxl") is None:
    subprocess.run(["pip", "install", "openpyxl"], check=True)

# ========== Streamlit Config ==========
st.set_page_config(page_title="Proposal Auto Generator", layout="wide")
st.title("📄 Techno-Commercial Proposal Auto Generator")

enrich_logo_path = r"enrich_logo.png"
try:
    st.sidebar.image(enrich_logo_path, width=150)
except:
    st.sidebar.warning("Logo not found at provided path.")

st.markdown("""
Upload your Excel sheet with **Parameters** and **Value** columns.  
The app will replace all placeholders in the Word template like `{{Parameter Name}}` automatically,  
including those in **headers** and **footers**.
""")

# Upload Excel
uploaded_excel = st.file_uploader("📤 Upload Excel File", type=["xlsx"])

# Template path (stored in GitHub repo)
TEMPLATE_PATH = "input_template.docx"


# ========== Core Function ==========
def fill_template(df, template_path):
    # Clean headers
    df.columns = df.columns.str.strip()
    df["Parameters"] = df["Parameters"].astype(str).str.strip()
    df["Value"] = df["Value"].astype(str)

    # Create lookup dictionary (case insensitive)
    param_dict = {p.lower(): v for p, v in zip(df["Parameters"], df["Value"])}

    # Load Word doc
    doc = Document(template_path)

    # --- helper: replace placeholders like {{Parameter Name}} (Unchanged) ---
    def replace_placeholders(text):
        for key, value in param_dict.items():
            # Use an efficient regex pattern
            pattern = r"\{\{\s*" + re.escape(key) + r"\s*\}\}"
            text = re.sub(pattern, value, text, flags=re.IGNORECASE)
        return text

    # --- helper: process paragraph runs (FIXED: Preserves style/font - Issue 3) ---
    def process_paragraphs(paragraphs):
        for para in paragraphs:
            if "{{" in para.text:
                full_text = "".join(run.text for run in para.runs)
                new_text = replace_placeholders(full_text)

                if new_text != full_text:
                    # Capture formatting properties of the first run BEFORE clearing
                    first_run = para.runs[0] if para.runs else None
                    first_run_style = first_run.style if first_run else None
                    
                    # Capture font size (which is an object like Pt)
                    first_run_font_size = first_run.font.size if first_run and first_run.font.size else None

                    # Clear all existing runs
                    for r in para.runs:
                        r.text = ""
                    
                    # Add the new text
                    new_run = para.add_run(new_text)

                    # Restore style/formatting
                    if first_run_style:
                        new_run.style = first_run_style
                    
                    # Explicitly set the font to Calibri (Fix for Issue 3)
                    new_run.font.name = 'Calibri'
                    
                    # Restore size if captured
                    if first_run_font_size:
                         new_run.font.size = first_run_font_size

    # --- helper: process cell content (NEW: Handles nested tables/cells - Issue 2) ---
    def process_cell(cell):
        # Process paragraphs within the cell (handles replacement and formatting)
        process_paragraphs(cell.paragraphs)

        # Recursively process any nested tables
        for nested_table in cell.tables:
            for row in nested_table.rows:
                for cell in row.cells:
                    process_cell(cell) # Recursive call

    # --- main body (includes 1st page content) ---
    process_paragraphs(doc.paragraphs)

    # --- tables (Now uses recursive cell processor for all main body tables/cells) ---
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                process_cell(cell)

    # --- headers & footers (all sections) ---
    for section in doc.sections:
        header = section.header
        footer = section.footer
        
        # Process paragraphs in header/footer
        process_paragraphs(header.paragraphs)
        process_paragraphs(footer.paragraphs)

        # Process tables in header/footer (Uses recursive cell processor now)
        for table in header.tables + footer.tables:
            for row in table.rows:
                for cell in row.cells:
                    process_cell(cell)

    return doc


# ========== App Flow (Unchanged) ==========
if uploaded_excel is not None:
    try:
        df = pd.read_excel(uploaded_excel, engine="openpyxl")
        st.success("✅ Excel loaded successfully!")
        st.dataframe(df)

        if st.button("🚀 Generate Word Proposal"):
            try:
                # The docx import is now globally available
                filled_doc = fill_template(df, TEMPLATE_PATH) 

                # Save to memory
                buffer = BytesIO()
                filled_doc.save(buffer)
                buffer.seek(0)

                # Download Word
                st.download_button(
                    label="⬇️ Download Word File",
                    data=buffer,
                    file_name="Generated_Proposal.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

                # Optional PDF export
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
                    st.warning("⚠️ PDF conversion skipped (requires MS Word on Windows).")

            except Exception as e:
                st.error(f"❌ Error generating proposal: {e}")

    except Exception as e:
        st.error(f"❌ Error reading Excel: {e}")
else:
    st.info("📥 Please upload your Excel file to begin.")
