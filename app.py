import streamlit as st
import pandas as pd
from docx import Document
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
    import re
    from docx import Document

    # Clean headers
    df.columns = df.columns.str.strip()
    df["Parameters"] = df["Parameters"].astype(str).str.strip()
    df["Value"] = df["Value"].astype(str)

    # Create lookup dictionary (case insensitive)
    param_dict = {p.lower(): v for p, v in zip(df["Parameters"], df["Value"])}

    # Load Word doc
    doc = Document(template_path)

    # --- helper: replace placeholders like {{Parameter Name}} ---
    def replace_placeholders(text):
        for key, value in param_dict.items():
            pattern = r"\{\{\s*" + re.escape(key) + r"\s*\}\}"
            text = re.sub(pattern, value, text, flags=re.IGNORECASE)
        return text

    # --- helper: process paragraph runs (merge + replace) ---
    def process_paragraphs(paragraphs):
        for para in paragraphs:
            if "{{" in para.text:
                full_text = "".join(run.text for run in para.runs)
                new_text = replace_placeholders(full_text)
                if new_text != full_text:
                    for r in para.runs:
                        r.text = ""
                    para.add_run(new_text)

    # --- main body ---
    process_paragraphs(doc.paragraphs)

    # --- tables ---
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if "{{" in cell.text:
                    cell.text = replace_placeholders(cell.text)

    # --- headers & footers (all sections) ---
    for section in doc.sections:
        header = section.header
        footer = section.footer
        process_paragraphs(header.paragraphs)
        process_paragraphs(footer.paragraphs)

        # Replace inside header/footer tables too
        for table in header.tables + footer.tables:
            for row in table.rows:
                for cell in row.cells:
                    if "{{" in cell.text:
                        cell.text = replace_placeholders(cell.text)

    return doc


# ========== App Flow ==========
if uploaded_excel is not None:
    try:
        df = pd.read_excel(uploaded_excel, engine="openpyxl")
        st.success("✅ Excel loaded successfully!")
        st.dataframe(df)

        if st.button("🚀 Generate Word Proposal"):
            try:
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
