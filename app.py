import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import tempfile, os, importlib.util, subprocess

# ============ Auto-install openpyxl if missing ============
if importlib.util.find_spec("openpyxl") is None:
    subprocess.run(["pip", "install", "openpyxl"], check=True)

# ============ Streamlit Page Setup ============
st.set_page_config(page_title="Proposal Auto Generator", layout="wide")
st.title("📄 Techno-Commercial Proposal Auto Generator")

st.markdown("""
Upload your Excel parameter sheet (any filename, one sheet).  
This app will fill your **Word proposal template** by replacing placeholders like `{{Name of Customer}}`,  
`{{Project Capacity}}`, etc., with values from Excel.
""")

# ============ File Upload ============
uploaded_excel = st.file_uploader("📤 Upload Excel file", type=["xlsx"])

# Path of your stored Word template (in GitHub repo)
TEMPLATE_PATH = "input_template.docx"   # Make sure this file exists in repo root


# ============ Function: Replace {{placeholders}} ============
def fill_template(df, template_path):
    # Create dictionary from Excel (Parameter : Value)
    df.columns = df.columns.str.strip()
    param_dict = dict(zip(df["Parameters"].astype(str).str.strip(), df["Value"].astype(str)))

    doc = Document(template_path)

    def replace_text(text):
        for key, value in param_dict.items():
            placeholder = f"{{{{{key}}}}}"  # example: {{Name of Customer}}
            if placeholder in text:
                text = text.replace(placeholder, value)
        return text

    # Replace in all paragraphs
    for para in doc.paragraphs:
        if "{{" in para.text and "}}" in para.text:
            inline = para.runs
            for i in range(len(inline)):
                inline[i].text = replace_text(inline[i].text)

    # Replace inside tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if "{{" in cell.text and "}}" in cell.text:
                    cell.text = replace_text(cell.text)

    return doc


# ============ Process Uploaded Excel ============
if uploaded_excel is not None:
    try:
        df = pd.read_excel(uploaded_excel, engine="openpyxl")
        st.success("✅ Excel file loaded successfully!")
        st.dataframe(df)

        if st.button("🚀 Generate Word Proposal"):
            filled_doc = fill_template(df, TEMPLATE_PATH)

            # Save Word output
            output_buffer = BytesIO()
            filled_doc.save(output_buffer)
            output_buffer.seek(0)

            st.download_button(
                label="⬇️ Download Word File",
                data=output_buffer,
                file_name="Generated_Proposal.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

            # Optional PDF Conversion
            try:
                from docx2pdf import convert
                with tempfile.TemporaryDirectory() as tmpdir:
                    temp_docx = os.path.join(tmpdir, "temp.docx")
                    temp_pdf = os.path.join(tmpdir, "temp.pdf")
                    filled_doc.save(temp_docx)
                    convert(temp_docx, temp_pdf)
                    with open(temp_pdf, "rb") as pdf_file:
                        st.download_button(
                            label="⬇️ Download PDF File",
                            data=pdf_file,
                            file_name="Generated_Proposal.pdf",
                            mime="application/pdf"
                        )
            except Exception:
                st.warning("⚠️ PDF conversion skipped (requires MS Word on Windows).")

    except Exception as e:
        st.error(f"❌ Error reading Excel file: {e}")
else:
    st.info("📥 Please upload your Excel parameter file to begin.")
