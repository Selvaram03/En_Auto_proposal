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

st.markdown("""
Upload your Excel sheet with **Parameters** and **Value** columns.  
The app will replace all placeholders in the Word template like `{{Parameter Name}}` automatically.
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

    # Function to replace placeholders like {{Parameter Name}}
    def replace_placeholders(text):
        matches = re.findall(r"\{\{(.*?)\}\}", text)
        for m in matches:
            key = m.strip().lower()
            if key in param_dict:
                text = text.replace(f"{{{{{m}}}}}", param_dict[key])
        return text

    # Replace in all paragraphs
    for para in doc.paragraphs:
        if "{{" in para.text:
            for run in para.runs:
                run.text = replace_placeholders(run.text)

    # Replace in all tables (cells)
    for table in doc.tables:
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

                # PDF export (optional)
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
