import pandas as pd
from docx import Document
from io import BytesIO
import streamlit as st
import tempfile, os

# ===== Streamlit App Config =====
st.set_page_config(page_title="Proposal Auto Generator", layout="wide")
st.title("📄 Techno-Commercial Proposal Auto Generator")

st.markdown("""
Upload your Excel parameter sheet (any filename, one sheet), and this tool will auto-fill 
the Word proposal template using your placeholders.
""")

# ===== Upload Excel =====
uploaded_excel = st.file_uploader("📤 Upload Excel file", type=["xlsx"])

# ===== Define Template Path (stored in your GitHub repo) =====
TEMPLATE_PATH = "input_template.docx"  # Must be in same folder on GitHub


# ===== Function: Replace Placeholders =====
def fill_template(df, template_path):
    # Convert Excel data into dictionary
    df.columns = df.columns.str.strip()
    param_dict = dict(zip(df["Parameters"].astype(str).str.strip(), df["Value"].astype(str)))

    doc = Document(template_path)

    # Replace in normal paragraphs
    for para in doc.paragraphs:
        for key, value in param_dict.items():
            placeholder = f"({key})"
            if placeholder in para.text:
                inline = para.runs
                for i in range(len(inline)):
                    if placeholder in inline[i].text:
                        inline[i].text = inline[i].text.replace(placeholder, value)

    # Replace inside tables too
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in param_dict.items():
                    placeholder = f"({key})"
                    if placeholder in cell.text:
                        cell.text = cell.text.replace(placeholder, value)

    return doc


# ===== Main Processing =====
if uploaded_excel is not None:
    try:
        # Explicitly use openpyxl engine
        df = pd.read_excel(uploaded_excel, engine="openpyxl")
    except Exception as e:
        st.error(f"❌ Error reading Excel file: {e}")
        st.stop()

    st.success("✅ Excel file uploaded successfully!")
    st.dataframe(df)

    if st.button("🚀 Generate Word Proposal"):
        try:
            filled_doc = fill_template(df, TEMPLATE_PATH)

            # Save output Word file
            output_buffer = BytesIO()
            filled_doc.save(output_buffer)
            output_buffer.seek(0)

            st.download_button(
                label="⬇️ Download Word File",
                data=output_buffer,
                file_name="Generated_Proposal.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

            # Optional PDF conversion using temporary file
            try:
                from docx2pdf import convert
                with tempfile.TemporaryDirectory() as tmpdirname:
                    temp_docx = os.path.join(tmpdirname, "temp.docx")
                    temp_pdf = os.path.join(tmpdirname, "temp.pdf")
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
            st.error(f"⚠️ Error while generating proposal: {e}")
