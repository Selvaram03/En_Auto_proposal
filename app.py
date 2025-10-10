from io import BytesIO
import pandas as pd
from docx import Document
import streamlit as st
from docx2pdf import convert
import tempfile
import os

st.set_page_config(page_title="Auto Proposal Generator", layout="wide")
st.title("📄 Techno-Commercial Proposal Auto-Generator")

st.markdown("""
Upload your Excel parameter sheet, and this tool will auto-fill the Word proposal template.
""")

# --- File Uploads ---
uploaded_excel = st.file_uploader("📤 Upload Excel file", type=["xlsx"])
template_path = "input_template.docx"   # stored in your Git repo

# --- Function: Replace placeholders in Word ---
def fill_template(excel_df, template_path):
    param_dict = dict(zip(excel_df['Parameters'].str.strip(), excel_df['Value'].astype(str)))

    doc = Document(template_path)

    # Replace placeholders in paragraphs
    for para in doc.paragraphs:
        for key, value in param_dict.items():
            placeholder = f"({key})"
            if placeholder in para.text:
                inline = para.runs
                for i in range(len(inline)):
                    if placeholder in inline[i].text:
                        inline[i].text = inline[i].text.replace(placeholder, value)

    # Replace placeholders inside tables too
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in param_dict.items():
                    placeholder = f"({key})"
                    if placeholder in cell.text:
                        cell.text = cell.text.replace(placeholder, value)

    return doc


# --- Process Excel Upload ---
if uploaded_excel is not None:
    df = pd.read_excel(uploaded_excel)
    st.write("✅ Parameters loaded:")
    st.dataframe(df)

    if st.button("Generate Word Proposal"):
        filled_doc = fill_template(df, template_path)

        # Save filled Word file to buffer
        output_buffer = BytesIO()
        filled_doc.save(output_buffer)
        output_buffer.seek(0)

        st.success("🎉 Proposal generated successfully!")
        st.download_button(
            label="⬇️ Download Word File",
            data=output_buffer,
            file_name="Generated_Proposal.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

        # Optional: Generate PDF
        with tempfile.TemporaryDirectory() as tmpdirname:
            temp_docx = os.path.join(tmpdirname, "temp.docx")
            temp_pdf = os.path.join(tmpdirname, "temp.pdf")
            filled_doc.save(temp_docx)
            try:
                convert(temp_docx, temp_pdf)
                with open(temp_pdf, "rb") as pdf_file:
                    st.download_button(
                        label="⬇️ Download PDF File",
                        data=pdf_file,
                        file_name="Generated_Proposal.pdf",
                        mime="application/pdf"
                    )
            except Exception:
                st.warning("⚠️ PDF conversion skipped (docx2pdf requires MS Word on Windows).")
