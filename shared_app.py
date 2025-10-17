# shared_app.py
import re
from docx import Document

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
