import os
from datetime import datetime
import pandas as pd
from docxtpl import DocxTemplate
import win32com.client as win32

from config import *

today = datetime.now().strftime("%d %B %Y")

# Read Excel
df = pd.read_excel(EXCEL_FILE)
df.columns = df.columns.str.strip()

if STATUS_COLUMN not in df.columns:
    df[STATUS_COLUMN] = ""

os.makedirs(GENERATED_FOLDER, exist_ok=True)

# Start Microsoft Word
word = win32.Dispatch("Word.Application")
word.Visible = False

for _, row in df.iterrows():

    name = str(row[NAME_COLUMN]).strip()

    print(f"Generating {name}...")

    # Open Template
    doc = DocxTemplate(TEMPLATE_FILE)

    context = {
        "name": name,
        "date": today
    }

    doc.render(context)

    docx_path = os.path.abspath(
        os.path.join(GENERATED_FOLDER, f"{name}.docx")
    )

    pdf_path = os.path.abspath(
        os.path.join(GENERATED_FOLDER, f"{name}.pdf")
    )

    doc.save(docx_path)

    # Convert to PDF using Microsoft Word
    word_doc = word.Documents.Open(docx_path)

    word_doc.SaveAs(pdf_path, FileFormat=17)

    word_doc.Close()

    print(f"✓ {name}")

word.Quit()

print("\nAll Offer Letters Generated Successfully.")