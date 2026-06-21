import hashlib
import os
from datetime import datetime
from pathlib import Path

import pandas as pd
from docxtpl import DocxTemplate
import win32com.client as win32
from tqdm import tqdm

from config import *
from utils import build_safe_attachment_name, setup_logging, unique_filepath, validate_email

logger = setup_logging("generate_offer_letters")
OUTPUT_FOLDER = Path(GENERATED_FOLDER)


def generate_internship_id(company_name: str, intern_name: str) -> str:
    """Generate deterministic 9-character ID: 4 company letters + 2 name letters + 3 hash digits"""
    try:
        company_prefix = (company_name or '').strip().upper()[:4].ljust(4, 'X')
        company_prefix = ''.join(ch for ch in company_prefix if ch.isalpha() or ch == 'X')[:4]
        name_prefix = (intern_name or '').strip().upper()[:2].ljust(2, 'X')
        name_prefix = ''.join(ch for ch in name_prefix if ch.isalpha() or ch == 'X')[:2]
        seed = f"{company_prefix}{name_prefix}".encode('utf-8')
        digest = hashlib.md5(seed).hexdigest()
        hash_digits = str(int(digest[:6], 16) % 1000).zfill(3)
        return f"{company_prefix}{name_prefix}{hash_digits}"
    except Exception as e:
        logger.error("Error generating internship ID: %s", e)
        return "XXXX00000"


def load_data():
    if not os.path.exists(EXCEL_FILE):
        logger.error("Excel file not found: %s", EXCEL_FILE)
        raise FileNotFoundError(EXCEL_FILE)

    df = pd.read_excel(EXCEL_FILE)
    df.columns = df.columns.str.strip()

    if NAME_COLUMN not in df.columns:
        logger.error("Missing required column: %s", NAME_COLUMN)
        raise KeyError(f"Missing required column: {NAME_COLUMN}")

    if STATUS_COLUMN not in df.columns:
        df[STATUS_COLUMN] = ""

    if INTERN_ID_COLUMN not in df.columns:
        df[INTERN_ID_COLUMN] = ""
    
    # Convert INTERN_ID_COLUMN to string type to allow storing alphanumeric IDs
    df[INTERN_ID_COLUMN] = df[INTERN_ID_COLUMN].fillna("").astype(str)

    return df


def render_row(row, row_index, word, df):
    name = str(row[NAME_COLUMN]).strip()
    email = str(row.get(EMAIL_COLUMN, "") or "").strip()
    company_name = str(row.get(COMPANY_COLUMN, "") or "").strip()
    role = str(row.get(ROLES_COLUMN, "") or "").strip()

    if not name:
        raise ValueError(f"Row {row_index + 1}: missing {NAME_COLUMN}")

    if email and not validate_email(email):
        logger.warning("Invalid email for %s: %s", name, email)

    if not os.path.exists(TEMPLATE_FILE):
        raise FileNotFoundError(f"Template file not found: {TEMPLATE_FILE}")

    intern_id = str(row.get(INTERN_ID_COLUMN, "") or "").strip()
    if not intern_id:
        intern_id = generate_internship_id(company_name, name)
        df.at[row_index, INTERN_ID_COLUMN] = intern_id

    # Extract all internship details from the row
    context = {
        "name": name,
        "company_name": company_name,
        "role": role,
        "date": datetime.now().strftime(DATE_FORMAT),
        "duty": str(row.get(DUTY_COLUMN, "")).strip(),
        "duration": str(row.get(DURATION_COLUMN, "")).strip(),
        "mode": str(row.get(MODE_COLUMN, "")).strip(),
        "commitment": str(row.get(COMMITMENT_COLUMN, "")).strip(),
        "stipend": str(row.get(STIPEND_COLUMN, "")).strip(),
        "intern_id": intern_id,
    }

    doc = DocxTemplate(TEMPLATE_FILE)
    doc.render(context)

    base_name = build_safe_attachment_name(name, email)
    if not base_name:
        base_name = f"intern-{row_index + 1}"

    docx_path = unique_filepath(OUTPUT_FOLDER / f"{base_name}.docx")
    pdf_path = unique_filepath(OUTPUT_FOLDER / f"{base_name}.pdf")

    doc.save(str(docx_path))

    word_doc = word.Documents.Open(str(docx_path))
    word_doc.SaveAs(str(pdf_path), FileFormat=17)
    word_doc.Close()

    logger.info("Generated %s and %s for ID: %s", docx_path.name, pdf_path.name, intern_id)


def main():
    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

    df = load_data()
    total = len(df)
    generated = 0
    skipped = 0
    failed = 0

    word = None
    try:
        if not os.path.exists(TEMPLATE_FILE):
            raise FileNotFoundError(f"Template file not found: {TEMPLATE_FILE}")

        word = win32.Dispatch("Word.Application")
        word.Visible = False

        for index, row in tqdm(df.iterrows(), total=total, desc="Generating", unit="letter"):
            try:
                render_row(row, index, word, df)
                generated += 1
            except ValueError as exc:
                skipped += 1
                logger.warning("Skipping row %s: %s", index + 1, exc)
            except Exception:
                failed += 1
                logger.exception("Failed to generate row %s", index + 1)
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                logger.exception("Unable to close Word application")
        
        # Save updated Excel with generated IDs
        try:
            df.to_excel(EXCEL_FILE, index=False)
            logger.info("Updated Excel file with generated IDs")
        except Exception as e:
            logger.error("Failed to update Excel file: %s", e)

    logger.info("Finished generation. Generated=%s, Skipped=%s, Failed=%s", generated, skipped, failed)
    print("\n===== Summary =====")
    print(f"Total rows: {total}")
    print(f"Generated: {generated}")
    print(f"Skipped: {skipped}")
    print(f"Failed: {failed}")


if __name__ == "__main__":
    main()
