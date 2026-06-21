import os
import time
import smtplib
import random
import string
from datetime import datetime
from pathlib import Path
from email.message import EmailMessage
from shutil import copy2
from docxtpl import DocxTemplate
import win32com.client as win32

import pandas as pd
from tqdm import tqdm

from config import *
from utils import build_safe_attachment_name, find_matching_file, setup_logging, validate_email, unique_filepath

logger = setup_logging("send_offer_letters")
OUTPUT_FOLDER = Path(GENERATED_FOLDER)


def load_data():
    if not os.path.exists(EXCEL_FILE):
        logger.error("Excel file not found: %s", EXCEL_FILE)
        raise FileNotFoundError(EXCEL_FILE)

    df = pd.read_excel(EXCEL_FILE)
    df.columns = df.columns.str.strip()

    if STATUS_COLUMN not in df.columns:
        df[STATUS_COLUMN] = ""

    if INTERN_ID_COLUMN not in df.columns:
        df[INTERN_ID_COLUMN] = ""

    df[STATUS_COLUMN] = df[STATUS_COLUMN].fillna("").astype(str)
    df[INTERN_ID_COLUMN] = df[INTERN_ID_COLUMN].fillna("").astype(str)
    return df


def get_pdf_path(name: str, email: str):
    base_name = build_safe_attachment_name(name, email)
    file_path = find_matching_file(OUTPUT_FOLDER, base_name, ".pdf")
    if file_path:
        return file_path

    alternate_name = OUTPUT_FOLDER / f"{base_name}.pdf"
    return alternate_name if alternate_name.exists() else None


def generate_internship_id(company_name: str, intern_name: str) -> str:
    """Generate deterministic 9-character ID: 4 company letters + 2 name letters + 3 hash digits"""
    try:
        import hashlib
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


def get_sop_path() -> Path:
    """Get SOP file path from Template folder (as PDF)"""
    sop_path = Path(SOP_FILE)
    # Check for PDF version of SOP
    if not sop_path.exists():
        pdf_sop = sop_path.with_suffix('.pdf')
        if pdf_sop.exists():
            return pdf_sop
        logger.warning("SOP file not found at %s or %s", SOP_FILE, pdf_sop)
        return None
    return sop_path


def prepare_biopay_agreement(duration: str, date: str, role: str, company_name: str) -> Path:
    """Copy Biopay Agreement, update placeholders, and convert to PDF"""
    biopay_template = Path(BIOPAY_AGREEMENT_FILE)
    if not biopay_template.exists():
        logger.warning("Biopay Agreement template not found at %s", BIOPAY_AGREEMENT_FILE)
        return None

    try:
        doc = DocxTemplate(str(biopay_template))
        context = {
            "duration": duration,
            "date": date,
            "role": role,
            "company_name": company_name,
        }
        doc.render(context)
        
        docx_output = unique_filepath(OUTPUT_FOLDER / "Biopay_Agreement_Updated.docx")
        doc.save(str(docx_output))
        logger.info("Updated Biopay Agreement with duration: %s and date: %s", duration, date)
        
        # Convert DOCX to PDF
        pdf_output = docx_output.with_suffix('.pdf')
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        try:
            word_doc = word.Documents.Open(str(docx_output))
            word_doc.SaveAs(str(pdf_output), FileFormat=17)
            word_doc.Close()
            logger.info("Converted Biopay Agreement to PDF: %s", pdf_output)
        finally:
            word.Quit()
        
        # Delete the DOCX file as we only need the PDF
        try:
            os.remove(str(docx_output))
        except Exception as e:
            logger.warning("Could not delete temp DOCX file: %s", e)
        
        return pdf_output
    except Exception as e:
        logger.error("Error preparing Biopay Agreement: %s", e)
        return None


def send_email(smtp, index, row):
    name = str(row[NAME_COLUMN]).strip()
    email = str(row.get(EMAIL_COLUMN, "") or "").strip()
    status = str(row.get(STATUS_COLUMN, "") or "").strip().lower()
    role = str(row.get(ROLES_COLUMN, "") or "").strip()
    duty = str(row.get(DUTY_COLUMN, "") or "").strip()
    company_name = str(row.get(COMPANY_COLUMN, "") or "").strip()
    duration = str(row.get(DURATION_COLUMN, "") or "").strip()

    default_id = ""
    if status == "sent":
        logger.info("Already sent: %s", name)
        return "already_sent", default_id

    if not name:
        logger.warning("Skipping row %s: missing %s", index + 1, NAME_COLUMN)
        return "skipped", default_id

    if not email:
        logger.warning("Skipping %s: no email provided", name)
        return "skipped", default_id

    if not validate_email(email):
        logger.warning("Skipping %s: invalid email %s", name, email)
        return "invalid_email", default_id

    if not role:
        logger.warning("Skipping %s: no role specified", name)
        return "skipped", default_id

    if not duty:
        logger.warning("Skipping %s: no duty specified", name)
        return "skipped", default_id

    if not company_name:
        logger.warning("Skipping %s: no company name specified", name)
        return "skipped", default_id

    if not duration:
        logger.warning("Skipping %s: no duration specified", name)
        return "skipped", default_id

    pdf_path = get_pdf_path(name, email)
    if not pdf_path or not pdf_path.exists():
        logger.error("Missing PDF for %s at expected path %s", name, OUTPUT_FOLDER)
        return "missing_file", default_id

    # Use existing internship ID from Excel if available, otherwise generate
    internship_id = str(row.get(INTERN_ID_COLUMN, "") or "").strip()
    if not internship_id:
        internship_id = generate_internship_id(company_name, name)

    # Select email template based on role
    email_template = EMAIL_BODY_TECHNICAL if role.lower() == "technical" else EMAIL_BODY_NON_TECHNICAL
    email_body = email_template.format(
        name=name,
        role=role,
        duty=duty,
        company_name=company_name,
        duration=duration,
        internship_id=internship_id
    )

    subject = EMAIL_SUBJECT
    try:
        subject = EMAIL_SUBJECT.format(role=role, company_name=company_name)
    except Exception:
        logger.debug("EMAIL_SUBJECT formatting skipped or invalid: %s", EMAIL_SUBJECT)

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = GMAIL_EMAIL
    msg["To"] = email
    msg.set_content(email_body)

    # Attach Offer Letter PDF
    with open(pdf_path, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="pdf",
            filename=pdf_path.name,
        )

    # Attach SOP (as PDF)
    sop_path = get_sop_path()
    if sop_path:
        with open(sop_path, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="pdf",
                filename="SOP.pdf",
            )
        logger.info("Attached SOP for %s", name)
    else:
        logger.warning("SOP not available for %s", name)

    offer_date = datetime.now().strftime(DATE_FORMAT)
    # Attach Biopay Agreement (converted to PDF with duration and date updated)
    biopay_path = prepare_biopay_agreement(duration, offer_date, role, company_name)
    if biopay_path:
        with open(biopay_path, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="pdf",
                filename="Biopay_Agreement.pdf",
            )
        logger.info("Attached Biopay Agreement for %s", name)
    else:
        logger.warning("Biopay Agreement not available for %s", name)

    smtp.send_message(msg)
    logger.info("Sent email to %s <%s> with ID %s", name, email, internship_id)
    return "sent", internship_id


def main():
    df = load_data()
    total = len(df)
    summary = {
        "sent": 0,
        "already_sent": 0,
        "skipped": 0,
        "invalid_email": 0,
        "missing_file": 0,
        "failed": 0,
    }

    if not OUTPUT_FOLDER.exists():
        OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

    logger.info("Starting email send process for %s rows", total)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(GMAIL_EMAIL, GMAIL_APP_PASSWORD)
            for index, row in tqdm(df.iterrows(), total=total, desc="Sending", unit="email"):
                try:
                    result, internship_id = send_email(smtp, index, row)
                    if result:
                        summary[result] = summary.get(result, 0) + 1
                        if result == "sent":
                            df.at[index, STATUS_COLUMN] = "Sent"
                            df.at[index, INTERN_ID_COLUMN] = internship_id
                            time.sleep(2)
                except Exception:
                    summary["failed"] += 1
                    logger.exception("Failed to send to row %s", index + 1)
    except Exception:
        logger.exception("SMTP connection failed")
        raise
    finally:
        df.to_excel(EXCEL_FILE, index=False)
        logger.info("Saved updated Excel status: %s", EXCEL_FILE)

    print("\n===== Summary =====")
    print(f"Total rows: {total}")
    print(f"Sent: {summary['sent']}")
    print(f"Already sent: {summary['already_sent']}")
    print(f"Skipped: {summary['skipped']}")
    print(f"Invalid emails: {summary['invalid_email']}")
    print(f"Missing PDFs: {summary['missing_file']}")
    print(f"Failed: {summary['failed']}")
    logger.info("Send summary: %s", summary)


if __name__ == "__main__":
    main()
