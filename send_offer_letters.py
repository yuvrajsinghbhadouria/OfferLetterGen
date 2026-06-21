import os
import smtplib
import time
import threading
from datetime import datetime
from pathlib import Path
from email.message import EmailMessage
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
from tqdm import tqdm

from config import *
from utils import build_safe_attachment_name, generate_internship_id, setup_logging, validate_email

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


def build_pdf_lookup() -> dict[str, Path]:
    lookup = {}
    for path in OUTPUT_FOLDER.glob("*.pdf"):
        key = path.stem.lower()
        lookup[key] = path
    return lookup


def get_cached_pdf_path(name: str, email: str, lookup: dict[str, Path]):
    base_name = build_safe_attachment_name(name, email).lower()
    return lookup.get(base_name)


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


def get_biopay_agreement_path() -> Path:
    """Get Biopay Agreement PDF from Template folder"""
    biopay_path = Path(BIOPAY_AGREEMENT_FILE)
    if biopay_path.exists():
        return biopay_path
    
    logger.warning("Biopay Agreement not found at %s", biopay_path)
    return None


def precache_attachment_bytes(sop_path, biopay_path):
    """Read SOP and Biopay Agreement into memory ONCE. Returns (sop_bytes, biopay_bytes)."""
    sop_bytes = None
    biopay_bytes = None

    if sop_path and sop_path.exists():
        sop_bytes = sop_path.read_bytes()
        logger.info("Pre-cached SOP (%s bytes)", len(sop_bytes))

    if biopay_path and biopay_path.exists():
        biopay_bytes = biopay_path.read_bytes()
        logger.info("Pre-cached Biopay Agreement (%s bytes)", len(biopay_bytes))

    return sop_bytes, biopay_bytes


def precache_offer_pdfs(pdf_lookup: dict[str, Path]) -> dict[str, bytes]:
    """Read all offer letter PDFs into memory ONCE for fast attachment."""
    pdf_bytes_cache = {}
    for key, path in pdf_lookup.items():
        try:
            pdf_bytes_cache[key] = path.read_bytes()
        except Exception as e:
            logger.warning("Could not pre-read PDF %s: %s", path.name, e)
    logger.info("Pre-cached %s offer letter PDFs", len(pdf_bytes_cache))
    return pdf_bytes_cache


def build_email_job(index, row, pdf_lookup, pdf_bytes_cache, sop_bytes, biopay_bytes):
    """Validate and build a complete email job dict. Returns (job_dict, result_status) — 
    if result_status is not None, the row was skipped/invalid."""
    name = str(row[NAME_COLUMN]).strip()
    email = str(row.get(EMAIL_COLUMN, "") or "").strip()
    status = str(row.get(STATUS_COLUMN, "") or "").strip().lower()
    role = str(row.get(ROLES_COLUMN, "") or "").strip()
    duty = str(row.get(DUTY_COLUMN, "") or "").strip()
    company_name = str(row.get(COMPANY_COLUMN, "") or "").strip()
    duration = str(row.get(DURATION_COLUMN, "") or "").strip()

    if status == "sent":
        logger.info("Already sent: %s", name)
        return None, "already_sent"

    if not name:
        logger.warning("Skipping row %s: missing %s", index + 1, NAME_COLUMN)
        return None, "skipped"

    if not email:
        logger.warning("Skipping %s: no email provided", name)
        return None, "skipped"

    if not validate_email(email):
        logger.warning("Skipping %s: invalid email %s", name, email)
        return None, "invalid_email"

    if not role:
        logger.warning("Skipping %s: no role specified", name)
        return None, "skipped"

    if not duty:
        logger.warning("Skipping %s: no duty specified", name)
        return None, "skipped"

    if not company_name:
        logger.warning("Skipping %s: no company name specified", name)
        return None, "skipped"

    if not duration:
        logger.warning("Skipping %s: no duration specified", name)
        return None, "skipped"

    # Check for PDF
    pdf_key = build_safe_attachment_name(name, email).lower()
    offer_pdf_bytes = pdf_bytes_cache.get(pdf_key)
    if offer_pdf_bytes is None:
        # Fallback: check on disk
        pdf_path = get_cached_pdf_path(name, email, pdf_lookup)
        if not pdf_path:
            logger.error("Missing PDF for %s", name)
            return None, "missing_file"
        offer_pdf_bytes = pdf_path.read_bytes()

    pdf_filename = pdf_lookup.get(pdf_key, Path(f"{pdf_key}.pdf")).name

    # Internship ID
    internship_id = str(row.get(INTERN_ID_COLUMN, "") or "").strip()
    if not internship_id:
        internship_id = generate_internship_id(company_name, name)

    # Build email template
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

    # Build the full EmailMessage object
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = GMAIL_EMAIL
    msg["To"] = email
    msg.set_content(email_body)

    # Attach Offer Letter PDF (from pre-cached bytes)
    msg.add_attachment(
        offer_pdf_bytes,
        maintype="application",
        subtype="pdf",
        filename=pdf_filename,
    )

    # Attach SOP (from pre-cached bytes)
    if sop_bytes:
        msg.add_attachment(
            sop_bytes,
            maintype="application",
            subtype="pdf",
            filename="SOP.pdf",
        )

    # Attach Biopay Agreement (from pre-cached bytes)
    if biopay_bytes:
        msg.add_attachment(
            biopay_bytes,
            maintype="application",
            subtype="pdf",
            filename="Biopay_Agreement.pdf",
        )

    return {
        "index": index,
        "name": name,
        "email": email,
        "msg": msg,
        "internship_id": internship_id,
    }, None


def _create_smtp_connection():
    """Create an SMTP connection using config settings. Supports SSL and STARTTLS."""
    if SMTP_USE_TLS:
        # STARTTLS on port 587 (Outlook, some custom servers)
        conn = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
        conn.ehlo()
        conn.starttls()
        conn.ehlo()
    else:
        # Direct SSL on port 465 (Gmail, Yahoo, Zoho)
        conn = smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT)
    conn.login(SENDER_EMAIL, SENDER_PASSWORD)
    return conn


def send_single_email(job):
    """Send a single email using a per-thread SMTP connection. Thread-safe."""
    thread_id = threading.current_thread().name
    name = job["name"]
    email = job["email"]
    internship_id = job["internship_id"]

    try:
        with _create_smtp_connection() as smtp:
            smtp.send_message(job["msg"])
        logger.info("[%s] Sent email to %s <%s> with ID %s", thread_id, name, email, internship_id)
        return job["index"], "sent", internship_id
    except Exception as e:
        logger.exception("[%s] Failed to send to %s <%s>", thread_id, name, email)
        return job["index"], "failed", internship_id


def main():
    start_time = time.perf_counter()

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

    # ---- Pre-cache all file data into memory ----
    pdf_lookup = build_pdf_lookup()
    sop_path = get_sop_path()
    biopay_path = get_biopay_agreement_path()

    sop_bytes, biopay_bytes = precache_attachment_bytes(sop_path, biopay_path)
    pdf_bytes_cache = precache_offer_pdfs(pdf_lookup)

    logger.info("Starting email send process for %s rows (workers=%s)", total, MAX_EMAIL_WORKERS)

    # ---- Build all email jobs (validation + message construction) ----
    email_jobs = []
    for index, row in df.iterrows():
        job, skip_reason = build_email_job(
            index, row, pdf_lookup, pdf_bytes_cache, sop_bytes, biopay_bytes
        )
        if skip_reason:
            summary[skip_reason] = summary.get(skip_reason, 0) + 1
        elif job:
            email_jobs.append(job)

    logger.info("Prepared %s emails to send, %s skipped", len(email_jobs), total - len(email_jobs))

    # ---- Send emails concurrently ----
    if email_jobs:
        df_lock = threading.Lock()

        with ThreadPoolExecutor(max_workers=MAX_EMAIL_WORKERS) as pool:
            futures = {pool.submit(send_single_email, job): job for job in email_jobs}
            for future in tqdm(
                as_completed(futures), total=len(futures), desc="Sending", unit="email"
            ):
                try:
                    idx, result, internship_id = future.result()
                    summary[result] = summary.get(result, 0) + 1
                    if result == "sent":
                        with df_lock:
                            df.at[idx, STATUS_COLUMN] = "Sent"
                            df.at[idx, INTERN_ID_COLUMN] = internship_id
                except Exception:
                    summary["failed"] += 1
                    logger.exception("Unexpected error in email future")

    # ---- Save updated Excel ----
    try:
        df.to_excel(EXCEL_FILE, index=False)
        logger.info("Saved updated Excel status: %s", EXCEL_FILE)
    except Exception as e:
        logger.error("Failed to save Excel: %s", e)

    elapsed = time.perf_counter() - start_time
    print("\n===== Summary =====")
    print(f"Total rows: {total}")
    print(f"Sent: {summary['sent']}")
    print(f"Already sent: {summary['already_sent']}")
    print(f"Skipped: {summary['skipped']}")
    print(f"Invalid emails: {summary['invalid_email']}")
    print(f"Missing PDFs: {summary['missing_file']}")
    print(f"Failed: {summary['failed']}")
    print(f"Time: {elapsed:.1f}s ({elapsed / max(summary['sent'], 1):.2f}s per email)")
    logger.info("Send summary: %s (%.1fs)", summary, elapsed)


if __name__ == "__main__":
    main()
