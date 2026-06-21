import os
import time
import tempfile
from datetime import datetime
from pathlib import Path
from concurrent.futures import ProcessPoolExecutor, as_completed

import pandas as pd
from tqdm import tqdm

from config import *
from utils import build_safe_attachment_name, generate_internship_id, setup_logging, unique_filepath, validate_email

logger = setup_logging("generate_offer_letters")
OUTPUT_FOLDER = Path(GENERATED_FOLDER)


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


# ─────────────────────────────────────────────────────
# PIPELINE WORKER — runs in its own process
#
#   Render DOCX in memory
#        ↓
#   Hand to Word via invisible temp file in %TEMP%
#        ↓
#   Export PDF to Generated/
#        ↓
#   Temp DOCX deleted instantly — never kept
# ─────────────────────────────────────────────────────

def _pipeline_worker(file_jobs):
    """Each PROCESS: render in memory → Word → PDF. No DOCX kept."""
    import pythoncom
    import win32com.client as win32
    from docxtpl import DocxTemplate
    import time as _time

    pythoncom.CoInitialize()
    word = None
    results = []
    worker_start = _time.perf_counter()

    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False

        for template_path, context, pdf_path_str, row_index, intern_id in file_jobs:
            file_start = _time.perf_counter()
            tmp_path = None
            success = False
            try:
                # 1) Render DOCX entirely in memory
                doc = DocxTemplate(template_path)
                doc.render(context)

                # 2) Write to invisible system temp (not project folder)
                with tempfile.NamedTemporaryFile(
                    suffix=".docx", delete=False, dir=tempfile.gettempdir()
                ) as tmp:
                    doc.save(tmp)
                    tmp_path = tmp.name

                # 3) Open in Word → export PDF
                word_doc = word.Documents.Open(tmp_path)
                word_doc.SaveAs(pdf_path_str, FileFormat=17)
                word_doc.Close(SaveChanges=False)
                success = True

            except Exception:
                pass
            finally:
                # 4) Delete temp DOCX instantly — never kept
                if tmp_path:
                    try:
                        os.unlink(tmp_path)
                    except Exception:
                        pass

            file_elapsed = _time.perf_counter() - file_start
            results.append((row_index, intern_id, pdf_path_str, success, file_elapsed))

    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    worker_elapsed = _time.perf_counter() - worker_start
    return results, worker_elapsed


def main():
    start_time = time.perf_counter()
    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

    df = load_data()
    total = len(df)
    generated = 0
    skipped = 0
    failed = 0

    # Validate template ONCE
    if not os.path.exists(TEMPLATE_FILE):
        raise FileNotFoundError(f"Template file not found: {TEMPLATE_FILE}")

    template_path = str(Path(TEMPLATE_FILE).resolve())

    # ── Prepare all jobs (fast, no I/O) ──
    # Each job: (template_path, context, pdf_path_str, row_index, intern_id)
    all_jobs = []

    for index, row in df.iterrows():
        name = str(row[NAME_COLUMN]).strip()
        email = str(row.get(EMAIL_COLUMN, "") or "").strip()
        company_name = str(row.get(COMPANY_COLUMN, "") or "").strip()
        role = str(row.get(ROLES_COLUMN, "") or "").strip()

        if not name:
            skipped += 1
            logger.warning("Skipping row %s: missing %s", index + 1, NAME_COLUMN)
            continue

        if email and not validate_email(email):
            logger.warning("Invalid email for %s: %s", name, email)

        intern_id = str(row.get(INTERN_ID_COLUMN, "") or "").strip()
        if not intern_id:
            intern_id = generate_internship_id(company_name, name)
            df.at[index, INTERN_ID_COLUMN] = intern_id

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

        base_name = build_safe_attachment_name(name, email)
        if not base_name:
            base_name = f"intern-{index + 1}"

        pdf_path = unique_filepath(OUTPUT_FOLDER / f"{base_name}.pdf")

        all_jobs.append((template_path, context, str(pdf_path), index, intern_id))

    if not all_jobs:
        logger.info("Nothing to generate.")
        return

    # ══════════════════════════════════════════════════
    # ALL FILES SIMULTANEOUSLY — 1 process per file
    # Render in memory → Word → PDF → no DOCX kept
    # ══════════════════════════════════════════════════
    num_workers = len(all_jobs)
    logger.info(
        "Generating %s PDFs SIMULTANEOUSLY (%s processes, 1 file each)...",
        len(all_jobs), num_workers,
    )

    with ProcessPoolExecutor(max_workers=num_workers) as pool:
        job_futures = {pool.submit(_pipeline_worker, [job]): i for i, job in enumerate(all_jobs)}

        for future in tqdm(
            as_completed(job_futures), total=len(job_futures),
            desc="Generating PDFs", unit="file"
        ):
            worker_idx = job_futures[future]
            try:
                results, worker_elapsed = future.result()
                for row_index, intern_id, pdf_path_str, success, file_elapsed in results:
                    if success:
                        generated += 1
                        logger.info(
                            "[File %s] ✓ %s | ID: %s | %.1fs",
                            worker_idx + 1, Path(pdf_path_str).name, intern_id, file_elapsed,
                        )
                    else:
                        failed += 1
                        logger.error("[File %s] ✗ Failed row %s", worker_idx + 1, row_index + 1)
                logger.info("[File %s] process time: %.1fs", worker_idx + 1, worker_elapsed)
            except Exception:
                failed += 1
                logger.exception("[File %s] process crashed", worker_idx + 1)

    # ── Save updated Excel ──
    try:
        df.to_excel(EXCEL_FILE, index=False)
        logger.info("Updated Excel file with generated IDs")
    except Exception as e:
        logger.error("Failed to update Excel file: %s", e)

    elapsed = time.perf_counter() - start_time
    logger.info("Finished in %.1fs. Generated=%s, Skipped=%s, Failed=%s", elapsed, generated, skipped, failed)
    print("\n===== Summary =====")
    print(f"Total rows:  {total}")
    print(f"Generated:   {generated}")
    print(f"Skipped:     {skipped}")
    print(f"Failed:      {failed}")
    print(f"Workers:     {num_workers} simultaneous processes")
    print(f"Total time:  {elapsed:.1f}s ({elapsed / max(generated, 1):.2f}s per letter)")


if __name__ == "__main__":
    main()
