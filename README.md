# Offer Letter Automation

Generates personalized offer letters from a Word template, converts them to PDF, and sends them by email.

**Supported platform:** Windows (requires Microsoft Word for DOCX→PDF conversion via COM automation)

**Main scripts:**
- `generate_offer_letters.py` — render templates, save DOCX and PDF, and populate `Intern Id` in the Excel file.
- `send_offer_letters.py` — send generated PDFs (and related attachments) by Gmail SMTP and update the `Status` column.

**Config:**
- `config.py` centralizes paths and email templates. Default folders created at runtime: `Data/`, `Template/`, `Generated/`, `Logs/`.
- Copy `.env.example` to `.env` and fill values for `GMAIL_EMAIL`, `GMAIL_APP_PASSWORD`, and other overrides.

**Key config variables**
- `GMAIL_EMAIL`, `GMAIL_APP_PASSWORD` — Gmail sender and app password.
- `COMPANY_NAME` — company name used in email templates.
- `EMAIL_SUBJECT` — subject template (formatted with `{role}` and `{company_name}` when possible).
- `EMAIL_BODY_TECHNICAL`, `EMAIL_BODY_NON_TECHNICAL` — full email body templates used depending on the `Roles` column.
- `EXCEL_FILE` — default `Data/interns.xlsx`.
- `TEMPLATE_FILE`, `SOP_FILE`, `BIOPAY_AGREEMENT_FILE` — files under `Template/` used by the scripts.
- Excel column names (as used in code): `Name`, `Student mail ID`, `Status`, `Roles`, `Duty`, `Duration`, `Company Name`, `Intern Id`, `Mode`, `Commitment`, `Stipend`.

## Requirements

- Python 3.x on Windows
- Microsoft Word installed (for PDF conversion via `pywin32`)
- Gmail account with an app password (if using Gmail SMTP)

Install Python dependencies:

```powershell
python -m pip install -r requirements.txt
```

## Usage

1) Generate offer letters (creates DOCX + PDF and fills `Intern Id` in the Excel file):

```powershell
python generate_offer_letters.py
```

2) Send offer letters by email (attaches Offer Letter PDF, SOP, and Biopay Agreement PDF):

```powershell
python send_offer_letters.py
```

What `generate_offer_letters.py` does:
- Reads rows from `Data/interns.xlsx` and enforces required columns.
- Generates a deterministic `Intern Id` when missing and writes it back to the Excel file.
- Renders `Template/Offer_Letter_Template.docx` using `docxtpl` context values (see config placeholders).
- Saves a safe, unique DOCX filename and converts it to PDF using Word COM.

What `send_offer_letters.py` does:
- Reads `Data/interns.xlsx`, skips rows where `Status` is `Sent`.
- Validates email addresses and required fields (`Name`, `Roles`, `Duty`, `Duration`, `Company Name`).
- Looks up the generated Offer Letter PDF in `Generated/` (by safe filename), attaches it.
- Attaches `SOP.pdf` (if available) and a generated `Biopay_Agreement.pdf` (the DOCX template is rendered with `duration` and `date` and converted to PDF).
- Sends email via `smtp.gmail.com:465` using SSL and updates `Status` to `Sent` and saves `Intern Id`.

## Excel input format

The Excel file should include these columns (exact names used by the scripts):
- `Name`
- `Student mail ID`
- `Roles` (e.g. `Technical` or `Non-Technical` — picks email body template)
- `Duty`
- `Duration`
- `Company Name`
- `Status` (created if missing)
- `Intern Id` (generated if missing)

The scripts will create `Status` and `Intern Id` columns when absent and persist updates to the same Excel file.

## Logs

Log file is written to `Logs/offer_letters.log` and controlled via `LOG_LEVEL` in `.env` or `config.py`.

## Notes & troubleshooting

- The DOCX→PDF conversion uses the local Microsoft Word application via `pywin32`; it only works on Windows with Word installed.
- Email sending uses Gmail SSL on port 465 — create an app password for `GMAIL_APP_PASSWORD`.
- Template placeholders used by `docxtpl` include: `{name}`, `{company_name}`, `{role}`, `{date}`, `{duty}`, `{duration}`, `{mode}`, `{commitment}`, `{stipend}`, `{intern_id}`.
- If a PDF is missing for a row, the email will not be sent for that record and it will be logged.


