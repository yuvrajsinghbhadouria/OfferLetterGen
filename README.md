# Offer Letter Automation

This project generates personalized offer letters from a Word template and sends them by email using Gmail.

## Project Structure

- `config.py` - central configuration for email, files, folders, and placeholders.
- `generate_offer_letters.py` - reads intern data from `Data/interns.xlsx`, renders a Word template, saves each document as `.docx`, then converts it to `.pdf` using Microsoft Word.
- `send_offer_letters.py` - reads the same Excel file, sends each generated PDF by email, and updates the `Status` column to `Sent`.
- `Data/` - input data folder, expected to contain `interns.xlsx`.
- `Template/` - Word template folder, expected to contain `Offer_Letter_Template.docx`.
- `Generated/` - output folder for generated `.docx` and `.pdf` offer letters.
- `Logs/` - created automatically by `config.py` for future log use.

## Requirements

- Windows (required for Microsoft Word COM automation)
- Python 3.x
- Microsoft Word installed
- Gmail account with app password enabled

## Python Dependencies

Install required packages before running the scripts:

```powershell
python -m pip install -r requirements.txt
```

## Configuration

Copy `.env.example` to `.env` and fill in your values. The scripts will automatically load Gmail credentials, paths, logging settings, and email template text from `.env`.

If you prefer not to use `.env`, you can continue editing `config.py` directly, but `.env` is recommended for credentials and environment-specific overrides.

Update `config.py` with your organization details and file paths.

Main settings:

- `GMAIL_EMAIL` - sender Gmail address
- `GMAIL_APP_PASSWORD` - Gmail app password (16-character password)
- `COMPANY_NAME` - company name used in email content
- `EMAIL_SUBJECT` - subject line for outgoing emails
- `EMAIL_BODY` - email body template; uses `{name}` placeholder
- `EXCEL_FILE` - path to `Data/interns.xlsx`
- `TEMPLATE_FILE` - path to `Template/Offer_Letter_Template.docx`
- `GENERATED_FOLDER` - output folder for generated files
- `NAME_COLUMN`, `EMAIL_COLUMN`, `STATUS_COLUMN` - Excel column names
- `NAME_PLACEHOLDER`, `DATE_PLACEHOLDER` - template placeholders used by `docxtpl`

`config.py` also ensures the `Data/`, `Template/`, `Generated/`, and `Logs/` folders exist.

## Usage

### 1. Generate offer letters

Run:

```powershell
python generate_offer_letters.py
```

What it does:

- Reads intern names from `Data/interns.xlsx`
- Replaces template placeholders with intern name and current date
- Saves a `.docx` file for each intern in `Generated/`
- Converts each document to `.pdf`

### 2. Send offer letters by email

Run:

```powershell
python send_offer_letters.py
```

What it does:

- Reads `Data/interns.xlsx`
- Skips rows already marked `Sent` in the `Status` column
- Attaches the corresponding PDF from `Generated/`
- Sends email via Gmail SMTP
- Updates the Excel `Status` column to `Sent`

## Excel Input Format

The Excel file should include at least these columns:

- `Name`
- `Student mail ID`
- `Status`

The script will create the `Status` column if it does not exist.

## Notes

- The project relies on Microsoft Word automation (`win32com.client`). It will only work on Windows with Word installed.
- Gmail sending uses SSL on port `465`.
- If a PDF file is missing for a row, the email will not be sent for that record.
- The `Status` field prevents duplicate emails when the script is run multiple times.

## Troubleshooting

- Ensure the Word template placeholders match the names expected by `docxtpl`.
- Verify `Data/interns.xlsx` uses the exact column names configured in `config.py`.
- For Gmail, enable app passwords and verify the login credentials.
- If Word automation fails, confirm Microsoft Word is installed and accessible.
