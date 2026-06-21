import os
from pathlib import Path

from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

# =====================================================
# EMAIL / SMTP CONFIGURATION
# =====================================================
# Works with any provider: Gmail, Outlook, Yahoo, custom domains, etc.
# Set SMTP_HOST and SMTP_PORT in .env to override auto-detection.
#
# Common SMTP settings:
#   Gmail:        smtp.gmail.com        port 465 (SSL) or 587 (TLS)
#   Outlook:      smtp.office365.com    port 587 (TLS)
#   Yahoo:        smtp.mail.yahoo.com   port 465 (SSL)
#   Zoho:         smtp.zoho.com         port 465 (SSL)
#   Custom:       your-smtp-server.com  port 465 or 587
# =====================================================

# Sender credentials — MUST be set in .env, no hardcoded fallbacks
SENDER_EMAIL = os.getenv("SENDER_EMAIL") or os.getenv("GMAIL_EMAIL", "")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD") or os.getenv("GMAIL_APP_PASSWORD", "")

if not SENDER_EMAIL:
    raise EnvironmentError(
        "SENDER_EMAIL is not set. Create a .env file with SENDER_EMAIL=you@domain.com"
    )
if not SENDER_PASSWORD:
    raise EnvironmentError(
        "SENDER_PASSWORD is not set. Create a .env file with SENDER_PASSWORD=your_app_password"
    )

# Auto-detect SMTP host from email domain if not explicitly set
_SMTP_DEFAULTS = {
    "gmail.com":       ("smtp.gmail.com",        465, False),
    "outlook.com":     ("smtp.office365.com",    587, True),
    "hotmail.com":     ("smtp.office365.com",    587, True),
    "live.com":        ("smtp.office365.com",    587, True),
    "yahoo.com":       ("smtp.mail.yahoo.com",   465, False),
    "yahoo.co.in":     ("smtp.mail.yahoo.com",   465, False),
    "zoho.com":        ("smtp.zoho.com",         465, False),
    "zoho.in":         ("smtp.zoho.in",          465, False),
}

def _detect_smtp_settings(email: str):
    domain = email.rsplit("@", 1)[-1].lower().strip()
    return _SMTP_DEFAULTS.get(domain, (f"smtp.{domain}", 587, True))

_auto_host, _auto_port, _auto_tls = _detect_smtp_settings(SENDER_EMAIL)

SMTP_HOST = os.getenv("SMTP_HOST", _auto_host)
SMTP_PORT = int(os.getenv("SMTP_PORT", str(_auto_port)))
SMTP_USE_TLS = os.getenv("SMTP_USE_TLS", str(_auto_tls)).lower() in ("true", "1", "yes")

# Backward compatibility aliases
GMAIL_EMAIL = SENDER_EMAIL
GMAIL_APP_PASSWORD = SENDER_PASSWORD

# =====================================================
# COMPANY DETAILS
# =====================================================

COMPANY_NAME = os.getenv("COMPANY_NAME", "Company Name")
EMAIL_SUBJECT = os.getenv(
    "EMAIL_SUBJECT",
    "Offer Letter {role} Summer Internship | Company Name"
)

EMAIL_BODY_TECHNICAL = os.getenv(
    "EMAIL_BODY_TECHNICAL",
    """
Dear {name},

Greetings from {company_name}!

Congratulations on your selection as a Technical  Intern.

Internship ID: {internship_id}

Your Role: {role}

Duration: {duration}

Duties & Responsibilities:
{duty}

Please find attached your Offer Letter, SOP, and Biopay Agreement for detailed terms and conditions.

Kindly review the documents and confirm your acceptance by replying to this email and Submit the Agreement Attached with your Photograph and Resume within next 7 Days

We are excited to welcome you to our technical team and look forward to working with you.

Best Regards,

Yuvraj Singh
Founder & CEO
{company_name}
"""
)

EMAIL_BODY_NON_TECHNICAL = os.getenv(
    "EMAIL_BODY_NON_TECHNICAL",
    """
Dear {name},

Greetings from {company_name}!

Congratulations on your selection as a Non-Technical Intern.

Internship ID: {internship_id}

Your Role: {role}

Duration: {duration}

Duties & Responsibilities:
{duty}

Please find attached your Offer Letter, SOP, and Biopay Agreement for detailed terms and conditions.

Kindly review the documents and confirm your acceptance by replying to this email and Submit the Agreement Attached with your Photograph and Resume within next 7 Days. 

We are excited to have you on our team and look forward to working with you.

Best Regards,

Yuvraj Singh
Founder & CEO
{company_name}
"""
)

if isinstance(EMAIL_BODY_TECHNICAL, str):
    EMAIL_BODY_TECHNICAL = EMAIL_BODY_TECHNICAL.replace("\\n", "\n")
if isinstance(EMAIL_BODY_NON_TECHNICAL, str):
    EMAIL_BODY_NON_TECHNICAL = EMAIL_BODY_NON_TECHNICAL.replace("\\n", "\n")

# =====================================================
# FOLDER STRUCTURE
# =====================================================

DATA_FOLDER = str(BASE_DIR / os.getenv("DATA_FOLDER", "Data"))
TEMPLATE_FOLDER = str(BASE_DIR / os.getenv("TEMPLATE_FOLDER", "Template"))
GENERATED_FOLDER = str(BASE_DIR / os.getenv("GENERATED_FOLDER", "Generated"))
LOG_FOLDER = str(BASE_DIR / os.getenv("LOG_FOLDER", "Logs"))

# =====================================================
# FILES
# =====================================================

EXCEL_FILE = os.path.join(
    DATA_FOLDER,
    "interns.xlsx"
)

TEMPLATE_FILE = os.path.join(
    TEMPLATE_FOLDER,
    "Offer_Letter_Template.docx"
)

SOP_FILE = os.path.join(
    TEMPLATE_FOLDER,
    "SOP.pdf"
)

BIOPAY_AGREEMENT_FILE = os.path.join(
    TEMPLATE_FOLDER,
    "BIOPAY AGREEMENT.pdf"
)

DURATION_PLACEHOLDER = os.getenv("DURATION_PLACEHOLDER", "[Duration]")

# =====================================================
# EXCEL COLUMN NAMES
# =====================================================

NAME_COLUMN = "Name"
EMAIL_COLUMN = "Student mail ID"
STATUS_COLUMN = "Status"
ROLES_COLUMN = "Roles"
DUTY_COLUMN = "Duty"
COMPANY_COLUMN = "Company Name"
DURATION_COLUMN = "Duration"
INTERN_ID_COLUMN = "Intern Id"
MODE_COLUMN = "Mode"
COMMITMENT_COLUMN = "Commitment"
STIPEND_COLUMN = "Stipend"

# =====================================================
# WORD PLACEHOLDERS
# =====================================================

NAME_PLACEHOLDER = os.getenv("NAME_PLACEHOLDER", "[Intern Name]")

DATE_PLACEHOLDER = os.getenv("DATE_PLACEHOLDER", "[Date]")

# =====================================================
# LOGGING SETTINGS
# =====================================================

LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()
LOG_FILE = str(Path(LOG_FOLDER) / "offer_letters.log")
DATE_FORMAT = os.getenv("DATE_FORMAT", "%d %B %Y")

# =====================================================
# PERFORMANCE / PARALLELISM SETTINGS
# =====================================================

MAX_DOCX_WORKERS = int(os.getenv("MAX_DOCX_WORKERS", "4"))
MAX_EMAIL_WORKERS = int(os.getenv("MAX_EMAIL_WORKERS", "4"))

# =====================================================
# CREATE FOLDERS IF MISSING
# =====================================================

os.makedirs(DATA_FOLDER, exist_ok=True)
os.makedirs(TEMPLATE_FOLDER, exist_ok=True)
os.makedirs(GENERATED_FOLDER, exist_ok=True)
os.makedirs(LOG_FOLDER, exist_ok=True)