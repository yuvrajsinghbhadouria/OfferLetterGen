import os
from pathlib import Path

from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

# =====================================================
# EMAIL CONFIGURATION
# =====================================================

GMAIL_EMAIL = os.getenv("GMAIL_EMAIL", "connectbiopay@gmail.com")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD", "lbvu loor dwvd khng ")

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

Name
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

Name
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
    "Biopay_Agreement.docx"
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
# CREATE FOLDERS IF MISSING
# =====================================================

os.makedirs(DATA_FOLDER, exist_ok=True)
os.makedirs(TEMPLATE_FOLDER, exist_ok=True)
os.makedirs(GENERATED_FOLDER, exist_ok=True)
os.makedirs(LOG_FOLDER, exist_ok=True)