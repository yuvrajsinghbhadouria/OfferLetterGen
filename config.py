import os

# =====================================================
# EMAIL CONFIGURATION
# =====================================================

GMAIL_EMAIL = "yourgmail@gmail.com"
GMAIL_APP_PASSWORD = "16 Charachter Passwprd"

# =====================================================
# COMPANY DETAILS
# =====================================================

COMPANY_NAME = "Company Name"

EMAIL_SUBJECT = "Offer Letter – Technical Summer Internship | Company Name"

EMAIL_BODY = """
Dear {name},

Greetings from Company Name!

Congratulations on your selection as a Technical Summer Intern.

Please find attached your Offer Letter.

Kindly review the document and confirm your acceptance by replying to this email.

We are excited to welcome you to Company Name and look forward to working with you.

Best Regards,

Name
Founder & CEO
Company Name
"""

# =====================================================
# FOLDER STRUCTURE
# =====================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_FOLDER = os.path.join(BASE_DIR, "Data")

TEMPLATE_FOLDER = os.path.join(BASE_DIR, "Template")

GENERATED_FOLDER = os.path.join(BASE_DIR, "Generated")

LOG_FOLDER = os.path.join(BASE_DIR, "Logs")

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

# =====================================================
# EXCEL COLUMN NAMES
# =====================================================

NAME_COLUMN = "Name"

EMAIL_COLUMN = "Student mail ID"

STATUS_COLUMN = "Status"

# =====================================================
# WORD PLACEHOLDERS
# =====================================================

NAME_PLACEHOLDER = "[Intern Name]"

DATE_PLACEHOLDER = "[Date]"

# =====================================================
# CREATE FOLDERS IF MISSING
# =====================================================

os.makedirs(DATA_FOLDER, exist_ok=True)
os.makedirs(TEMPLATE_FOLDER, exist_ok=True)
os.makedirs(GENERATED_FOLDER, exist_ok=True)
os.makedirs(LOG_FOLDER, exist_ok=True)