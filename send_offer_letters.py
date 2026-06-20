import os
import time
import smtplib
import pandas as pd
from email.message import EmailMessage

from config import *

# =====================================================
# READ EXCEL
# =====================================================

df = pd.read_excel(EXCEL_FILE)
df.columns = df.columns.str.strip()

if STATUS_COLUMN not in df.columns:
    df[STATUS_COLUMN] = ""
# Make Status a text column
df[STATUS_COLUMN] = df[STATUS_COLUMN].fillna("").astype(str)
total = len(df)
sent = 0
failed = 0
skipped = 0

print("\n==============================")
print(" CONNECT BIOPAY MAIL SENDER")
print("==============================\n")

# =====================================================
# SEND EMAILS
# =====================================================

with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:

    smtp.login(GMAIL_EMAIL, GMAIL_APP_PASSWORD)

    for index, row in df.iterrows():

        name = str(row[NAME_COLUMN]).strip()
        email = str(row[EMAIL_COLUMN]).strip()

        if str(row[STATUS_COLUMN]).strip().lower() == "sent":
            print(f"⏭ Skipping {name}")
            skipped += 1
            continue

        pdf_path = os.path.join(
            GENERATED_FOLDER,
            f"{name}.pdf"
        )

        if not os.path.exists(pdf_path):
            print(f"❌ PDF Missing : {name}")
            failed += 1
            continue

        try:

            msg = EmailMessage()

            msg["Subject"] = EMAIL_SUBJECT
            msg["From"] = GMAIL_EMAIL
            msg["To"] = email

            msg.set_content(
                EMAIL_BODY.format(name=name)
            )

            with open(pdf_path, "rb") as f:

                msg.add_attachment(
                    f.read(),
                    maintype="application",
                    subtype="pdf",
                    filename=os.path.basename(pdf_path)
                )

            smtp.send_message(msg)

            df.at[index, STATUS_COLUMN] = "Sent"

            sent += 1

            print(f"✅ {sent}/{total}  {name}")

            # Wait 2 seconds to avoid Gmail rate limits
            time.sleep(2)

        except Exception as e:

            failed += 1

            print(f"❌ Failed : {name}")
            print(e)

# =====================================================
# SAVE EXCEL
# =====================================================

df.to_excel(EXCEL_FILE, index=False)

print("\n=================================")
print("          COMPLETED")
print("=================================")
print(f"Sent     : {sent}")
print(f"Skipped  : {skipped}")
print(f"Failed   : {failed}")
print("=================================")