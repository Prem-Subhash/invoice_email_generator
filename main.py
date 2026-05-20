import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import os
import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv
from datetime import datetime

# ==========================================
# LOAD ENVIRONMENT VARIABLES
# ==========================================

load_dotenv()

SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = int(os.getenv("SMTP_PORT"))
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

# ==========================================
# READ EXCEL FILE
# ==========================================

file_path = "excel/Invoice_Data.xlsx"

df = pd.read_excel(file_path)

# ==========================================
# CREATE MAIN INVOICE FOLDER
# ==========================================

os.makedirs("invoices", exist_ok=True)

# ==========================================
# PROCESS EACH INVOICE
# ==========================================

for index, row in df.iterrows():

    # ==========================================
    # CLEAN DATA
    # ==========================================

    resource_name = str(row[' Resource Name']).strip()

    month = str(row['Month']).strip()

    vendor_email = str(row['Vendor Email']).strip()

    current_date = datetime.now().strftime("%Y%m%d")

    current_year = datetime.now().strftime("%Y")

    # ==========================================
    # CREATE YEAR / MONTH FOLDER
    # ==========================================

    month_folder = month.replace(" ", "_")

    invoice_folder = os.path.join(
        "invoices",
        current_year,
        month_folder
    )

    os.makedirs(invoice_folder, exist_ok=True)

    # ==========================================
    # GENERATE PDF FILE NAME
    # ==========================================

    pdf_filename = (
        f"Invoice_{resource_name}_{month}_{current_date}.pdf"
    )

    pdf_file = os.path.join(
        invoice_folder,
        pdf_filename
    )

    # ==========================================
    # CREATE PDF
    # ==========================================

    c = canvas.Canvas(pdf_file, pagesize=letter)

    # Title
    c.setFont("Helvetica-Bold", 18)

    c.drawString(220, 750, "INVOICE")

    # Invoice Details
    c.setFont("Helvetica", 12)

    c.drawString(50, 700, f"Company Name: {row['Company Name']}")

    c.drawString(50, 680, f"Resource Name: {resource_name}")

    c.drawString(50, 660, f"Vendor Name: {row['Vendor Name']}")

    c.drawString(50, 640, f"Month: {month}")

    c.drawString(50, 620, f"No of Hours: {row['No of Hours ']}")

    c.drawString(50, 600, f"Bill Rate: ${row['Bill rate']}")

    c.drawString(50, 580, f"Invoice Amount: ${row['Invoice Amount']}")

    # Company Address
    c.drawString(50, 540, "Company Address:")

    c.drawString(70, 520, str(row['Company Address']))

    # Vendor Address
    c.drawString(50, 480, "Vendor Address:")

    c.drawString(70, 460, str(row['Vendor address']))

    # Footer
    c.drawString(50, 400, "Thank you for your business!")

    # Save PDF
    c.save()

    print("\n===================================")

    print("Invoice Generated Successfully")

    print(f"PDF Stored At: {pdf_file}")

    print("===================================")

    # ==========================================
    # VALIDATE EMAIL
    # ==========================================

    if vendor_email == "" or vendor_email.lower() == "nan":

        print("Skipped: Empty Vendor Email")

        with open("logs.txt", "a") as log:

            log.write(
                f"""

==================================================
TIME: {datetime.now()}
STATUS: SKIPPED
REASON: Empty Vendor Email
RESOURCE: {resource_name}
MONTH: {month}
PDF: {pdf_filename}
STORAGE: {pdf_file}
==================================================

"""
            )

        continue

    # ==========================================
    # CREATE EMAIL
    # ==========================================

    msg = EmailMessage()

    msg['Subject'] = f"Invoice for {month}"

    msg['From'] = EMAIL_ADDRESS

    msg['To'] = vendor_email

    msg.set_content(f"""
Hello,

Please find attached invoice for {month}.

Regards,
Accounts Team
""")

    # ==========================================
    # ATTACH PDF
    # ==========================================

    with open(pdf_file, "rb") as f:

        file_data = f.read()

    msg.add_attachment(
        file_data,
        maintype="application",
        subtype="pdf",
        filename=pdf_filename
    )

    # ==========================================
    # SEND EMAIL
    # ==========================================

    try:

        with smtplib.SMTP_SSL(
            SMTP_SERVER,
            SMTP_PORT
        ) as smtp:

            smtp.login(
                EMAIL_ADDRESS,
                EMAIL_PASSWORD
            )

            smtp.send_message(msg)

        print(f"Email Sent Successfully To: {vendor_email}")

        # ==========================================
        # SUCCESS LOG
        # ==========================================

        with open("logs.txt", "a") as log:

            log.write(
                f"""

==================================================
TIME: {datetime.now()}
STATUS: SUCCESS
TO: {vendor_email}
RESOURCE: {resource_name}
MONTH: {month}
AMOUNT: ${row['Invoice Amount']}
PDF: {pdf_filename}
STORAGE: {pdf_file}
==================================================

"""
            )

    except Exception as e:

        print(f"Failed To Send Email To: {vendor_email}")

        print(e)

        # ==========================================
        # FAILURE LOG
        # ==========================================

        with open("logs.txt", "a") as log:

            log.write(
                f"""

==================================================
TIME: {datetime.now()}
STATUS: FAILED
TO: {vendor_email}
RESOURCE: {resource_name}
MONTH: {month}
PDF: {pdf_filename}
STORAGE: {pdf_file}
ERROR: {str(e)}
==================================================

"""
            )

# ==========================================
# FINAL MESSAGE
# ==========================================

print("\nAll invoices processed successfully!")