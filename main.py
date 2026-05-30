import os
import smtplib
import pandas as pd

from datetime import datetime
from dotenv import load_dotenv

from email.message import EmailMessage

from reportlab.pdfgen import canvas

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter

# ==========================================
# LOAD ENV VARIABLES
# ==========================================

load_dotenv()

SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = int(os.getenv("SMTP_PORT"))
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

# ==========================================
# COMPANY DETAILS
# ==========================================

COMPANY_NAME = "SLN Solution LLC"

CONTACT_EMAIL = "srikanth@sln-solutions.com"

CONTACT_PHONE = "+1 804-304-6616"

# ==========================================
# LOGO PATH
# ==========================================

LOGO_PATH = "assets/sln_logo.png"

# ==========================================
# READ EXCEL FILE
# ==========================================

df = pd.read_excel("excel/Invoice_Data.xlsx")

# ==========================================
# CREATE STORAGE
# ==========================================

os.makedirs("invoices", exist_ok=True)

# ==========================================
# INVOICE COUNTER
# ==========================================

invoice_counter = 1

# ==========================================
# SPLIT WEEKLY HOURS
# ==========================================

def split_hours(total_hours):

    weekly_hours = []

    while total_hours > 40:

        weekly_hours.append(40)

        total_hours -= 40

    if total_hours > 0:

        weekly_hours.append(total_hours)

    return weekly_hours

# ==========================================
# PROCESS EACH ROW
# ==========================================

for index, row in df.iterrows():

    # ==========================================
    # BASIC DATA
    # ==========================================

    resource_name = str(
        row[' Resource Name']
    ).strip()

    vendor_name = str(
        row['Vendor Name']
    ).strip()

    vendor_email = str(
        row['Vendor Email']
    ).strip()

    month = str(
        row['Month']
    ).strip()

    total_hours = int(
        row['No of Hours ']
    )

    bill_rate = float(
        row['Bill rate']
    )

    invoice_amount = float(
        row['Invoice Amount']
    )

    vendor_address = str(
        row['Vendor address']
    ).strip()

    # ==========================================
    # INVOICE NUMBER
    # ==========================================

    invoice_number = f"SLN{invoice_counter:05d}"

    invoice_counter += 1

    # ==========================================
    # WEEKLY HOURS SPLIT
    # ==========================================

    weekly_hours = split_hours(total_hours)

    # ==========================================
    # CREATE YEAR/MONTH FOLDER
    # ==========================================

    current_year = datetime.now().strftime("%Y")

    invoice_folder = os.path.join(
        "invoices",
        current_year,
        month
    )

    os.makedirs(invoice_folder, exist_ok=True)

    # ==========================================
    # PDF FILE NAME
    # ==========================================

    
    
    pdf_filename = (
    f"Invoice_SLN_"
    f"{vendor_name}_"
    f"{resource_name}_"
    f"{month}_"
    f"{index}.pdf"
   )

    pdf_file = os.path.join(
        invoice_folder,
        pdf_filename
    )

    # ==========================================
    # CREATE PDF
    # ==========================================

    c = canvas.Canvas(
        pdf_file,
        pagesize=letter
    )

    # ==========================================
    # LOGO
    # ==========================================

    c.drawImage(
        LOGO_PATH,
        235,
        720,
        width=160,
        height=70,
        preserveAspectRatio=True,
        mask='auto'
    )

    # ==========================================
    # INVOICE TITLE
    # ==========================================

    c.setFont("Helvetica-Bold", 22)

    c.drawString(
        235,
        665,
        "INVOICE"
    )

    # ==========================================
    # INVOICE INFO
    # ==========================================

    c.setFont("Helvetica-Bold", 10)

    c.drawString(
        430,
        645,
        f"Invoice #: {invoice_number}"
    )

    c.drawString(
        430,
        630,
        f"Date: {datetime.now().strftime('%B %d, %Y')}"
    )

    # ==========================================
    # COMPANY DETAILS
    # ==========================================

    c.setFont("Helvetica-Bold", 10)

    c.drawString(
        50,
        630,
        COMPANY_NAME
    )

    c.setFont("Helvetica", 9)

    c.drawString(
        50,
        615,
        "470 Olde Worthington Rd Suite 200"
    )

    c.drawString(
        50,
        602,
        "Westerville OH 43082"
    )

    # ==========================================
    # TO / SHIP TO BOXES
    # ==========================================

    c.rect(40, 480, 250, 110)

    c.rect(290, 480, 250, 110)

    # LEFT BOX

    c.setFont("Helvetica-Bold", 10)

    c.drawString(
        45,
        575,
        "To:"
    )

    c.setFont("Helvetica-Bold", 11)

    c.drawString(
        45,
        550,
        vendor_name
    )

    c.setFont("Helvetica", 10)

    c.drawString(
        45,
        535,
        vendor_address
    )

    # RIGHT BOX

    c.setFont("Helvetica-Bold", 10)

    c.drawString(
        295,
        575,
        "Ship To:"
    )

    c.setFont("Helvetica-Bold", 11)

    c.drawString(
        295,
        550,
        vendor_name
    )

    c.setFont("Helvetica", 10)

    c.drawString(
        295,
        535,
        vendor_address
    )

    # ==========================================
    # HEADER TABLE
    # ==========================================

    table_y = 445

    col0 = 40
    col1 = 230
    col2 = 290
    col3 = 360
    col4 = 480
    col5 = 540

    table_height = 35

    header_y = table_y + 24
    value_y = table_y + 8

    c.rect(col0, table_y, col5 - col0, table_height)

    c.line(col1, table_y, col1, table_y + table_height)
    c.line(col2, table_y, col2, table_y + table_height)
    c.line(col3, table_y, col3, table_y + table_height)
    c.line(col4, table_y, col4, table_y + table_height)

    # HEADERS

    c.setFont("Helvetica-Bold", 7)

    c.drawCentredString(
        (col0 + col1) / 2.0,
        header_y,
        "PERIOD OF SERVICES"
    )

    c.drawCentredString(
        (col1 + col2) / 2.0,
        header_y,
        "P.O. NUMBER"
    )

    c.drawCentredString(
        (col2 + col3) / 2.0,
        header_y,
        "REQUISITIONER"
    )

    c.drawCentredString(
        (col3 + col4) / 2.0,
        header_y,
        "SHIPPED VIA EMAIL TO"
    )

    c.drawCentredString(
        (col4 + col5) / 2.0,
        header_y,
        "TERMS"
    )

    # VALUES

    c.setFont("Helvetica", 7)

    c.drawCentredString(
        (col0 + col1) / 2.0,
        value_y,
        f"{month.upper()} {total_hours} HRS FOR ${bill_rate:,.0f}"
    )

    c.drawCentredString(
        (col2 + col3) / 2.0,
        value_y,
        resource_name[:12]
    )

    email_text = vendor_email[:25]

    c.drawCentredString(
        (col3 + col4) / 2.0,
        value_y,
        email_text
    )

    c.drawCentredString(
        (col4 + col5) / 2.0,
        value_y,
        "Due on receipt"
    )

    # ==========================================
    # MAIN HOURS TABLE
    # ==========================================

    main_y = 315

    c.rect(40, main_y, 500, 110)

    c.line(40, main_y + 80, 540, main_y + 80)

    c.line(200, main_y, 200, main_y + 110)

    c.line(430, main_y, 430, main_y + 110)

    # TABLE HEADERS

    c.setFont("Helvetica-Bold", 9)

    c.drawCentredString(
        120,
        main_y + 90,
        "QUANTITY"
    )

    c.drawCentredString(
        315,
        main_y + 90,
        "Project Details"
    )

    c.drawCentredString(
        485,
        main_y + 90,
        "TOTAL"
    )

    # ==========================================
    # WEEKLY HOURS
    # ==========================================

    c.setFont("Helvetica-Bold", 8)

    week_y = main_y + 65

    for i, hours in enumerate(weekly_hours):

        week_text = (
            f"WEEK ENDING 04/{(i+1)*5:02d} – {hours} HRS."
        )

        c.drawString(
            60,
            week_y,
            week_text
        )

        week_y -= 15

    # ==========================================
    # PROJECT DETAILS
    # ==========================================

    c.setFont("Helvetica-Bold", 10)

    c.drawCentredString(
        315,
        main_y + 45,
        f"IT SERVICES FOR {resource_name.upper()}"
    )

    c.drawCentredString(
        315,
        main_y + 30,
        f"{month.upper()} MONTH"
    )

    # ==========================================
    # TOTAL AMOUNT
    # ==========================================

    c.setFont("Helvetica-Bold", 16)

    c.drawCentredString(
        485,
        main_y + 40,
        f"${invoice_amount:,.2f}"
    )

    # ==========================================
    # PAYMENT SECTION
    # ==========================================

    payment_y = 170

    c.rect(40, payment_y, 500, 120)

    c.line(
        190,
        payment_y,
        190,
        payment_y + 120
    )

    c.line(
        360,
        payment_y,
        360,
        payment_y + 120
    )

    # LEFT SIDE

    c.setFont("Helvetica-Bold", 8)

    c.drawString(
        45,
        270,
        "Check Details:"
    )

    c.setFont("Helvetica", 8)

    c.drawString(
        45,
        250,
        COMPANY_NAME
    )

    c.drawString(
        45,
        238,
        "470 Olde Worthington Rd"
    )

    c.drawString(
        45,
        226,
        "Suite 200 Westerville OH 43082"
    )

    # MIDDLE SIDE (WIRE)

    c.setFont("Helvetica-Bold", 8)

    c.drawString(
        195,
        270,
        "Wire"
    )

    c.setFont("Helvetica", 8)

    c.drawString(
        195,
        250,
        "Payee Name: SLN Solution LLC"
    )

    c.drawString(
        195,
        238,
        "Account Number: 374006933920"
    )

    c.drawString(
        195,
        226,
        "Routing Number: 071214579"
    )

    c.drawString(
        195,
        214,
        "Wire: 026009593"
    )

    # SUBTOTAL

    c.setFont("Helvetica-Bold", 11)

    c.drawString(
        370,
        250,
        "SUBTOTAL"
    )

    c.drawRightString(
        530,
        250,
        f"${invoice_amount:,.2f}"
    )

    # SALES TAX

    c.drawString(
        370,
        225,
        "SALES TAX"
    )

    c.drawRightString(
        530,
        225,
        "$0.00"
    )

    # TOTAL DUE

    c.setFont("Helvetica-Bold", 14)

    c.drawString(
        370,
        190,
        "TOTAL DUE"
    )

    c.drawRightString(
        530,
        190,
        f"${invoice_amount:,.2f}"
    )

    # ==========================================
    # FOOTER
    # ==========================================

c.line(40, 150, 540, 150)

c.setFont("Helvetica", 7)

c.drawString(
    45,
    145,
    "If you have any questions concerning this invoice,"
)

c.drawString(
    45,
    135,
    f"please contact {CONTACT_EMAIL}"
)

c.drawString(
    45,
    125,
    f"Phone: {CONTACT_PHONE}"
)

# ==========================================
# THANK YOU
# ==========================================

c.setFont("Helvetica-Bold", 18)

c.setFillColor(colors.darkblue)

c.drawCentredString(
    300,
    55,
    "Thank you for your business!"
)

# ==========================================
# SAVE PDF
# ==========================================

c.save()

if vendor_email == "" or vendor_email.lower() == "nan":

        print("Skipped Empty Vendor Email")

        pass
try:

        msg = EmailMessage()

        msg['Subject'] = (
            f"Invoice {invoice_number} - {month}"
        )

        msg['From'] = EMAIL_ADDRESS

        msg['To'] = vendor_email

        msg.set_content(f"""
Hello,

Please find attached invoice for {month}.

Invoice Number: {invoice_number}

Regards,
SLN Solution LLC
""")

        # ATTACH PDF

        with open(pdf_file, "rb") as f:

            file_data = f.read()

            file_name = os.path.basename(pdf_file)

        msg.add_attachment(
            file_data,
            maintype="application",
            subtype="pdf",
            filename=file_name
        )

        # SEND EMAIL

        with smtplib.SMTP_SSL(
            SMTP_SERVER,
            SMTP_PORT
        ) as smtp:

            smtp.login(
                EMAIL_ADDRESS,
                EMAIL_PASSWORD
            )

            smtp.send_message(msg)

        # SUCCESS LOG

        with open("logs.txt", "a") as log:

            log.write(
                f"\n[{datetime.now()}] "
                f"SUCCESS - {vendor_email} - "
                f"{invoice_number}"
            )

        print(f"Email Sent To: {vendor_email}")

except Exception as e:

        print(f"Failed Sending To: {vendor_email}")

        print(e)

        with open("logs.txt", "a") as log:

            log.write(
                f"\n[{datetime.now()}] "
                f"FAILED - {vendor_email} - "
                f"{str(e)}"
            )

print("\n================================")

print("Professional Invoice Generated")

print(f"Invoice Number: {invoice_number}")

print(f"PDF Saved At: {pdf_file}")

print("================================")