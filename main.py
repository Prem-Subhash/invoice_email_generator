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

    invoice_amount =  round(
        total_hours * bill_rate,
        2
    )

    vendor_address = str(
        row['Vendor address']
    ).strip()

    # ==========================================
    # INVOICE NUMBER
    # ==========================================

    invoice_number = f"SLN{invoice_counter}"

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
    # DUPLICATE CHECK
    # ==========================================

    if os.path.exists(pdf_file):
        print(
            f"Skipped - Already Processed: {pdf_filename}"
        )
        
        continue

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
        width=140,
        height=60,
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

    c.rect(40, 490, 250, 100)

    c.rect(290, 490, 250, 100)

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

    c.rect(40, table_y, 500, 45)

    c.line(130, table_y, 130, table_y + 45)

    c.line(200, table_y, 200, table_y + 45)

    c.line(290, table_y, 290, table_y + 45)

    c.line(450, table_y, 450, table_y + 45)

    # HEADERS

    c.setFont("Helvetica-Bold", 7)

    c.drawString(
        50,
        475,
        "PERIOD OF SERVICES"
    )

    c.drawString(
        145,
        475,
        "P.O. NUMBER"
    )

    c.drawString(
        220,
        475,
        "REQUISITIONER"
    )

    c.drawString(
        300,
        475,
        "SHIPPED VIA email to"
    )

    c.drawString(
        450,
        475,
        "TERMS"
    )

    # VALUES

    c.setFont("Helvetica", 6)

    c.drawString(
        50,
        455,
        f"{month.upper()} HOURS"
    )

    c.drawString(
        220,
        455,
        resource_name
    )

    email_text = vendor_email

    c.drawString(
        295,
        455,
        email_text
    )

    c.drawString(
        455,
        455,
        "Due on receipt"
    )

    # ==========================================
    # MAIN HOURS TABLE
    # ==========================================

    main_y = 315

    c.rect(40, main_y, 500, 110)

    c.line(40, main_y + 80, 540, main_y + 80)

    c.line(170, main_y, 170, main_y + 110)

    c.line(430, main_y, 430, main_y + 110)

    # TABLE HEADERS

    c.setFont("Helvetica-Bold", 9)

    c.drawString(
        80,
        main_y + 90,
        "QUANTITY"
    )

    c.drawString(
        250,
        main_y + 90,
        "Project Details"
    )

    c.drawString(
        455,
        main_y + 90,
        "TOTAL"
    )

    # ==========================================
    # WEEKLY HOURS
    # ==========================================

    c.setFont("Helvetica-Bold", 7)

    week_y = main_y + 65

    for i, hours in enumerate(weekly_hours):

        week_text = (
            f"WEEK ENDING 04/{(i+1)*5:02d} – {hours} HRS."
        )

        c.drawString(
            45,
            week_y,
            week_text
        )

        week_y -= 15

    # ==========================================
    # PROJECT DETAILS
    # ==========================================

    c.setFont("Helvetica-Bold", 10)

    c.drawString(
        215,
        main_y + 45,
        f"IT SERVICES FOR {resource_name.upper()}"
    )

    c.drawString(
        255,
        main_y + 30,
        f"{month.upper()} MONTH"
    )

    # ==========================================
    # TOTAL AMOUNT
    # ==========================================

    c.setFont("Helvetica-Bold", 12)

    c.drawString(
        445,
        main_y + 35,
        f"${invoice_amount:,.2f}"
    )

    # ==========================================
    # PAYMENT SECTION
    # ==========================================

    payment_y = 170

    c.rect(40, payment_y, 500, 120)

    c.line(
        210,
        payment_y,
        210,
        payment_y + 120
    )

    # LEFT SIDE

    c.setFont("Helvetica-Bold", 8)

    c.drawString(
        50,
        270,
        "Check Details:"
    )

    c.setFont("Helvetica", 8)

    c.drawString(
        50,
        250,
        COMPANY_NAME
    )

    c.drawString(
        50,
        238,
        "470 Olde Worthington Rd"
    )

    c.drawString(
        50,
        226,
        "Suite 200 Westerville OH 43082"
    )

    # RIGHT SIDE

    c.setFont("Helvetica-Bold", 8)

    c.drawString(
        230,
        270,
        "Wire"
    )

    c.setFont("Helvetica", 8)

    c.drawString(
        230,
        250,
        "Payee Name: SLN Solution LLC"
    )

    c.drawString(
        230,
        238,
        "Account Number: 374006933920"
    )

    c.drawString(
        230,
        226,
        "Routing Number: 071214579"
    )

    c.drawString(
        230,
        214,
        "Wire: 026009593"
    )

    # TOTALS

    c.setFont("Helvetica-Bold", 10)

    c.drawString(
        430,
        250,
        "SUBTOTAL"
    )

    c.drawString(
        490,
        250,
        f"${invoice_amount:,.2f}"
    )

    c.drawString(
        430,
        230,
        "SALES TAX"
    )

    c.drawString(
        490,
        230,
        "$0.00"
    )

    c.drawString(
        430,
        190,
        "TOTAL due"
    )

    c.drawString(
        490,
        190,
        f"${invoice_amount:,.2f}"
    )

    # ==========================================
    # FOOTER
    # ==========================================

    c.line(40, 160, 540, 160)

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

    c.setFont("Helvetica-Bold", 16)

    c.setFillColor(colors.darkblue)

    c.drawCentredString(
        300,
        60,
        "Thank you for your business!"
    )

    # ==========================================
    # SAVE PDF
    # ==========================================

    c.save()

    # ==========================================
    # SEND EMAIL
    # ==========================================

    if vendor_email == "" or vendor_email.lower() == "nan":

        print("Skipped Empty Vendor Email")

        continue

    try:

        msg = EmailMessage()
        msg['Subject'] = (
            f"SLN Solution LLC - Invoice {invoice_number} for {month}"
        )
        

        msg['From'] = EMAIL_ADDRESS

        msg['To'] = vendor_email

        msg.set_content(f"""
                        
Dear Team,

Please find attached the invoice for services rendered during {month}.

Invoice Number: {invoice_number}

If you have any questions regarding this invoice, please feel free to contact us.

Thank you for your continued partnership.

Regards,

SLN Solution LLC
470 Olde Worthington Rd Suite 200
Westerville OH 43082

Email: srikanth@sln-solutions.com
Phone: +1 804-304-6616
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

        # FAILURE LOG

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

print("\nAll invoices processed successfully!")