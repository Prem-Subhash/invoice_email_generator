INVOICE AUTOMATION PROJECT

========================================

PROJECT DESCRIPTION
----------------------------------------

This project automates invoice generation and email sending using Python.

Features:
- Reads invoice data from Excel
- Generates PDF invoices automatically
- Sends invoices through company webmail SMTP
- Supports multiple vendors
- Creates audit logs
- Stores invoices in organized year/month folders

========================================

PROJECT STRUCTURE
----------------------------------------

invoice_automation/
│
├── excel/
│      └── Invoice_Data.xlsx
│
├── invoices/
│      └── 2026/
│            └── April/
│
├── logs.txt
├── main.py
├── run_invoice.bat
├── requirements.txt
├── README.txt
└── .env

========================================

REQUIRED PACKAGES
----------------------------------------

Install required packages:

pip install -r requirements.txt

========================================

ENVIRONMENT VARIABLES
----------------------------------------

Create .env file with:

SMTP_SERVER=mail.slnsolution.com
SMTP_PORT=465
EMAIL_ADDRESS=your_email@slnsolution.com
EMAIL_PASSWORD=your_password

========================================

HOW TO RUN
----------------------------------------

Method 1:
Run manually:

py main.py

Method 2:
Run batch file:

run_invoice.bat

========================================

INPUT FILE FORMAT
----------------------------------------

Excel File:
excel/Invoice_Data.xlsx

Required Columns:
- Company Name
- Resource Name
- Vendor Name
- Company Address
- Vendor address
- Vendor Email
- Company Email
- Month
- No of Hours
- Bill rate
- Invoice Amount

========================================

OUTPUT
----------------------------------------

Generated PDFs:
invoices/year/month/

Logs:
logs.txt

========================================

NOTES
----------------------------------------

- Keep Excel file closed while running script
- Do not share real SMTP passwords
- Update .env before production use