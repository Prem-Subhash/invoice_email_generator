import os
import pandas as pd
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

COMPANY_NAME = "SLN Solution LLC"
CONTACT_EMAIL = "srikanth@sln-solutions.com"
CONTACT_PHONE = "+1 804-304-6616"
LOGO_PATH = "assets/sln_logo.png"

df = pd.read_excel("excel/Invoice_Data.xlsx")
os.makedirs("test_invoices", exist_ok=True)

invoice_counter = 1

def split_hours(total_hours):
    weekly_hours = []
    while total_hours > 40:
        weekly_hours.append(40)
        total_hours -= 40
    if total_hours > 0:
        weekly_hours.append(total_hours)
    return weekly_hours

for index, row in df.iterrows():
    resource_name = str(row[' Resource Name']).strip()
    vendor_name = str(row['Vendor Name']).strip()
    vendor_email = str(row['Vendor Email']).strip()
    month = str(row['Month']).strip()
    total_hours = int(row['No of Hours '])
    bill_rate = float(row['Bill rate'])
    invoice_amount = round(total_hours * bill_rate, 2)
    vendor_address = str(row['Vendor address']).strip()
    period_of_services = str(row['Period Of Services']).strip()
    if period_of_services.lower() == 'nan':
        period_of_services = ""
    invoice_number = f"SLN{invoice_counter:04d}"
    invoice_counter += 1
    weekly_hours = split_hours(total_hours)

    pdf_filename = f"test_{vendor_name}_{resource_name}_{month}_{index}.pdf"
    pdf_file = os.path.join("test_invoices", pdf_filename)

    # We set margins of 56pt on left and right, so printable width is exactly 500pt (612 - 56 - 56 = 500).
    # Top and bottom margins are set to 40pt.
    doc = SimpleDocTemplate(
        pdf_file,
        pagesize=letter,
        leftMargin=56,
        rightMargin=56,
        topMargin=40,
        bottomMargin=40
    )

    styles = getSampleStyleSheet()
    
    # Custom Paragraph Styles
    title_style = ParagraphStyle(
        'InvoiceTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=24,
        leading=28,
        alignment=1, # Center
        textColor=colors.HexColor('#000000'),
        spaceAfter=15
    )

    company_name_style = ParagraphStyle(
        'CompanyName',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=13,
        textColor=colors.HexColor('#000000')
    )

    company_address_style = ParagraphStyle(
        'CompanyAddress',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        leading=13,
        textColor=colors.HexColor('#333333')
    )

    meta_style = ParagraphStyle(
        'MetaInfo',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        leading=14,
        alignment=2, # Right
        textColor=colors.HexColor('#000000')
    )

    box_header_style = ParagraphStyle(
        'BoxHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=12,
        textColor=colors.HexColor('#000000')
    )

    box_name_style = ParagraphStyle(
        'BoxName',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=13,
        textColor=colors.HexColor('#000000'),
        spaceBefore=4,
        spaceAfter=2
    )

    box_body_style = ParagraphStyle(
        'BoxBody',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9.5,
        leading=13,
        textColor=colors.HexColor('#333333')
    )

    table_header_style = ParagraphStyle(
        'TableHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8,
        leading=11,
        alignment=1, # Center
        textColor=colors.HexColor('#000000')
    )

    table_cell_style = ParagraphStyle(
        'TableCell',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=11,
        alignment=1, # Center
        textColor=colors.HexColor('#000000')
    )

    table_cell_bold_style = ParagraphStyle(
        'TableCellBold',
        parent=table_cell_style,
        fontName='Helvetica-Bold'
    )

    main_header_style = ParagraphStyle(
        'MainHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=12,
        alignment=1, # Center
        textColor=colors.HexColor('#000000')
    )

    qty_style = ParagraphStyle(
        'QtyText',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8,
        leading=12,
        textColor=colors.HexColor('#000000'),
        alignment=0 # Left
    )

    project_details_style = ParagraphStyle(
        'ProjectDetails',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=14,
        alignment=1, # Center
        textColor=colors.HexColor('#000000')
    )

    total_amount_style = ParagraphStyle(
        'TotalAmount',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=11,
        leading=14,
        alignment=1, # Center
        textColor=colors.HexColor('#000000')
    )

    payment_label_style = ParagraphStyle(
        'PaymentLabel',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8,
        leading=11,
        textColor=colors.HexColor('#000000'),
        alignment=0 # Left
    )

    payment_value_style = ParagraphStyle(
        'PaymentValue',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=11,
        textColor=colors.HexColor('#333333'),
        alignment=0 # Left
    )

    totals_label_style = ParagraphStyle(
        'TotalsLabel',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=12,
        textColor=colors.HexColor('#000000'),
        alignment=0 # Left
    )

    totals_label_bold_style = ParagraphStyle(
        'TotalsLabelBold',
        parent=totals_label_style,
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=13,
        textColor=colors.HexColor('#000000')
    )

    totals_value_style = ParagraphStyle(
        'TotalsValue',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=12,
        textColor=colors.HexColor('#333333'),
        alignment=2 # Right
    )

    totals_value_bold_style = ParagraphStyle(
        'TotalsValueBold',
        parent=totals_value_style,
        fontName='Helvetica-Bold',
        fontSize=11,
        leading=13,
        textColor=colors.HexColor('#000000')
    )

    footer_style = ParagraphStyle(
        'FooterText',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=11,
        textColor=colors.HexColor('#333333')
    )

    thank_you_style = ParagraphStyle(
        'ThankYou',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=16,
        leading=20,
        alignment=1, # Center
        textColor=colors.darkblue,
        spaceBefore=15
    )

    story = []

    # 1. Logo (Centered)
    logo = Image(LOGO_PATH, width=140, height=60)
    logo.hAlign = 'CENTER'
    story.append(logo)
    story.append(Spacer(1, 10))

    # 2. Title "INVOICE"
    story.append(Paragraph("INVOICE", title_style))

    # 3. Company Info & Metadata Table
    company_details_p = Paragraph(
        f"<b>{COMPANY_NAME}</b><br/>"
        f"470 Olde Worthington Rd Suite 200<br/>"
        f"Westerville OH 43082",
        company_address_style
    )
    meta_p = Paragraph(
        f"Invoice: #{invoice_number}<br/>"
        f"Date: {datetime.now().strftime('%B %d, %Y')}",
        meta_style
    )
    info_table = Table([[company_details_p, meta_p]], colWidths=[250, 250])
    info_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 15))

    # 4. To / Ship To Boxes
    vendor_address_formatted = vendor_address.replace('\n', '<br/>')
    to_p = Paragraph(
        f"To: {vendor_name}<br/>"
        f"<br/>"
        f"{vendor_name}<br/>"
        f"{vendor_address_formatted}",
        box_body_style
    )
    ship_to_p = Paragraph(
        f"Ship To: {vendor_name}<br/>"
        f"<br/>"
        f"{vendor_name}<br/>"
        f"{vendor_address_formatted}",
        box_body_style
    )
    
    # We style these to share a middle border by using a grid style on the Table
    to_ship_table = Table([[to_p, ship_to_p]], colWidths=[250, 250])
    to_ship_table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#888888')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
    ]))
    story.append(to_ship_table)
    story.append(Spacer(1, 15))

    # 5. Header Table (Meta Table)
    service_text = f"{total_hours} HRS @ ${bill_rate:,.2f}/HR"
    header_data = [
        [
            Paragraph("Period Of Services", table_header_style),
            Paragraph("P.O. NUMBER", table_header_style),
            Paragraph("REQUISITIONER", table_header_style),
            Paragraph("SHIPPED VIA email to", table_header_style),
            Paragraph("F.O.B. POINT", table_header_style),
            Paragraph("TERMS", table_header_style)
        ],
        [
            Paragraph(period_of_services, table_cell_bold_style),
            Paragraph("", table_cell_style),
            Paragraph(resource_name, table_cell_bold_style),
            Paragraph(f"<font color='#0a58ca'><u>{vendor_email}</u></font>", table_cell_style),
            Paragraph("-", table_cell_style),
            Paragraph("Due on receipt", table_cell_style)
        ]
    ]
    meta_table = Table(header_data, colWidths=[120, 50, 90, 120, 50, 70])
    meta_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f2f2f2')),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#888888')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 15))

    # 6. Main Items Table
    qty_content = f"{total_hours} HRS @ ${bill_rate:.2f}/HR"
    qty_p = Paragraph(qty_content, qty_style)

    project_details_text = f"IT SERVICES FOR {resource_name.upper()}<br/>{month.upper()} MONTH"
    project_p = Paragraph(project_details_text, project_details_style)

    total_p = Paragraph(f"${invoice_amount:,.2f}", total_amount_style)

    main_table_data = [
        [
            Paragraph("QUANTITY", main_header_style),
            Paragraph("Project Details", main_header_style),
            Paragraph("TOTAL", main_header_style)
        ],
        [qty_p, project_p, total_p]
    ]
    main_table = Table(main_table_data, colWidths=[130, 270, 100])
    main_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f2f2f2')),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#888888')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('LEFTPADDING', (0, 1), (0, 1), 12), # Indent the weekly hours for better layout
    ]))
    story.append(main_table)
    story.append(Spacer(1, 15))

    # 7. Payment Details & Subtotal Table
    check_p = Paragraph(
        f"<b>Check Details:</b><br/>"
        f"{COMPANY_NAME}<br/>"
        f"470 Olde Worthington Rd<br/>"
        f"Suite 200 Westerville OH 43082",
        payment_value_style
    )
    wire_p = Paragraph(
        f"<b>Wire</b><br/>"
        f"Payee Name: SLN Solution LLC<br/>"
        f"Account Number: 374006933920<br/>"
        f"Routing Number: 071214579<br/>"
        f"Wire: 026009593",
        payment_value_style
    )
    
    payment_data = [
        [check_p, wire_p, Paragraph("SUBTOTAL", totals_label_style), Paragraph(f"${invoice_amount:,.2f}", totals_value_style)],
        ["", "", Paragraph("SALES TAX", totals_label_style), Paragraph("$0.00", totals_value_style)],
        ["", "", Paragraph("SHIPPING & HANDLING", totals_label_style), Paragraph("-", totals_value_style)],
        ["", "", Paragraph("TOTAL due", totals_label_bold_style), Paragraph(f"${invoice_amount:,.2f}", totals_value_bold_style)],
    ]
    
    payment_table = Table(payment_data, colWidths=[170, 160, 80, 90])
    payment_table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#888888')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('SPAN', (0, 0), (0, 3)),
        ('SPAN', (1, 0), (1, 3)),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('LINEABOVE', (2, 3), (3, 3), 1, colors.HexColor('#888888')),
    ]))
    story.append(payment_table)
    story.append(Spacer(1, 15))

    # 8. Footer
    divider = Table([[""]], colWidths=[500], rowHeights=[1])
    divider.setStyle(TableStyle([
        ('LINEABOVE', (0, 0), (-1, -1), 1, colors.HexColor('#888888')),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
    ]))
    story.append(divider)
    story.append(Spacer(1, 10))

    footer_p = Paragraph(
        f"If you have any questions concerning this invoice,<br/>"
        f"please contact {CONTACT_EMAIL}<br/>"
        f"Phone: {CONTACT_PHONE}",
        footer_style
    )
    story.append(footer_p)

    # 9. Thank You
    story.append(Paragraph("Thank you for your business!", thank_you_style))

    doc.build(story)
    print(f"Generated: {pdf_file}")
