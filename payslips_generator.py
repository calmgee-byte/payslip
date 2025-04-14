
import pandas as pd
import json
import openpyxl
from fpdf import FPDF
import subprocess
import sys
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib import colors
from reportlab.lib.units import inch
from datetime import datetime
import os
from pathlib import Path
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from num2words import num2words

def create_payslip_directory():
    """Creates the payslips directory if it doesn't exist."""
    Path('payslips').mkdir(exist_ok=True)

def create_payslip_pdf(employee_data, pdf_path):
    """Creates a PDF payslip for a single employee using a layout similar to the provided sample."""
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=letter,
        leftMargin=0.75 * inch,
        rightMargin=0.75 * inch,
        topMargin=0.75 * inch,
        bottomMargin=0.5 * inch
    )
    elements = []
    
    styles = {
        'centered_bold': ParagraphStyle(name='centered_bold', alignment=1, fontSize=14, fontName='Helvetica-Bold'),
        'normal': ParagraphStyle(name='normal', alignment=0, fontSize=10, fontName='Helvetica'),
        'label': ParagraphStyle(name='label', alignment=0, fontSize=10, fontName='Helvetica-Bold'),
        'centered': ParagraphStyle(name='centered', alignment=1, fontSize=10, fontName='Helvetica'),
        'centered_bold_large': ParagraphStyle(name='centered_bold_large', alignment=1, fontSize=12, fontName='Helvetica-Bold')
    }
    
    # Header with logo
    logo_path = "path/to/your/logo.png"  # Replace this path with your actual logo file path
    if os.path.exists(logo_path):
        logo = Image(logo_path, width=1.5 * inch, height=1 * inch)
        elements.append(logo)
        elements.append(Spacer(1, 0.2 * inch))
    
    elements.append(Paragraph("Payslip", styles['centered_bold']))
    elements.append(Spacer(1, 0.05 * inch))
    elements.append(Paragraph("CALMTECH HOLDINGS", styles['centered_bold_large']))
    elements.append(Paragraph("21023 Pearson Point Road<br/>Gateway Avenue", styles['centered']))
    elements.append(Spacer(1, 0.2 * inch))
    
    # Employee info
    employee_info = [
        [Paragraph("<b>Date of Joining</b>: 2018-06-23", styles['normal']),
         Paragraph(f"<b>Employee Name</b>: {employee_data['Name']}", styles['normal'])],
        [Paragraph(f"<b>Pay Period</b>: {datetime.now().strftime('%B %Y')}", styles['normal']),
         Paragraph(f"<b>Designation</b>: Marketing Executive", styles['normal'])],
        [Paragraph(f"<b>Worked Days</b>: 26", styles['normal']),
         Paragraph(f"<b>Department</b>: Marketing", styles['normal'])]
    ]
    elements.append(Table(employee_info, colWidths=[3 * inch, 3 * inch], hAlign='LEFT'))
    elements.append(Spacer(1, 0.2 * inch))
    
    # Salary breakdown
    earnings = [
        ["Basic", f"{employee_data['Basic Salary']:,.0f}"],
        ["Incentive Pay", "1000"],
        ["House Rent Allowance", "400"],
        ["Meal Allowance", "200"],
        ["Total Earnings", f"{employee_data['Basic Salary'] + 1000 + 400 + 200:,.0f}"]
    ]
    deductions = [
        ["Net Pay", f"{employee_data['Net Salary']:,.0f}"]
    ]
    table_data = [
        [Paragraph("<b>Earnings</b>", styles['label']), "", Paragraph("<b>Deductions</b>", styles['label']), ""]
    ] + [[e[0], e[1], d[0], d[1]] for e, d in zip(earnings, deductions)]
    
    table = Table(table_data, colWidths=[2.2 * inch, 1.2 * inch, 2.2 * inch, 1.2 * inch])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (1, 0), colors.lightgrey),
        ('BACKGROUND', (2, 0), (3, 0), colors.lightgrey),
        ('GRID', (0, 0), (-1, -1), 0.8, colors.black),
        ('ALIGN', (1, 1), (1, -1), 'RIGHT'),
        ('ALIGN', (3, 1), (3, -1), 'RIGHT'),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 0.3 * inch))
    
    # Net pay in numbers and words
    net_pay_words = num2words(employee_data['Net Salary'], to='cardinal').replace(',', '').title()
    elements.append(Paragraph(f"{int(employee_data['Net Salary'])}", styles['centered']))
    elements.append(Paragraph(f"{net_pay_words}", styles['centered']))
    elements.append(Spacer(1, 0.3 * inch))
    
    # Signatures
    signature_table = Table([
        [Paragraph("Employer Signature", styles['centered']),
         Paragraph("Employee Signature", styles['centered'])],
        ["________________________", "________________________"]
    ], colWidths=[3 * inch, 3 * inch])
    elements.append(signature_table)
    elements.append(Spacer(1, 0.2 * inch))
    
    # Footer note
    elements.append(Paragraph("This is system generated payslip", styles['centered']))
    
    # Build document
    doc.build(elements)
