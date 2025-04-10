# import pandas as pd
# import json
# import openpyxl
# from fpdf import FPDF

# # Read the Excel file
# df = pd.read_excel("payslip project.xlsx")

# # Display current columns to verify names
# print("Current columns in the  Dataframe:")
# print(df.columns)

# # Select the required columns (using exact column names from your data)
# df = df[['NAME', 'Employee ID', 'Email Address', 'Basic Salary', 'Allowances', 'Deductions']]

# # Net salary calculation (using exact column names from your data)
# df['Net Salary'] = df['Basic Salary'] + df['Allowances'] - df['Deductions']

# # Display results with formatted currency
# pd.options.display.float_format = '${:,.2f}'.format # Fixed the typo here
# print("\nEmployee Salary Detail:")
# print(df[['Employee ID', 'Name', 'Basic Salary', 'Allowances', 'Deductions', 'Net Salary']])

# from reportlab.lib.pagesizes import letter
# from reportlab.pdfgen import canvas
# from reportlab.lib.styles import ParagraphStyle
# from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
# from reportlab.lib import colors
# from reportlab.lib.units import inch
# from datetime import datetime
# import os
# from pathlib import Path
# import pandas as pd

# def create_payslip_directory():
#     """Create payslips directly if it doesn't exist."""
#     Path('payslips').mkdir(exist_ok=True)

# def create_payslip_pdf(employee_data, pdf_path):
#     """Create a professional payslip PDF for an employee."""
#     # Create PDF document
#     doc = SimpleDocTemplate(
#         pdf_path,
#         pagesize=letter,
#         leftMargin=0.75 * inch,
#         rightMargin=1 * inch,
#         topMargin=1 * inch,
#         bottomMargin=0.5 * inch
#     )

#     # Initialize elements list
#     elements = []

#     # Create styles
#     styles = {
#       'heading': ParagraphStyle(
#           name='heading',
#           fontSize=16,
#           leading=20,
#           alignment=1,
#           textColor=colors.black,
#           fontName='Helvetica-Bold'
#         ),
#         'subheading': ParagraphStyle(
#             name='subheading',
#             fontSize=12,
#             leading=14,
#             alignment=0,
#             textColor=colors.black,
#             fontName='Helvetica-Bold'
#         ),
#         'body': ParagraphStyle(
#             name='body',
#             fontSize=10,
#             leading=12,
#             alignment=0,
#             textColor=colors.black,
#             fontName='Helvetica'
#         )
#     }

#     # Add header section
#     elements.append(Paragraph("Monthly Payslip", styles['heading']))
#     elements.append(Spacer(1, 0.2 * inch))

#     # Add employee details
#     elements.append(Paragraph(f"Employee Name: {employee_data['Name']}", styles['body']))
#     elements.append(Paragraph(f"Employee ID: {employee_data['Employee ID']}", styles['body']))
#     elements.append(Paragraph(f"Date: {datetime.now().strftime('%B %Y')}", styles['body']))
#     elements.append(Spacer(1, 0,3 * inch))

#     # Add salary details table
#     salary_data = [
#         ['Basic Salary:', f"${employee_data['Basic Salary']:,.2f}"],
#         ['Allowances:', f"${employee_data['Allowances']:,.2f}"],
#         ['Deductions:', f"${employee_data['Deductions']:,.2f}"],
#         ['Net Salary:', f"${employee_data['Net Salary']:,.2f}"]
#     ]

#     salary_table = Table(salary_data, style=[
#         ('GRID', (0,0), (-1,-1), 1, colors.black),
#         ('FONTNAME', (0,0), (-1,-1), 1, 'Helvetica'),
#         ('FONTSIZE', (0,0), (-1,-1), 10),
#         ('BACKGROUND', (0,0), (-1,0), colors.grey),
#         ('TEXTCOLOR', (0,0), (-1,0), colors.black),
#         ('ALIGN', (0,0), (-1,-1), 'LEFT'),
#         ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
#     ])

#     elements.append(salary_table)
#     elements.append(Spacer(Spacer(1, 0.3 * inch)))

#     #Add footer section
#     elements.append(Paragraph(
#         "This payslip was generated automically. Please contact the payroll department "
#         "if you notice any discrepancies.",
#         styles['body']
#     ))
    
#     # Build PDF document
#     doc.build(elements)

# def generate_payslips(df):
#     """Generate payslips for all employees in the Dataframe."""
#     create_payslip_directory()   
  
#     for _,row in df.iterrows():
#         try:
#             pdf_path = f"payslips/{row['Employee ID']}.pdf"
#             create_payslip_pdf(row, pdf_path)
#             print(f"Payslip generated for {row['Name']} (ID: {row['Employee ID']})")
#         except Exception as e:
#             print(f"Error generating payslip for {row['Name']}: {str(e)}")    

# # Example usage with your data
# if __name__=="__main__":
#     # Your employee data
#     data ={
#         'Employee ID' : ['TC10', 'TCO9', 'TC08', 'TC07', 'TCO6', 'TC05', 'TC04', 'TCO3', 'TCO2', 'TCO1', 'TC11', 'TC12'],
#         'Name': ['Taps Ciriboto', 'Mbali Mushayi', 'Demy Bingura', 'Arty Sibanda', 'Emphraim Buruvuru', 'Kudzai Tivatye', 'Reece Kurangwa', 'Lloyd Chogari', 'Thobela Sithole', 'Shaine Kaduhwa', 'Nashe Graphix', 'Rue Gurure'],
#         'Basic Salary': [2100.00, 2100.00, 1800.00, 1750.00, 1500.00, 2000.00, 2000.00, 2000.00, 2000.00, 2500.00, 1500.00, 1500.00],
#         'Allowances': [500.00, 500.00, 450.00, 450.00, 450.00, 450.00, 450.00, 450.00, 450.00, 500.00, 450.00, 450.00],
#         'Deductions': [95.00, 95.00, 45.00, 30.00, 40.00, 40.00, 50.00, 45.00, 45.00, 25.00, 15.00, 15.00],
#         'Net Salary' :
#     }

# import smtplib
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText
# from email.mime.application import MIMEApplication
# import os
# from pathlib import Path
# import pandas as pd
# from datetime import datetime

# def send_payslip_email(config, employee_data, pdf_path):
#     """Send payslip as email attachment to employee."""
#     try:
#         #Create message
#         msg = MIMEMultipart()
#         msg['Subject'] = "Your Payslip for This Month"
#         msg['From'] = config['FROM_EMAIL']
#         msg['TO'] = employee_data['EMAIL']

#         # Email body
#         body = f"""
# Dear {employee_data['NAME']},

# Please access your monthly payslip attached to this email.

# Enjoy your day,
# Accounts Department
# """
        
#         # Attach message body
#         msg.attach(MIMEText(Body, 'plain'))

#         # Attach PDF
#         with open(pdf_path, 'rb') as f:
#             attachment = MIMEApplication(f.read(), _subtype='pdf')
#             attachment.add_header('Content-Disposition', 'attachment', filename=Path(pdf_path).name)
#             msg.attach(attachment)

#         # Send email using TLS
#         with smtplib.SMTP(config['SMTP_SERVER'], config['SMTP_PORT']) as server:
#             server.starttls()
#             server.login(config['FROM_EMAIL'], config['EMAIL_PASSWORD'])
#             server.send_message(msg)

#         print(f"Email sent successfully to {employee_data['EMAIL']}")
#         return True

#     except Exception as e:
#         print(f"Error sending email to {employee_data['EMAIL']}: {str(e)}")
#         return False

# def send_all_payslips(df, config):
#     """Send payslips to all employees in the Dataframe."""
#     success_count = 0
#     total_employees = len(df)

#     for _, row in df.iterrows():
#         pdf_path = f"payslips/{row['EMPLOYEE_ID']}.pdf"

#         if os.path.exists(pdf_path):
#             if send_payslip_email(config, row, pdf_path):
#                 success_count += 1
#             else:
#                 print(f"PDF not found for {row['Name']} (ID: {row['Employee_ID']})")

#     print(f"\nSummary:")
#     print(f"Total employees: {total_employees}")
#     print(f"Emails sent successfully: {success_count}")
#     print(f"Failed emails: {total_employees - success_count }")

# def load_config():
#     """Load configuration from environment variables with default values."""
#     config = {
#         'SMTP_SERVER': os.getenv('SMTP_SERVER', 'smtp.gmail.com'),
#         'SMTP_PORT': int(os.getenv('SMTP_PORT', '587')),
#         'FROM_EMAIL': os.getenv('FROM_EMAIL', "spliffking16@gmail.com"),
#         'EMAIL_PASSWORD': os.getenv('EMAIL_PASSWORD', "bumu suqb dwcb zklp")
#     }

#     # Verify required configuration
#     required_keys = ['SMTP_SERVER', 'SMTP_PORT', 'FROM_EMAIL', 'EMAIL_PASSWORD']
#     missing_keys = [key for key in required_keys if not config[key]]

#     if missing_keys:
#         print(f"Missing required  configuration: {', '.join(missing_keys)}")
#         print("\nPlease set these environment variables or provide default values in the code.")
#         return None
#     return config

# if __name__ == "__main__":
#     # Load configuration
#     config = load_config()
#     if config is None:
#         exit(1)                         
    
#     # Your employee data
#     data = {
#         'Employee ID': ['TC10', 'TC09', 'TCO8', 'TC07', 'TC06', 'TC05', 'TC04', 'TCO3', 'TCO2', 'TC01', 'TC11', 'TC12'],
#         'Name': ['Taps Ciriboto', 'Mbali Mushayi', 'Demy Bingura', 'Arty Sibanda', 'Emphraim Buruvuru', 'Kudzai Tivatye', 'Reece Kurangwa', 'Lloyd Chogari', 'Thobela Sithole', 'Shaine Kaduhwa', 'Nashe Graphix', 'Rue Gurure'],
#         'Email Address': ['ciribotonicole@gmail.com', 'tafadzwamushayi3@gmail.com', 'demycadwell@gmail.com', 'sibandaaartii15@gmail.com', 'eoburuvuru@gmail.com', 'nicolaskudzai696@gmail.com', 'Kurangwareece@gmail.com', 'lloyddonnel44@gmail.com', 'thobelacarltonsithole@gmail.com', 'kaduhwashaine20@gmail.com', 'nashegraphix@gmail.com', 'ruvimbo448@gmail.com'],                 ]
#         'Basic Salary': [2100.00, 2100.00, 1800.00, 1750.00, 1500.00, 2000.00, 2000.00, 2000.00, 2000.00, 2500.00, 1500.00, 1500.00],
#         'Allowances': [500.00, 500.00, 450.00, 450.00, 450.00, 450.00, 450.00, 450.00, 450.00, 500.00, 450.00, 450.00],
#         'Deductions': [95.00, 95.00, 45.00, 30.00, 40.00, 40.00, 50.00, 45.00, 45.00, 25.00, 15.00, 15.00],
#         'Net Salary':
#     }

#     df = pd.DataFrame(data)
#     send_all_payslips(df, config)


# import pandas as pd
# import json
# import openpyxl
# import subprocess
# import sys

# # Automatically install num2words if not present
# try:
#     from num2words import num2words
# except ImportError:
#     subprocess.check_call([sys.executable, "-m", "pip", "install", "num2words"])
#     from num2words import num2words

# from reportlab.lib.pagesizes import letter
# from reportlab.pdfgen import canvas
# from reportlab.lib.styles import ParagraphStyle
# from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
# from reportlab.lib import colors
# from reportlab.lib.units import inch
# from datetime import datetime
# import os
# from pathlib import Path
# import smtplib
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText
# from email.mime.application import MIMEApplication

# # Create payslips directory
# def create_payslip_directory():
#     Path('payslips').mkdir(exist_ok=True)

# # Create PDF payslip
# def create_payslip_pdf(employee_data, pdf_path):
#     doc = SimpleDocTemplate(
#         pdf_path,
#         pagesize=letter,
#         leftMargin=0.75 * inch,
#         rightMargin=1 * inch,
#         topMargin=1 * inch,
#         bottomMargin=0.5 * inch
#     )

#     elements = []

#     styles = {
#         'heading': ParagraphStyle(name='heading', fontSize=16, leading=20, alignment=1, textColor=colors.black, fontName='Helvetica-Bold'),
#         'subheading': ParagraphStyle(name='subheading', fontSize=12, leading=14, alignment=0, textColor=colors.black, fontName='Helvetica-Bold'),
#         'body': ParagraphStyle(name='body', fontSize=10, leading=12, alignment=0, textColor=colors.black, fontName='Helvetica')
#     }

#     elements.append(Paragraph("Monthly Payslip", styles['heading']))
#     elements.append(Spacer(1, 0.2 * inch))

#     elements.append(Paragraph(f"Employee Name: {employee_data['Name']}", styles['body']))
#     elements.append(Paragraph(f"Employee ID: {employee_data['Employee ID']}", styles['body']))
#     elements.append(Paragraph(f"Date: {datetime.now().strftime('%B %Y')}", styles['body']))
#     elements.append(Spacer(1, 0.3 * inch))

#     salary_data = [
#         ['Basic Salary:', f"${employee_data['Basic Salary']:,.2f}"],
#         ['Allowances:', f"${employee_data['Allowances']:,.2f}"],
#         ['Deductions:', f"${employee_data['Deductions']:,.2f}"],
#         ['Net Salary:', f"${employee_data['Net Salary']:,.2f}"]
#     ]

#     salary_table = Table(salary_data, style=[
#         ('GRID', (0, 0), (-1, -1), 1, colors.black),
#         ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
#         ('FONTSIZE', (0, 0), (-1, -1), 10),
#         ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
#         ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
#         ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
#         ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')
#     ])

#     elements.append(salary_table)
#     elements.append(Spacer(1, 0.3 * inch))

#     # ✅ Convert net salary to words and display both centered
#     net_pay = employee_data['Net Salary']
#     net_pay_words = num2words(net_pay, to='cardinal').replace(',', '').title()

#     elements.append(Spacer(1, 0.1 * inch))
#     elements.append(Paragraph(f"<para align='center'><b>Net Salary:</b> ${net_pay:,.2f}</para>", styles['body']))
#     elements.append(Paragraph(f"<para align='center'><b>In Words:</b> {net_pay_words} Dollars</para>", styles['body']))
#     elements.append(Spacer(1, 0.3 * inch))

#     elements.append(Paragraph(
#         "This payslip was generated automatically. Please contact the payroll department "
#         "if you notice any discrepancies.",
#         styles['body']
#     ))

#     doc.build(elements)

# # Generate payslips for all employees
# def generate_payslips(df):
#     create_payslip_directory()
#     for _, row in df.iterrows():
#         try:
#             pdf_path = f"payslips/{row['Employee ID']}.pdf"
#             create_payslip_pdf(row, pdf_path)
#             print(f"Payslip generated for {row['Name']} (ID: {row['Employee ID']})")
#         except Exception as e:
#             print(f"Error generating payslip for {row['Name']}: {str(e)}")

# # Send payslip via email
# def send_payslip_email(config, employee_data, pdf_path):
#     try:
#         msg = MIMEMultipart()
#         msg['Subject'] = "Your Payslip for This Month"
#         msg['From'] = config['FROM_EMAIL']
#         msg['To'] = employee_data['Email Address']

#         body = f"""
# Dear {employee_data['Name']},

# Please access your monthly payslip attached to this email.

# Enjoy your day,
# Accounts Department
# """
#         msg.attach(MIMEText(body, 'plain'))

#         with open(pdf_path, 'rb') as f:
#             attachment = MIMEApplication(f.read(), _subtype='pdf')
#             attachment.add_header('Content-Disposition', 'attachment', filename=Path(pdf_path).name)
#             msg.attach(attachment)

#         with smtplib.SMTP(config['SMTP_SERVER'], config['SMTP_PORT']) as server:
#             server.starttls()
#             server.login(config['FROM_EMAIL'], config['EMAIL_PASSWORD'])
#             server.send_message(msg)

#         print(f"Email sent successfully to {employee_data['Email Address']}")
#         return True
#     except Exception as e:
#         print(f"Error sending email to {employee_data['Email Address']}: {str(e)}")
#         return False

# # Send all payslips
# def send_all_payslips(df, config):
#     success_count = 0
#     total_employees = len(df)

#     for _, row in df.iterrows():
#         pdf_path = f"payslips/{row['Employee ID']}.pdf"
#         if os.path.exists(pdf_path):
#             if send_payslip_email(config, row, pdf_path):
#                 success_count += 1
#         else:
#             print(f"PDF not found for {row['Name']} (ID: {row['Employee ID']})")

#     print(f"\nSummary:")
#     print(f"Total employees: {total_employees}")
#     print(f"Emails sent successfully: {success_count}")
#     print(f"Failed emails: {total_employees - success_count}")

# # Load email config
# def load_config():
#     config = {
#         'SMTP_SERVER': os.getenv('SMTP_SERVER', 'smtp.gmail.com'),
#         'SMTP_PORT': int(os.getenv('SMTP_PORT', '587')),
#         'FROM_EMAIL': os.getenv('FROM_EMAIL', ""),
#         'EMAIL_PASSWORD': os.getenv('EMAIL_PASSWORD', "")
#     }

#     missing_keys = [key for key, value in config.items() if not value]
#     if missing_keys:
#         print(f"Missing required configuration: {', '.join(missing_keys)}")
#         return None
#     return config

# # Entry point
# if __name__ == "__main__":
#     config = load_config()
#     if config is None:
#         sys.exit(1)

#     # Example employee data (you can switch this with reading from Excel if needed)
#     data = {
#         'Employee ID': ['TC10', 'TC09', 'TC08', 'TC07', 'TC06', 'TC05', 'TC04', 'TC03', 'TC02', 'TC01', 'TC11', 'TC12'],
#         'Name': ['Taps Ciriboto', 'Mbali Mushayi', 'Demy Bingura', 'Arty Sibanda', 'Emphraim Buruvuru', 'Kudzai Tivatye', 'Reece Kurangwa', 'Lloyd Chogari', 'Thobela Sithole', 'Shaine Kaduhwa', 'Nashe Graphix', 'Rue Gurure'],
#         'Email Address': ['ciribotonicole@gmail.com', 'tafadzwamushayi3@gmail.com', 'demycadwell@gmail.com', 'sibandaaartii15@gmail.com', 'eoburuvuru@gmail.com', 'nicolaskudzai696@gmail.com', 'Kurangwareece@gmail.com', 'lloyddonnel44@gmail.com', 'thobelacarltonsithole@gmail.com', 'kaduhwashaine20@gmail.com', 'nashegraphix@gmail.com', 'ruvimbo448@gmail.com'],
#         'Basic Salary': [2100.00, 2100.00, 1800.00, 1750.00, 1500.00, 2000.00, 2000.00, 2000.00, 2000.00, 2500.00, 1500.00, 1500.00],
#         'Allowances': [500.00, 500.00, 450.00, 450.00, 450.00, 450.00, 450.00, 450.00, 450.00, 500.00, 450.00, 450.00],
#         'Deductions': [95.00, 95.00, 45.00, 30.00, 40.00, 40.00, 50.00, 45.00, 45.00, 25.00, 15.00, 15.00]
#     }

#     df = pd.DataFrame(data)
#     df['Net Salary'] = df['Basic Salary'] + df['Allowances'] - df['Deductions']

#     generate_payslips(df)
#     send_all_payslips(df, config)

# import pandas as pd
# import json
# import openpyxl
# from fpdf import FPDF
# import subprocess
# import sys
# from reportlab.lib.pagesizes import letter
# from reportlab.pdfgen import canvas
# from reportlab.lib.styles import ParagraphStyle
# from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
# from reportlab.lib import colors
# from reportlab.lib.units import inch
# from datetime import datetime
# import os
# from pathlib import Path
# import smtplib
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText
# from email.mime.application import MIMEApplication
# from num2words import num2words

# # Automatically install num2words if not present
# try:
#     from num2words import num2words
# except ImportError:
#     subprocess.check_call([sys.executable, "-m", "pip", "install", "num2words"])
#     from num2words import num2words

# # Create payslips directory
# def create_payslip_directory():
#     Path('payslips').mkdir(exist_ok=True)

# # Create PDF payslip
# def create_payslip_pdf(employee_data, pdf_path):
#     doc = SimpleDocTemplate(
#         pdf_path,
#         pagesize=letter,
#         leftMargin=0.75 * inch,
#         rightMargin=1 * inch,
#         topMargin=1 * inch,
#         bottomMargin=0.5 * inch
#     )

#     elements = []

#     styles = {
#         'heading': ParagraphStyle(name='heading', fontSize=16, leading=20, alignment=1, textColor=colors.black, fontName='Helvetica-Bold'),
#         'subheading': ParagraphStyle(name='subheading', fontSize=12, leading=14, alignment=0, textColor=colors.black, fontName='Helvetica-Bold'),
#         'body': ParagraphStyle(name='body', fontSize=10, leading=12, alignment=0, textColor=colors.black, fontName='Helvetica')
#     }

#     # Company name
#     elements.append(Paragraph("CALMTECH HOLDING", styles['heading']))
#     elements.append(Spacer(1, 0.1 * inch))

#     # Payslip title
#     elements.append(Paragraph("Monthly Payslip", styles['heading']))
#     elements.append(Spacer(1, 0.2 * inch))

#     elements.append(Paragraph(f"Employee Name: {employee_data['Name']}", styles['body']))
#     elements.append(Paragraph(f"Employee ID: {employee_data['Employee ID']}", styles['body']))
#     elements.append(Paragraph(f"Date: {datetime.now().strftime('%B %Y')}", styles['body']))
#     elements.append(Spacer(1, 0.3 * inch))

#     salary_data = [
#         ['Basic Salary:', f"${employee_data['Basic Salary']:,.2f}"],
#         ['Allowances:', f"${employee_data['Allowances']:,.2f}"],
#         ['Deductions:', f"${employee_data['Deductions']:,.2f}"],
#         ['Net Salary:', f"${employee_data['Net Salary']:,.2f}"]
#     ]

#     salary_table = Table(salary_data, style=[ 
#         ('GRID', (0, 0), (-1, -1), 1, colors.black),
#         ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
#         ('FONTSIZE', (0, 0), (-1, -1), 10),
#         ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
#         ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
#         ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
#         ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')
#     ])

#     elements.append(salary_table)
#     elements.append(Spacer(1, 0.3 * inch))

#     # ✅ Convert net salary to words and display both centered
#     net_pay = employee_data['Net Salary']
#     net_pay_words = num2words(net_pay, to='cardinal').replace(',', '').title()

#     elements.append(Spacer(1, 0.1 * inch))
#     elements.append(Paragraph(f"<para align='center'><b>Net Salary:</b> ${net_pay:,.2f}</para>", styles['body']))
#     elements.append(Paragraph(f"<para align='center'><b>In Words:</b> {net_pay_words} Dollars</para>", styles['body']))
#     elements.append(Spacer(1, 0.3 * inch))

#     elements.append(Paragraph(
#         "This payslip was generated automatically. Please contact the payroll department "
#         "if you notice any discrepancies.",
#         styles['body']
#     ))

#     doc.build(elements)

# # Generate payslips for all employees
# def generate_payslips(df):
#     create_payslip_directory()
#     for _, row in df.iterrows():
#         try:
#             pdf_path = f"payslips/{row['Employee ID']}.pdf"
#             create_payslip_pdf(row, pdf_path)
#             print(f"Payslip generated for {row['Name']} (ID: {row['Employee ID']})")
#         except Exception as e:
#             print(f"Error generating payslip for {row['Name']}: {str(e)}")

# # Send payslip via email
# def send_payslip_email(config, employee_data, pdf_path):
#     try:
#         msg = MIMEMultipart()
#         msg['Subject'] = "Your Payslip for This Month"
#         msg['From'] = config['FROM_EMAIL']
#         msg['To'] = employee_data['Email Address']

#         body = f"""
# Dear {employee_data['Name']},

# Please access your monthly payslip attached to this email.

# Enjoy your day,
# Accounts Department
# """
#         msg.attach(MIMEText(body, 'plain'))

#         with open(pdf_path, 'rb') as f:
#             attachment = MIMEApplication(f.read(), _subtype='pdf')
#             attachment.add_header('Content-Disposition', 'attachment', filename=Path(pdf_path).name)
#             msg.attach(attachment)

#         with smtplib.SMTP(config['SMTP_SERVER'], config['SMTP_PORT']) as server:
#             server.starttls()
#             server.login(config['FROM_EMAIL'], config['EMAIL_PASSWORD'])
#             server.send_message(msg)

#         print(f"Email sent successfully to {employee_data['Email Address']}")
#         return True
#     except Exception as e:
#         print(f"Error sending email to {employee_data['Email Address']}: {str(e)}")
#         return False

# # Send all payslips
# def send_all_payslips(df, config):
#     success_count = 0
#     total_employees = len(df)

#     for _, row in df.iterrows():
#         pdf_path = f"payslips/{row['Employee ID']}.pdf"
#         if os.path.exists(pdf_path):
#             if send_payslip_email(config, row, pdf_path):
#                 success_count += 1
#         else:
#             print(f"PDF not found for {row['Name']} (ID: {row['Employee ID']})")

#     print(f"\nSummary:")
#     print(f"Total employees: {total_employees}")
#     print(f"Emails sent successfully: {success_count}")
#     print(f"Failed emails: {total_employees - success_count}")

# # Load email config (Direct configuration method for testing)
# def load_config():
#     config = {
#         'SMTP_SERVER': 'smtp.gmail.com',  # Use appropriate SMTP server for your email provider
#         'SMTP_PORT': 587,
#         'FROM_EMAIL': "youremail@example.com",  # Replace with your email
#         'EMAIL_PASSWORD': "yourpassword"       # Replace with your email password
#     }

#     missing_keys = [key for key, value in config.items() if not value]
#     if missing_keys:
#         print(f"Missing required configuration: {', '.join(missing_keys)}")
#         return None
#     return config

# # Entry point
# if __name__ == "__main__":
#     config = load_config()
#     if config is None:
#         sys.exit(1)

#     # Example employee data (you can switch this with reading from Excel if needed)
#     data = {
#         'Employee ID': ['TC10', 'TC09', 'TC08', 'TC07', 'TC06', 'TC05', 'TC04', 'TC03', 'TC02', 'TC01', 'TC11', 'TC12'],
#         'Name': ['Taps Ciriboto', 'Mbali Mushayi', 'Demy Bingura', 'Arty Sibanda', 'Emphraim Buruvuru', 'Kudzai Tivatye', 'Reece Kurangwa', 'Lloyd Chogari', 'Thobela Sithole', 'Shaine Kaduhwa', 'Nashe Graphix', 'Rue Gurure'],
#         'Email Address': ['ciribotonicole@gmail.com', 'tafadzwamushayi3@gmail.com', 'demycadwell@gmail.com', 'sibandaaartii15@gmail.com', 'eoburuvuru@gmail.com', 'nicolaskudzai696@gmail.com', 'Kurangwareece@gmail.com', 'lloyddonnel44@gmail.com', 'thobelacarltonsithole@gmail.com', 'kaduhwashaine20@gmail.com', 'nashegraphix@gmail.com', 'ruvimbo448@gmail.com'],
#         'Basic Salary': [2100.00, 2100.00, 1800.00, 1750.00, 1500.00, 2000.00, 2000.00, 2000.00, 2000.00, 2500.00, 1500.00, 1500.00],
#         'Allowances': [500.00, 500.00, 450.00, 450.00, 450.00, 450.00, 450.00, 450.00, 450.00, 500.00, 450.00, 450.00],
#         'Deductions': [95.00, 95.00, 45.00, 30.00, 40.00, 40.00, 50.00, 45.00, 45.00, 25.00, 15.00, 15.00]
#     }

#     df = pd.DataFrame(data)
#     df['Net Salary'] = df['Basic Salary'] + df['Allowances'] - df['Deductions']

#     generate_payslips(df)
#     send_all_payslips(df, config)


import pandas as pd
import json
import openpyxl
from fpdf import FPDF
import subprocess
import sys
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
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
    """Creates a PDF payslip for a single employee."""
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=letter,
        leftMargin=0.75 * inch,
        rightMargin=1 * inch,
        topMargin=1 * inch,
        bottomMargin=0.5 * inch
    )
    elements = []
    
    styles = {
        'heading': ParagraphStyle(name='heading', fontSize=16, leading=20,
                                alignment=1, textColor=colors.white,
                                fontName='Helvetica-Bold'),
        'salary_box_heading': ParagraphStyle(name='salary_box_heading',
                                           fontSize=14, leading=18,
                                           alignment=1, textColor=colors.black,
                                           fontName='Helvetica-Bold'),
        'salary_label': ParagraphStyle(name='salary_label',
                                     fontSize=11, leading=13,
                                     alignment=0, textColor=colors.black,
                                     fontName='Helvetica-Bold'),
        'salary_value': ParagraphStyle(name='salary_value',
                                     fontSize=11, leading=13,
                                     alignment=0, textColor=colors.black,
                                     fontName='Helvetica'),
        'body': ParagraphStyle(name='body', fontSize=10, leading=12,
                             alignment=0, textColor=colors.black,
                             fontName='Helvetica')
    }
    
    header_frame = Table([
        ['CALMTECH HOLDING'],
        ['Monthly Payslip']
    ], style=[
        ('BACKGROUND', (0, 0), (0, 1), colors.blue),
        ('TEXTCOLOR', (0, 0), (0, 1), colors.white),
        ('FONTNAME', (0, 0), (0, 1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (0, 1), [16, 12]),
        ('ALIGN', (0, 0), (0, 1), 'CENTER'),
        ('VALIGN', (0, 0), (0, 1), 'MIDDLE')
    ])
    elements.append(header_frame)
    
    elements.append(Spacer(1, 0.2 * inch))
    elements.append(Paragraph(f"Employee Name: {employee_data['Name']}", styles['body']))
    elements.append(Paragraph(f"Employee ID: {employee_data['Employee ID']}", styles['body']))
    elements.append(Paragraph(f"Date: {datetime.now().strftime('%B %Y')}", styles['body']))
    
    salary_box = Table([
        ['SALARY BREAKDOWN'],
        ['Basic Salary:', f"${employee_data['Basic Salary']:,.2f}"],
        ['Allowances:', f"${employee_data['Allowances']:,.2f}"],
        ['Deductions:', f"${employee_data['Deductions']:,.2f}"],
        ['Net Salary:', f"${employee_data['Net Salary']:,.2f}"]
    ], style=[
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('BOX', (0, 0), (-1, -1), 2, colors.black),
        ('GRID', (0, 1), (-1, -1), 1, colors.black),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), [14, 11]),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')
    ])
    elements.append(Spacer(1, 0.3 * inch))
    elements.append(salary_box)
    
    elements.append(Spacer(1, 0.3 * inch))
    net_pay = employee_data['Net Salary']
    net_pay_words = num2words(net_pay, to='cardinal').replace(',', '').title()
    elements.append(Paragraph(f"<para align='center'><b>Net Salary:</b> ${net_pay:,.2f}</para>", styles['body']))
    elements.append(Paragraph(f"<para align='center'><b>In Words:</b> {net_pay_words} Dollars</para>", styles['body']))
    
    elements.append(Spacer(1, 0.3 * inch))
    elements.append(Paragraph(
        "This payslip was generated automatically. Please contact the payroll department "
        "if you notice any discrepancies.",
        styles['body']
    ))
    
    doc.build(elements)

def generate_payslips(df):
    """Generates payslip PDFs for all employees in the DataFrame."""
    create_payslip_directory()
    for _, row in df.iterrows():
        try:
            pdf_path = f"payslips/{row['Employee ID']}.pdf"
            create_payslip_pdf(row, pdf_path)
            print(f"Payslip generated for {row['Name']} (ID: {row['Employee ID']})")
        except Exception as e:
            print(f"Error generating payslip for {row['Name']}: {str(e)}")

def send_payslip_email(config, employee_data, pdf_path):
    """Sends a payslip via email to an employee."""
    try:
        msg = MIMEMultipart()
        msg['Subject'] = "Your Payslip for This Month"
        msg['From'] = config['FROM_EMAIL']
        msg['To'] = employee_data['Email Address']
        body = f"""Dear {employee_data['Name']},
Please access your monthly payslip attached to this email.
Enjoy your day,
Accounts Department"""
        msg.attach(MIMEText(body, 'plain'))
        with open(pdf_path, 'rb') as f:
            attachment = MIMEApplication(f.read(), _subtype='pdf')
            attachment.add_header('Content-Disposition', 'attachment', filename=Path(pdf_path).name)
            msg.attach(attachment)
        with smtplib.SMTP(config['SMTP_SERVER'], config['SMTP_PORT']) as server:
            server.starttls()
            server.login(config['FROM_EMAIL'], config['EMAIL_PASSWORD'])
            server.send_message(msg)
        print(f"Email sent successfully to {employee_data['Email Address']}")
        return True
    except Exception as e:
        print(f"Error sending email to {employee_data['Email Address']}: {str(e)}")
        return False

def send_all_payslips(df, config):
    """Sends payslips via email to all employees."""
    success_count = 0
    total_employees = len(df)
    for _, row in df.iterrows():
        pdf_path = f"payslips/{row['Employee ID']}.pdf"
        if os.path.exists(pdf_path):
            if send_payslip_email(config, row, pdf_path):
                success_count += 1
        else:
            print(f"PDF not found for {row['Name']} (ID: {row['Employee ID']})")
    print(f"\nSummary:")
    print(f"Total employees: {total_employees}")
    print(f"Emails sent successfully: {success_count}")
    print(f"Failed emails: {total_employees - success_count}")

def load_config():
    """Loads email configuration settings."""
    config = {
        'SMTP_SERVER': 'smtp.gmail.com',
        'SMTP_PORT': 587,
        'FROM_EMAIL': "spliffking16@gmail.com",
        'EMAIL_PASSWORD': "bumu suqb dwcb zklp"
    }
    missing_keys = [key for key, value in config.items() if not value]
    if missing_keys:
        print(f"Missing required configuration: {', '.join(missing_keys)}")
        return None
    return config

# Entry point
if __name__ == "__main__":
    # Load configuration
    config = load_config()
    if config is None:
        sys.exit(1)

    # Sample employee data
    data = {
        'Employee ID': ['TC10', 'TC09', 'TC08', 'TC07', 'TC06', 'TC05', 'TC04', 'TC03', 'TC02', 'TC01', 'TC11', 'TC12'],
        'Name': ['Taps Ciriboto', 'Mbali Mushayi', 'Demy Bingura', 'Arty Sibanda', 'Emphraim Buruvuru',
                'Kudzai Tivatye', 'Reece Kurangwa', 'Lloyd Chogari', 'Thobela Sithole', 'Shaine Kaduhwa',
                'Nashe Graphix', 'Rue Gurure'],
        'Email Address': ['ciribotonicole@gmail.com', 'tafadzwamushayi3@gmail.com', 'demycadwell@gmail.com',
                        'sibandaaartii15@gmail.com', 'eoburuvuru@gmail.com', 'nicolaskudzai696@gmail.com',
                        'Kurangwareece@gmail.com', 'lloyddonnel44@gmail.com', 'thobelacarltonsithole@gmail.com',
                        'kaduhwashaine20@gmail.com', 'nashegraphix@gmail.com', 'ruvimbo448@gmail.com'],
        'Basic Salary': [2100.00, 2100.00, 1800.00, 1750.00, 1500.00, 2000.00, 2000.00, 2000.00,
                       2000.00, 2500.00, 1500.00, 1500.00],
        'Allowances': [500.00, 500.00, 450.00, 450.00, 450.00, 450.00, 450.00, 450.00,
                      450.00, 500.00, 450.00, 450.00],
        'Deductions': [95.00, 95.00, 45.00, 30.00, 40.00, 40.00, 50.00, 45.00,
                      45.00, 25.00, 15.00, 15.00]
    }
    
    df = pd.DataFrame(data)
    df['Net Salary'] = df['Basic Salary'] + df['Allowances'] - df['Deductions']
    
    generate_payslips(df)
    send_all_payslips(df, config)