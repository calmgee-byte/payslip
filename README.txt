===============================
Payslip Generator & Email Sender
===============================

Author: Carlton Thobela Sithole
Date: April 2025

DESCRIPTION:
------------
This Python project automates the generation of payslips from an Excel file and sends them to employees via email. It creates PDF payslips based on salary data, saves them to a "payslips" folder, and emails each employee their payslip as an attachment.

----------------------------------------
MAIN FEATURES:
----------------------------------------

1. Reads employee data from an Excel file (payslip project.xlsx).
2. Calculates net salary using the formula:
     Net Salary = Basic Salary + Allowances - Deductions
3. Creates a PDF payslip for each employee using ReportLab.
4. Emails each payslip to the corresponding employee using SMTP.

----------------------------------------
REQUIREMENTS:
----------------------------------------

Python 3.10+  
Install dependencies using pip:

pip install pandas openpyxl fpdf reportlab

For email functionality, make sure your SMTP settings (email and password) are correctly configured.

----------------------------------------
FILE STRUCTURE:
----------------------------------------

- payslip.py              --> Main script to generate and send payslips
- payslip project.xlsx    --> Excel file with employee salary data
- payslips/               --> Folder where generated PDF payslips are stored
- README.txt              --> This file

----------------------------------------
HOW TO RUN:
----------------------------------------

1. Open terminal or PowerShell and navigate to this folder:
   cd "C:\Users\uncommonstudent\OneDrive\Desktop\thobs"

2. Run the script using:
   python payslip.py

3. The script will:
   - Read the Excel file
   - Generate PDF payslips for each employee
   - Send emails with the payslips attached

----------------------------------------
EXCEL FILE FORMAT (Required Columns):
----------------------------------------

Your Excel sheet should contain the following columns:

- Employee ID
- Name (or Name with a trailing space, depending on the original file)
- Email Address
- Basic Salary
- Allowances
- Deductions

Make sure column headers are correctly spelled to match the script!

----------------------------------------
EMAIL SETUP:
----------------------------------------

The script uses a Gmail SMTP server by default.

To use Gmail:
- Enable "Less secure app access" in your Gmail account or use an App Password.
- Set the following environment variables or directly edit them in the script:

SMTP_SERVER = smtp.gmail.com  
SMTP_PORT = 587  
FROM_EMAIL = your_email@gmail.com  
EMAIL_PASSWORD = your_app_password

----------------------------------------
IMPORTANT NOTES:
----------------------------------------

- Ensure your "payslip project.xlsx" file is in the same folder as payslip.py.
- Make sure all employees have valid email addresses listed.
- Review PDF files inside the "payslips" folder before sending.

----------------------------------------
CONTACT:
----------------------------------------

If you run into any issues or need help improving this project,
reach out to Carlton at: thobelacarltonsithole@gmail.com

Enjoy automating your payroll!

