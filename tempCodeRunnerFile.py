import pandas as pd
from fpdf import FPDF
import os
import yagmail
from getpass import getpass
from dotenv import load_dotenv

# --- Admin Gmail Login ---
print("Admin Gmail Login")
admin_email = input("Enter your Gmail address: ")
admin_password = getpass("Enter your Gmail app password: ")

# --- Load Employee Data ---
excel_file = r"C:\Users\uncommonStudent\OneDrive\Desktop\test\Payslip generator\payslip.xlsx"
df = pd.read_excel(excel_file)

# Expected columns: ['Employee ID', 'Name', 'Email', 'Basic Salary', 'Allowance', 'Deduction']

# --- Generate Payslip PDF ---
def generate_payslip_pdf(row):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    pdf.cell(200, 10, txt=f"Payslip for {row['Name']}", ln=True, align='C')
    pdf.ln(10)

    pdf.cell(100, 10, f"Employee ID: {row['Employee ID']}", ln=True)
    pdf.cell(100, 10, f"Basic Salary: {row['Basic Salary']}", ln=True)
    pdf.cell(100, 10, f"Allowance: {row['Allowance']}", ln=True)
    pdf.cell(100, 10, f"Deductions: {row['Deduction']}", ln=True)

    net_pay = row['Basic Salary'] + row['Allowance'] - row['Deduction']
    pdf.cell(100, 10, f"Net Pay: {net_pay}", ln=True)

    filename = f"payslip_{row['Employee ID']}.pdf"
    pdf.output(filename)
    return filename

# --- Send Payslip via Email ---
def send_payslip(row, filename):
    try:
        yag = yagmail.SMTP(user=admin_email, password=admin_password)
        subject = f"Payslip for {row['Name']}"
        body = f"Dear {row['Name']},\n\nPlease find attached your payslip for this month.\n\nBest regards,\nAdmin Team"
        yag.send(to=row['Email'], subject=subject, contents=body, attachments=filename)
        print(f"✅ Email successfully sent to {row['Name']} ({row['Email']})")
    except Exception as e:
        print(f"❌ Failed to send email to {row['Email']}: {e}")

# --- Process All Employees ---
for _, row in df.iterrows():
    pdf_file = generate_payslip_pdf(row)
    send_payslip(row, pdf_file)
    os.remove(pdf_file)  # Optional: remove the PDF after sending

print("\n🎉 All payslips processed and sent.")
