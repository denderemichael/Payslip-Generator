import pandas as pd
from fpdf import FPDF
import os
import yagmail
from getpass import getpass
from dotenv import load_dotenv

# --- Admin Gmail Login ---
print("🔐 Admin Gmail Login")
admin_email = input("Enter your Gmail address: ")
admin_password = getpass("Enter your Gmail app password: ")  # App password only

# --- Load and Clean Excel Data ---
excel_file = r"C:\Users\uncommonStudent\OneDrive\Desktop\test\Payslip generator\payslip.xlsx"

try:
    df = pd.read_excel(excel_file)
    df.columns = df.columns.str.strip()  # Remove extra spaces from column names
    print("✅ Excel file loaded successfully.")
    print("📄 Columns found:", df.columns.tolist())  # For debugging
except Exception as e:
    print(f"❌ Failed to read Excel file: {e}")
    exit()

# --- Generate Payslip PDF ---
def generate_payslip_pdf(row):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    pdf.cell(200, 10, txt=f"Payslip for {row['Name']}", ln=True, align='C')
    pdf.ln(10)

    pdf.cell(100, 10, f"Employee ID: {row['Employees ID']}", ln=True)
    pdf.cell(100, 10, f"Basic Salary: {row['Basic  Salary']}", ln=True)
    pdf.cell(100, 10, f"Allowance: {row['Allowance']}", ln=True)
    pdf.cell(100, 10, f"Deductions: {row['Deduction']}", ln=True)

    net_pay = row['Basic  Salary'] + row['Allowance'] - row['Deduction']
    pdf.cell(100, 10, f"Net Pay: {net_pay}", ln=True)

    filename = f"payslip_{row['Employees ID']}.pdf"
    pdf.output(filename)
    return filename

# --- Send Email with Payslip ---
def send_payslip(row, filename):
    try:
        yag = yagmail.SMTP(user=admin_email, password=admin_password)
        subject = f"Payslip for {row['Name']}"
        body = f"""Dear {row['Name']},

Please find attached your payslip for this month.

Best regards,  
Admin Team"""
        yag.send(to=row['Email'], subject=subject, contents=body, attachments=filename)
        print(f"✅ Email successfully sent to {row['Name']} ({row['Email']})")
    except Exception as e:
        print(f"❌ Failed to send email to {row['Email']}: {e}")

# --- Process Each Employee ---
for index, row in df.iterrows():
    try:
        pdf_file = generate_payslip_pdf(row)
        send_payslip(row, pdf_file)
        os.remove(pdf_file)  # Clean up temporary PDF
    except Exception as err:
        print(f"❌ Error processing {row.get('Name', 'Unknown')}: {err}")

print("\n🎉 All payslips processed and emails sent (or attempted).")


# pepd qczr mtgx wrzx