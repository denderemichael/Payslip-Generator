import pandas as pd
from fpdf import FPDF
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import getpass
import os

# ========== USER INPUT ==========
sender_email = input("Enter your Gmail address: ")
sender_password = getpass.getpass("Enter your Gmail app password: ")

# ========== LOAD EXCEL DATA ==========
df = pd.read_excel(r"C:\Users\uncommonStudent\OneDrive\Desktop\test\employee\payslip.xlsx")
df.columns = df.columns.str.strip().str.replace(r'\s+', ' ', regex=True)
print("✅ Cleaned Columns:", df.columns.tolist())

# ========== CALCULATE NET SALARY ==========
df["Net Salary"] = df["Basic Salary"] + df["Allowance"] - df["Deduction"]
print("\n✅ With Net Salary:\n", df[["Employees ID", "Name", "Net Salary"]])

# ========== CREATE OUTPUT FOLDER ==========
output_folder = "payslips"
os.makedirs(output_folder, exist_ok=True)

# ========== EMAIL CONFIG ==========
subject = "Your Monthly Payslip"
base_body = "Please find attached your payslip for this month.\n\nRegards,\nHR Department"

# ========== INITIATE PAYSLIP LOGGING ==========
log_entries = []

# ========== PROCESS EACH EMPLOYEE ==========
for index, row in df.iterrows():
    name = row["Name"]
    receiver_email = row["Email"]
    basic = row["Basic Salary"]
    allowance = row["Allowance"]
    deductions = row["Deduction"]
    net = row["Net Salary"]

    # ========== CREATE PDF ==========
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt="Monthly Payslip", ln=True, align="C")
    pdf.ln(10)
    pdf.cell(200, 10, txt=f"Name: {name}", ln=True)
    pdf.cell(200, 10, txt=f"Basic Salary: ${basic:.2f}", ln=True)
    pdf.cell(200, 10, txt=f"Allowance: ${allowance:.2f}", ln=True)
    pdf.cell(200, 10, txt=f"Deductions: ${deductions:.2f}", ln=True)
    pdf.cell(200, 10, txt=f"Net Salary: ${net:.2f}", ln=True)

    # Create employee folder and file path
    employee_folder = os.path.join(output_folder, name)
    os.makedirs(employee_folder, exist_ok=True)
    filename = f"{name}.pdf"
    filepath = os.path.join(employee_folder, filename)
    pdf.output(filepath)

    print(f"📄 Generated payslip for {name}")

    # ========== COMPOSE EMAIL ==========
    personalized_body = f"Dear {name},\n\n{base_body}"
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message.attach(MIMEText(personalized_body, "plain"))

    # ========== ATTACH PDF ==========
    if os.path.exists(filepath):
        with open(filepath, "rb") as f:
            part = MIMEApplication(f.read(), _subtype="pdf")
            part.add_header('Content-Disposition', 'attachment', filename=filename)
            message.attach(part)
        print(f"📎 Attached PDF for {name}")
    else:
        print(f"❌ PDF not found for {name}")
        continue

    # ========== SEND EMAIL ==========
    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, receiver_email, message.as_string())
        print(f"✅ Email sent to {name} ({receiver_email})\n")

        # Log success
        log_entries.append({
            "Name": name,
            "Email": receiver_email,
            "PDF Path": filepath,
            "Status": "Sent"
        })
    except Exception as e:
        print(f"❌ Error sending email to {name}: {e}\n")
        log_entries.append({
            "Name": name,
            "Email": receiver_email,
            "PDF Path": filepath,
            "Status": f"Failed: {e}"
        })

# ========== SAVE DATABASE FILE ==========
log_df = pd.DataFrame(log_entries)
log_df.to_csv("payslip_log.csv", index=False)
print("🗃️ Payslip log saved as payslip_log.csv")
