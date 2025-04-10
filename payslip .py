import pandas as pd
from fpdf import FPDF
import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv

# Load environment variables from .env
load_dotenv()

EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

# Ensure output directory exists
os.makedirs("payslips", exist_ok=True)

def generate_payslip(employee):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)

    # Header
    pdf.cell(0, 10, "Monthly Payslip", ln=True, align="C")
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, f"Employee ID: {employee['Employee ID']}", ln=True)
    pdf.cell(0, 10, f"Name: {employee['Name']}", ln=True)
    pdf.cell(0, 10, f"Basic Salary: ${employee['Basic Salary']:.2f}", ln=True)
    pdf.cell(0, 10, f"Allowances: ${employee['Allowances']:.2f}", ln=True)
    pdf.cell(0, 10, f"Deductions: ${employee['Deductions']:.2f}", ln=True)
    pdf.cell(0, 10, f"Net Salary: ${employee['Net Salary']:.2f}", ln=True)

    # Save the PDF to the payslips directory
    filename = f"payslips/{employee['Employee ID']}.pdf"
    pdf.output(filename)
    return filename

def send_email(to_email, filename, employee_name):
    try:
        # Create the email components
        subject = "Your Payslip for This Month"
        body = f"Dear {employee_name},\n\nPlease find attached your payslip for this month.\n\nBest regards,\nHR Department"
        
        # Email setup
        msg = MIMEMultipart()
        msg['From'] = EMAIL_ADDRESS
        msg['To'] = to_email
        msg['Subject'] = subject

        # Attach the email body
        msg.attach(MIMEText(body, 'plain'))

        # Attach the payslip file
        attachment = open(filename, 'rb')
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename={os.path.basename(filename)}")
        msg.attach(part)
        attachment.close()

        # Connect to the SMTP server and send the email
        server = smtplib.SMTP('smtp.gmail.com', 587)  # Use the appropriate SMTP server and port
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        server.sendmail(EMAIL_ADDRESS, to_email, msg.as_string())
        server.quit()

        print(f"Email sent to {employee_name} at {to_email}")
    except Exception as e:
        print(f"Failed to send email to {employee_name}: {e}")

def main():
    try:
        # Load the employee data from the Excel file
        df = pd.read_excel("employees.xlsx")

        # Fill missing values with 0 for calculations
        df.fillna(0, inplace=True)

        # Compute net salary for each employee
        df["Net Salary"] = df["Basic Salary"] + df["Allowances"] - df["Deductions"]

        # Iterate over each employee and send their payslip
        for _, row in df.iterrows():
            payslip_file = generate_payslip(row)  # Generate the payslip PDF
            send_email(row["Email"], payslip_file, row["Name"])  # Send the email

        print("All payslips processed and sent successfully.")
    except Exception as e:
        print(f"Error during processing: {e}")

if __name__ == "__main__":
    main()