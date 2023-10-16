import openpyxl
import smtplib
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os

# Specify the path to the Excel file within your repository
excel_file_path = 'Book12.xlsx'

# Load the Excel workbook
try:
    workbook = openpyxl.load_workbook(excel_file_path)
except FileNotFoundError:
    print(f"Excel file '{excel_file_path}' not found.")
    # Handle the error or exit gracefully

sheet = workbook['Sheet1']

today = datetime.today().strftime('%m-%d')

# Your email and password (consider using environment variables for security)
from_email = os.environ.get('trailidsam@gmail.com')  # Use environment variables to protect your email and password
password = os.environ.get('sufapdhwpmytxyla')

# Initialize the SMTP server
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(from_email, password)

# Iterate through the Excel sheet and send emails
for row in sheet.iter_rows(values_only=True):
    name, dob_str, email = row
    if today == dob_str:
        # Send birthday email
        subject = 'Happy Birthday!'
        message = f"Dear {name},\n\nHappy Birthday! 🎉🎂\n\nBest wishes, Your Name"

        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = email
        msg['Subject'] = subject

        msg.attach(MIMEText(message, 'plain'))

        server.sendmail(from_email, email, msg.as_string())

# Close the SMTP server and the Excel workbook
server.quit()
workbook.close()
