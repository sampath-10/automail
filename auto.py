import openpyxl
import smtplib
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import requests
from io import BytesIO

# Define the GitHub raw URL to the Excel file
github_raw_url = 'https://raw.githubusercontent.com/sampath-10/automail/main/Book12.xlsx'

# Fetch the Excel file from the GitHub URL
response = requests.get(github_raw_url)
if response.status_code == 200:
    file_content = BytesIO(response.content)
    workbook = openpyxl.load_workbook(file_content)
    sheet = workbook['Sheet1']
else:
    print("Failed to fetch the Excel file. Check the URL or your internet connection.")

today = datetime.today().strftime('%m-%d')

# Your email and password (consider using environment variables for security)
from_email = 'trailidsam@gmail.com'
password = 'sufapdhwpmytxyla'

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
        # ... Send the email using smtplib

        # Send reminder emails to others
        for row2 in sheet.iter_rows(values_only=True):
            name2, _, email2 = row2
            if email2 and email2 != email:
                subject2 = f"Today is {name}'s birthday!"
                message2 = f"Hi {name2},\n\nJust a reminder that today is {name}'s birthday. Don't forget to send your warm wishes!"
                # ... Send the email using smtplib

# Close the SMTP server and the Excel workbook
server.quit()
workbook.close()
