#!/usr/bin/env python

import smtplib
from datetime import datetime
import openpyxl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Load the Excel sheet with names and DOB
excel_file = 'Book21.xlsx'
workbook = openpyxl.load_workbook(excel_file)
sheet = workbook['Sheet1']

# Define your email settings
email_address = 'trailidsam@gmail.com'
email_password = 'sufapdhwpmytxyla'

# Connect to the SMTP server
smtp_server = 'smtp.gmail.com'  # Update this for your email provider
smtp_port = 587  # Update this for your email provider

server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()
server.login(email_address, email_password)

# Get today's date
today = datetime.today().strftime('%m-%d')

# Iterate through the Excel sheet
for row in sheet.iter_rows(min_row=2, values_only=True):
    name, dob = row

    if today == dob:
        # It's their birthday, send an email
        message = MIMEMultipart()
        message['From'] = email_address
        message['To'] = 'sampathgaming04@gmail.com'
        message['Subject'] = f'Happy Birthday, {name}!'

        # Customize the email body as you like
        body = f"Dear {name},\n\nHappy Birthday!\n\nBest wishes,\nVANQUISHERS"
        message.attach(MIMEText(body, 'plain'))

        # Send the email
        server.sendmail(email_address, email_address, message.as_string())
        print(f"mail sent to {name} sucessfully ")

# Quit the SMTP server
server.quit()
