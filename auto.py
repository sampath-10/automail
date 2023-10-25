#!/usr/bin/env python
import openpyxl
import smtplib
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart # Import BytesIO to work with file content in memory
fp = r'Book21.xlsx'
workbook = openpyxl.load_workbook(fp)
sheet = workbook['Sheet1']
today = datetime.today().strftime('%m/%d')
from_email = 'trailidsam@gmail.com'
password = 'sufapdhwpmytxyla'
to_email = 'sampathgaming04@gmail.com'
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(from_email, password)
for row in sheet.iter_rows(values_only=True):
    name, dob_str = row
    print(dob_str==today)
    if today == dob_str:
        subject = 'Happy Birthday!'
        message = f"Today is {name} 's Birthday! 🎉🎂\n\nBest wishes, Vanquishers"
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(message, 'plain'))
        server.sendmail(from_email, to_email, msg.as_string())
        print(f" {name} Birthday email sent to  ({to_email})")
        
server.quit()
workbook.close()
