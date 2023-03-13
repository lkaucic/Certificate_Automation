# AUTOMATIC CERTIFICATOR
# ...................................
# Given .xlsx file with form responses, and certificate template, 
# this script generates certificates and sends them to all recipients via SMTP
# ...................................
# AUTHOR: LUKA KAUCIC
# DATE: 2023-03-12
# CONTACT: luka.kaucic08@gmail.com
# ...................................
# IMPORT SECTION
# ...................................
import pandas as pd
import numpy as np
from docx import Document
from docxtpl import DocxTemplate
import smtplib
from pathlib import Path
import ssl
import os
from email.message import EmailMessage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
from email import encoders
# ...................................
# GENERATE CERTIFICATES FROM MS FORM
# ...................................
form_responses = pd.read_excel("//Downloads//Microsoft_Form_results.xlsx")    # Path to donwloaded Microsoft form report
names = form_responses["Name and Surname"].values  # This corresponds to first column of response we want to store
emails = form_responses["e-mail"].values # This corresponds to second column of respon se we want to store
info = list(zip(names,emails))
certificate = DocxTemplate("//Event Certificate Template.docx") # Path to certificate template (.docx document)
for attendee in info:
    context = {
    'student_name': attendee[0],
    }
    certificate.render(context)
    certificate.save(f"//Filled Certificates//Certifikat_{attendee[0]}.docx") #Path where filled certificates will be stored
# ...................................
# SEND E-MAILS WITH CERTIFICATES
# ...................................
smtp_port = 587                 # Standard secure SMTP port
smtp_server = "smtp.gmail.com"  # Google SMTP Server
sender = "my_email@gmail.com" # MY GMAIL
password = "password" # APP PASSWORD
recievers = [email for name, email in info]
subject = "MSA workshop certificate"
body = f"""
        This mail contains certificate for the following workshop:
        'Name of the workshop' 
        
        Best regards,
        my_name
       """
for person in info:
    
    msg = MIMEMultipart()
    msg['From'] = sender
    msg['To'] = person[1]
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))
    filepath = f"//Filled Certificates//Certifikat_{person[0]}.docx"   # path to attachment
    attachment= MIMEApplication(open(filepath, "rb").read())  # r for read and b for binary
    attachment.add_header('Content-Disposition', 'attachment', filename=f"Certifikat_{person[0]}.docx")
    msg.attach(attachment)



    # Cast as string
    text = msg.as_string()
    print("Connecting to server...")
    TIE_server = smtplib.SMTP(smtp_server, smtp_port)
    TIE_server.starttls()
    TIE_server.login(sender, password)
    print("Succesfully connected to server")
    print()


    # Send emails to "person" as list is iterated
    print(f"Sending email to: {person[0]}...")
    TIE_server.sendmail(sender, person[1], text)
    print(f"Email sent to: {person[0]}, {person[1]}")
    print()

# Close the port
TIE_server.quit()