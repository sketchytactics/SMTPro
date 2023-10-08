import smtplib as stp
import tkinter as tk
from tkinter import filedialog
import os
import pandas as pd
import yaml
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


report_log = open('Email_Processing_Report.txt', 'w')

root = tk.Tk()
root.withdraw()

with open(os.path.join(os.getcwd(), 'SMTProConfig', 'config.yaml'), 'r') as yaml_file:
    config = yaml.safe_load(yaml_file)

# email configs
smtp_server = config['smtp_server']
smtp_port = config['smtp_port']
smtp_username = config['smtp_user']
smtp_password = config['smtp_password']
sender_email = config['smtp_sender']
server = None

outbox_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
outbox_df = pd.read_excel(outbox_path, sheet_name=0)
outbox_df['Receiver CC'] = outbox_df['Receiver CC'].fillna(False)
outbox_df['Attachment Name'] = outbox_df['Attachment Name'].fillna(False)

for index, row in outbox_df.iterrows():
    invoice = row['Invoice']
    receiver_email = row['Receiver Email']
    receiver_cc = row['Receiver CC']
    subject = row['Subject']
    greeting = row['Greeting']
    email_body1 = row['Email Body 1']
    email_body2 = row['Email Body 2']
    signature1 = row['Signature 1']
    signature2 = row['Signature 2']
    attachment_name = row['Attachment Name']

# message config
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = receiver_email
    if receiver_cc:
        message['Cc'] = receiver_cc

    message['Subject'] = subject

# body config
    body_text = f'{greeting}\n\n{email_body1}\n\n{email_body2}\n\n{signature1}\n\n{signature2}'
    message.attach(MIMEText(body_text, 'plain'))

# attachment config
    if attachment_name:
        file_path = os.path.join('OutboxAttachments', attachment_name)
        if not os.path.exists(file_path):
            error_msg = f'{attachment_name} not found in outbox'
            report_log.write(error_msg + '\n')
        else:
            with open(file_path, 'rb') as attachment:
                part = MIMEApplication(attachment.read())
                part.add_header('Content-Disposition', f'attachment; filename={attachment_name}')
                message.attach(part)

# email execution
    try:
        server = stp.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.sendmail(sender_email, receiver_email, message.as_string())
        status = f'{invoice} Successful'
        report_log.write(status + '\n')
    except Exception as e:
        status = f'{invoice} ERROR: {str(e)}'
        report_log.write(status + '\n')

if server:
    server.quit()

report_log.close()
yaml_file.close()
