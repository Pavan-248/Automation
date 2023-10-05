import smtplib
import ssl
from email.message import EmailMessage
import pandas as pd

# Read email addresses from an Excel file
df = pd.read_excel(r"C:\Users\pavan\OneDrive\Documents\Confirmation_list.xlsx")  # Replace with your Excel file's path
email_addresses = df['Email'].tolist()

# Define email sender
email_sender = ''
email_password = ''

# Set the subject and body of the email
subject = 'TEDxSIST Test Mail'
body = """
It's Working
"""

# Add SSL (layer of security)
context = ssl.create_default_context()

# Log in and send separate emails to each recipient using Bcc
with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
    smtp.login(email_sender, email_password)

    for email_receiver in email_addresses:
        em = EmailMessage()
        em['From'] = email_sender
        em['Subject'] = subject
        em.set_content(body)
        em['Bcc'] = email_receiver  # Set the "Bcc" header for this specific recipient

        smtp.send_message(em)
        print(f'Email sent successfully to {email_receiver}')

print('All emails sent successfully!')
