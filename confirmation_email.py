import smtplib
import ssl
from email.message import EmailMessage
import pandas as pd

df = pd.read_excel(r"Confirmation_list.xlsx")  # Replace with your Excel file's path
email_addresses = df['Email'].tolist()

email_sender = ''
email_password = ''

subject = 'Test Mail'
body = """
It's Working
"""

# Add SSL (layer of security)
context = ssl.create_default_context()

with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
    smtp.login(email_sender, email_password)

    for email_receiver in email_addresses:
        em = EmailMessage()
        em['From'] = email_sender
        em['Subject'] = subject
        em.set_content(body)
        em['Bcc'] = email_receiver 

        smtp.send_message(em)
        print(f'Email sent successfully to {email_receiver}')

print('All emails sent successfully!')
