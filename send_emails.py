import os
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import unidecode
import emoji
import re

# Function to clean text
def clean_text(text):
    try:
        text = unidecode.unidecode(text)  # Normalize to ASCII
        text = emoji.replace_emoji(text, replace='')  # Remove emojis
        text = re.sub(r'[^A-Za-z0-9 ]+', '', text)  # Remove special characters, keeping letters, numbers, and spaces
        first_name = text.split()[0]  # Extract the first name
        return first_name.capitalize()  # Capitalize the first letter
    except Exception as e:
        print(f"Error cleaning text: {text}. Error: {e}")



# Read the Excel file
# file_path = 'IGEmail_email_flencr_app_32_following.xlsx'  # Make sure this is the correct path to your Excel file
file_path = 'testing123.xlsx'  # Make sure this is the correct path to your Excel file
df = pd.read_excel(file_path)

# Clean the 'Real Name' column
df['Fullname'] = df['Fullname'].apply(clean_text)


# Extract necessary columns
df_filtered = df[['Fullname', 'Followers', 'Public Email']]

def generate_email_content(row):
    email_template = os.getenv('EMAIL_CONTENT')
    return email_template

df_filtered['email_content'] = df_filtered.apply(generate_email_content, axis=1)

# Set up your email server and login details
smtp_server = 'smtp.mail.me.com'  # iCloud SMTP server
smtp_port = 587  # Port number
email_address = 'taabish@flencr.com'  # Your iCloud email address
primary_email = os.getenv('EMAIL_ADDRESS')  # Your primary email address
email_password = os.getenv('EMAIL_PASSWORD')  # Your iCloud email password

def send_email(to_address, subject, content):
    msg = MIMEMultipart()
    msg['From'] = email_address
    msg['To'] = to_address
    msg['Subject'] = subject
    msg['Bcc'] = email_address  # Add BCC to ensure it appears in the Sent folder
    msg.attach(MIMEText(content, 'plain'))

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(primary_email, email_password)
            server.send_message(msg)
        print(f"Email sent successfully to {to_address}")
    except smtplib.SMTPAuthenticationError as e:
        print(f"Failed to authenticate: {e}")
    except Exception as e:
        print(f"Failed to send email to {to_address}: {e}")

# Send emails
for index, row in df_filtered.iterrows():
    send_email(row['Public Email'], f"{row['Fullname']}?", row['email_content'])
