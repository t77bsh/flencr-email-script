import os
import instaloader
import pandas as pd
from openpyxl import load_workbook
import re

# Function to load existing data from Excel and avoid duplicates
def load_existing_data(file_path):
    try:
        existing_data = pd.read_excel(file_path)
        existing_emails = existing_data['email'].tolist()
    except FileNotFoundError:
        existing_data = pd.DataFrame(columns=['username', 'real_name', 'followers_count', 'email'])
        existing_emails = []
    return existing_data, existing_emails

# Function to save data to Excel
def save_to_excel(data, file_path):
    existing_data, existing_emails = load_existing_data(file_path)
    new_data = pd.DataFrame(data)
    combined_data = pd.concat([existing_data, new_data])
    combined_data.drop_duplicates(subset='email', keep='first', inplace=True)
    combined_data.to_excel(file_path, index=False)

# Function to extract email from Instagram profile
def extract_email(profile):
    email = None

    # Regular expression to find valid email addresses
    email_regex = re.compile(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}')

    # Check biography for email
    if profile.biography:
        # Find all matches of the regex in the biography
        emails_in_bio = email_regex.findall(profile.biography)
        if emails_in_bio:
            email = emails_in_bio[0]  # Take the first found email

    # Fallback to business email if no email found in the bio
    if not email:
        email = profile.business_email

    return email

# Main script to extract data from Instagram
def main():
    # Login to Instagram
    L = instaloader.Instaloader()
    username = 'flencr_app'
    password = os.getenv('INSTAGRAM_PASSWORD')
    L.login(username, password)

    # Load profile
    profile = instaloader.Profile.from_username(L.context, username)

    leads_data = []
    for followee in profile.get_followees():
        email = extract_email(followee)
        if email:
            leads_data.append({
                'username': followee.username,
                'real_name': followee.full_name,
                'followers_count': followee.followers,
                'email': email
            })

    save_to_excel(leads_data, 'leads_data.xlsx')

if __name__ == "__main__":
    main()
