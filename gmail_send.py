import os.path
import base64
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# Gmail API scope for sending mail
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

def authenticate_gmail():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    else:
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('gmail', 'v1', credentials=creds)

def create_message_with_attachment(sender, to, subject, body_html, file_path):
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject

    # Attach the HTML body
    message.attach(MIMEText(body_html, 'html'))

    # Attach the PDF file
    with open(file_path, 'rb') as f:
        part = MIMEApplication(f.read(), _subtype='pdf')
        part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file_path))
        message.attach(part)

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    return {'raw': raw}

def main():
    sender_email = "tps039@gmail.com"
    data = pd.read_excel("contacts.xlsx")
    file_path = "Resume.pdf"

    # Read HTML body from txt file
    with open("email_body.txt", "r", encoding="utf-8") as f:
        body_template = f.read()

    service = authenticate_gmail()

    for _, row in data.iterrows():
        name = row['Name']
        email = row['Email']
        company = row['Company']

        subject = f"Application for Generative AI Intern Role"
        body_filled = body_template.format(name=name, company=company)

        message_obj = create_message_with_attachment(sender_email, email, subject, body_filled, file_path)
        try:
            service.users().messages().send(userId="me", body=message_obj).execute()
            print(f"✅ Sent to {email}")
        except Exception as error:
            print(f"❌ Failed to send to {email}: {error}")


if __name__ == '__main__':
    main()
