### Mail Libaries ###

#pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib

from os import path
from pickle import load as pickleLoader
from pickle import dump as pickleDumper
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from email import encoders

from base64 import urlsafe_b64encode

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase



SCOPES = ['https://mail.google.com/']
senderEmail = 'jensh6247@gmail.com'
token_location="credentials\\token.pickle"
creds_location="credentials\\credentials.json"


def gmail_authenticate():
    creds = None
    if path.exists(token_location):
        with open(token_location, "rb") as token:
            creds = pickleLoader(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(creds_location, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_location, "wb") as token:
            pickleDumper(creds, token)
    return build('gmail', 'v1', credentials=creds)

service = gmail_authenticate()

def add_attachment(message, filename):
    file_path = open(filename, 'rb')
    attached_file = MIMEBase('application', 'vnd.ms-excel')
    attached_file.set_payload(file_path.read())
    file_path.close()
    encoders.encode_base64(attached_file)
    filename = path.basename(filename)
    attached_file.add_header('Content-Disposition', 'attachment', filename=filename)
    message.attach(attached_file)

def build_message(destination, obj, body, attachments=[]):
    message = MIMEMultipart()
    message['to'] = destination
    message['from'] = senderEmail
    message['subject'] = obj
    message.attach(MIMEText(body))
    for filename in attachments:
        add_attachment(message, filename)
    return {'raw': urlsafe_b64encode(message.as_bytes()).decode()}

def send_message(service, destination, obj, body, attachments=[]):
    return service.users().messages().send(
      userId="me",
      body=build_message(destination, obj, body, attachments)
    ).execute()


