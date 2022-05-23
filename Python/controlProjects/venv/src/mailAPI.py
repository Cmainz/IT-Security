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
from mimetypes import guess_type as guess_mime_type



SCOPES = ['https://mail.google.com/']
senderEmail = 'jensh6247@gmail.com'

def gmail_authenticate():
    creds = None
    if path.exists("credentials\\token.pickle"):
        with open("credentials\\token.pickle", "rb") as token:
            creds = pickleLoader(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials\\credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open("credentials\\token.pickle", "wb") as token:
            pickleDumper(creds, token)
    return build('gmail', 'v1', credentials=creds)

service = gmail_authenticate()

def add_attachment(message, filename):
    filePath = open(filename, 'rb')
    attachedFile = MIMEBase('application', 'vnd.ms-excel')
    attachedFile.set_payload(filePath.read())
    filePath.close()
    encoders.encode_base64(attachedFile)
    filename = path.basename(filename)
    attachedFile.add_header('Content-Disposition', 'attachment', filename=filename)
    message.attach(attachedFile)

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


