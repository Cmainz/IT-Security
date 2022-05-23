from re import sub as substractor
from re import compile as compiler
from pickle import load as pickleLoader
from src.mailAPI import service
from base64 import urlsafe_b64decode
from os import path,getcwd

senders_dict = dict()
title_List = []
email_sender = []
downloadable_Msg = []


def previous_control():
  sent_emails = pickleLoader(open("Missingcontrols.dat", "rb"))
  for controls in sent_emails:
    for item in controls:
      title = str(item[0]) + " " + item[1] + " " + str(item[2])
      title_List.append(title)
      sender = item[4]
      print(sender)
      email_sender.append(sender)

  return print(title_List)

def finding_msg_id():
  results = service.users().messages().list(userId='me', labelIds=['INBOX']).execute()
  messages = results.get('messages', [])

  pattern = compiler(r'Fwd: |FWD: |re: |Re: ')

  if not messages:
    return(print("No messages found."))
  else:
    for message in messages:
      msg = service.users().messages().get(userId='me', id=message['id']).execute()
      title = msg['payload']['headers'][21]['value']
      reciever = msg['payload']['headers'][6]['value'][1:-1]
      mod_title = substractor(pattern, "", title)
      if mod_title in title_List and reciever in email_sender:
        msg_id = (msg["id"])
        downloadable_Msg.append(msg_id)
        print(mod_title)
  return print(downloadable_Msg)


def download_attachment():
  filename = ""
  for item in downloadable_Msg:
    message = service.users().messages().get(userId='me', id=item).execute()
    try:
      att_id = message['payload']['parts'][1]['body']['attachmentId']
      att = service.users().messages().attachments().get(userId='me', messageId=message['id'], id=att_id).execute()
      data = att['data']
      file_data = urlsafe_b64decode(data.encode('UTF-8'))
      filename = message['payload']['parts'][1]['filename']
      dl_path = path.join(getcwd() + '\\' + 'Downloaded controls' + '\\' + filename)
      print("download")
      with open(dl_path, 'wb') as f:
        f.write(file_data)
        f.close()
      service.users().messages().delete(userId='me', id=item).execute()
      
    except KeyError:
      service.users().messages().delete(userId='me', id=item).execute()
      return print("No attachments in Control. Control will be deleted")
    
  
  return print(filename)


##LOGIC##
previous_control()
finding_msg_id()
download_attachment()
