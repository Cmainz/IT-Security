from re import sub as substractor
from re import compile as compiler
from pickle import load as pickleLoader
from src.mailAPI import service
from base64 import urlsafe_b64decode
from os import path,getcwd

sendersDict = dict()
titleList = []
emailSender = []
downloadableMsg = []


def previousControl():
  sentEmails = pickleLoader(open("Missingcontrols.dat", "rb"))
  for controls in sentEmails:
    for item in controls:
      title = str(item[0]) + " " + item[1] + " " + str(item[2])
      titleList.append(title)
      sender = item[4]
      print(sender)
      emailSender.append(sender)

  return print(titleList)

def findingmsgID():
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
      modTitle = substractor(pattern, "", title)
      if modTitle in titleList and reciever in emailSender:
        msgId = (msg["id"])
        downloadableMsg.append(msgId)
        print(modTitle)
  return print(downloadableMsg)


def downloadAttachment():
  filename = ""
  for item in downloadableMsg:
    message = service.users().messages().get(userId='me', id=item).execute()
    try:
      attId = message['payload']['parts'][1]['body']['attachmentId']
      att = service.users().messages().attachments().get(userId='me', messageId=message['id'], id=attId).execute()
      data = att['data']
      file_data = urlsafe_b64decode(data.encode('UTF-8'))
      filename = message['payload']['parts'][1]['filename']
      Dlpath = path.join(getcwd() + '\\' 'Downloaded controls' + '\\' + filename)
      print("download")
      with open(Dlpath, 'wb') as f:
        f.write(file_data)
        f.close()
      service.users().messages().delete(userId='me', id=item).execute()
      
    except KeyError:
      service.users().messages().delete(userId='me', id=item).execute()
      return print("No attachments in Control. Control will be deleted")
    
  
  return print(filename)


##LOGIC##
previousControl()
findingmsgID()
downloadAttachment()
