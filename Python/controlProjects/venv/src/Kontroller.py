#!/usr/bin/env python3

from openpyxl import load_workbook
from os import getcwd, walk
from shutil import copyfile
from datetime import date
from src.mailAPI import send_message, service
import pickle

### Variables ###

sheet = load_workbook(filename="mainControllerDoc\\Kontroller.xlsx")
wsCtrl = sheet["Controls"]

wsControllers = sheet["Controllers"]

today = date.today()
todayDate = int(today.strftime("%d%m%Y"))

conInfo = {}
filesToSend = []
mailingList = []

maxControlRow = len(wsCtrl['A'])
maxContactsRow = len(wsControllers['A'])

### Classes ###


class Ctrls:
  
  ctrlsList=[]
  
  def __init__(self, number, control, due, verification,responsible):
    self.number = number
    self.control = control
    self.due = due
    self.verification = verification
    self.responsible = responsible
    
    
    self.ctrlsList.append(self)
    
  

### Functions ###

def contactInfoFunc():
  for value, val in wsControllers.iter_rows(
    min_row=2,
    max_row=maxContactsRow,
    min_col=1,
    max_col=2,
    values_only=True):
    if value is None:
      pass
    else:
      conInfo[value] = val
  return conInfo

def createCtrls():
  for value in wsCtrl.iter_rows(min_row=2,
                            max_row=maxControlRow,
                            min_col=1,
                            max_col=5,
                            values_only=True):
   
    if value[3] == "X":
      continue
      
    else:
      
      Ctrls(value[0],
            value[1],
            value[2],
            value[3],
            value[4]
            )

def classMaker(listItem):
  global contactInfo
  for item in listItem:
    checkForDue(item.number,
              item.control,
              item.due,
              item.verification,
              item.responsible)

  return mailingList


      
def checkForDue(value0, value1, value2, value3, value4,todayDate=todayDate) -> str:

  global notes
  
  try:
    type(int(value0)) == int
    dueDate=value2.date()
    dueDateint = int(dueDate.strftime("%d%m%Y"))
    if dueDateint - todayDate == 0:
      notes = "Send The email!"
      sendEmail = True

    elif dueDateint - todayDate == 10000000: # send a reminder if 10 days is left
      notes = "Send a reminder! He got 10 days left"
      sendEmail = True

    elif dueDateint - todayDate == 5000000: # send a reminder if 5 days is left
      notes = "Send a reminder!"
      sendEmail = True

    elif dueDateint - todayDate == -1000000: # send a reminder if delayed 1 day
      notes = "You are late! Please finish this control before end of date"
      sendEmail = True

    elif dueDateint - todayDate == -2000000: # send a reminder if delayed 2 days
      notes = "You are late!"
      sendEmail = True

    elif dueDateint - todayDate == -3000000: # send an email to security responsible
      notes = "this control has not been finished in time or has been incorrectly made."
      sendEmail = True

    else:
      sendEmail=False
      notes ="Nothing will be done"
      
    if value4 not in conInfo or conInfo[value4] == None:
      print("Update needed for {}".format(value4))
      return "Missing Contact Information"
    elif sendEmail == True:
      mailingList.append([value0, value1, dueDate, value3, conInfo[value4],notes])
    else:
      pass
          
  except ValueError:
    print("Something went wrong with your data.\n\"{}\" is not an index number. Control "
          "\"{}\" will not be correctly analysed \nPlease check your Excel Sheet".format(value0,value1))
    return "\"{}\" is not an index number. Control " \
           "\"{}\" will not be correctly analysed \nPlease check your " \
             "Excel Sheet".format(value0,value1)
  
  return notes


def makeControlDoc():
  for item in mailingList:
    control = item[1] + ".xlsx"
    control_title = str(item[0]) + " " + item[1] + " " + str(item[2]) + ".xlsx"
    
    for roots, dirs, files in walk("."):
      for file in files:
        if control in file:
          original = getcwd() + "\Templates" + "\\" + file
          target = getcwd() + "\Temps" + "\\" + control_title
          copyfile(original, target)
          filesToSend.append(target)
  return filesToSend

def sendingEmail():
  for item, file in zip(mailingList, filesToSend):
    send_message(service, "chr.maints@gmail.com", str(item[0])+" "+item[1]+" "+str(item[2]), item[5], [file])

### Logic ###


if __name__ == "__main__":
  contactInfoFunc() #Looks through the contactlists in Excel
  createCtrls() #Looks through all controls listed in Excel
  classMaker(Ctrls.ctrlsList) #transforms our list into forms and runs them through checkForDue()
  makeControlDoc() #Create conntrols from templates

  sendingEmail() #Sends emails
  pickle.dump(zip(mailingList),open("Missingcontrols.dat","wb")) # saving the emails for later usage