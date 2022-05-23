import unittest
import os
os.chdir(".."+"\\src")
from Kontroller import *
from datetime import datetime,date


class checkforDueUnittest(unittest.TestCase):

  def testForValueError(self) -> None:
    message=checkForDue("testIndex","controlTest",2,3,4)
    expected= "\"testIndex\" is not an index number. Control \"controlTest\" will not be correctly analysed \nPlease check your Excel Sheet"
    self.assertEqual(expected,message,"Error")

  def testForZeroDay(self):
    testDate = 19042021
    contactInfoFunc()
    datetimestr=datetime.strptime("19.04.2021", "%d.%m.%Y")
    message = checkForDue(1, "testControlZero", datetimestr, "","Jens Hansen",todayDate=testDate)
    expected = "Send The email!"
    self.assertEqual(expected, message)

  def testForFiveDays(self):
    testDate = 19042021
    contactInfoFunc()
    datetimestr = datetime.strptime("24.04.2021", "%d.%m.%Y")
    message = checkForDue(2, "testControlFive", datetimestr, "", "Jens Hansen", todayDate=testDate)
    expected = "Send a reminder!"
    self.assertEqual(expected, message)

  def testForTenDays(self):
    testDate=19042021
    contactInfoFunc()
    datetimestr = datetime.strptime("29.04.2021", "%d.%m.%Y")
    message =checkForDue(3,"testControlTen",datetimestr,"","Jens Hansen",todayDate=testDate)
    expected= "Send a reminder! He got 10 days left"
    self.assertEqual(expected,message)

  def testForPlentyDays(self):
    testDate = 19042021
    contactInfoFunc()
    datetimestr = datetime.strptime("19.07.2021", "%d.%m.%Y")
    message = checkForDue(4, "testControlPlenty", datetimestr, "","Jens Hansen", todayDate=testDate)
    expected = "Nothing will be done"
    self.assertEqual(expected, message)

  def testForNegativeOneDay(self):
    testDate = 19042021
    contactInfoFunc()
    datetimestr = datetime.strptime("18.04.2021", "%d.%m.%Y")
    message = checkForDue(5, "testControlLateOne", datetimestr, "","Jens Hansen", todayDate=testDate)
    expected = "You are late! Please finish this control before end of date"
    self.assertEqual(expected, message)

  def testForNegativeTwoDays(self):
    testDate = 19042021
    contactInfoFunc()
    datetimestr = datetime.strptime("17.04.2021", "%d.%m.%Y")
    message = checkForDue(6, "testControlLateTwo", datetimestr, "","Jens Hansen", todayDate=testDate)
    expected = "You are late!"

    self.assertEqual(expected, message)

  def testForFailedControl(self):
    testDate = 19042021
    contactInfoFunc()
    datetimestr = datetime.strptime("16.04.2021", "%d.%m.%Y")
    message = checkForDue(7, "testControlFailed", datetimestr, "","Jens Hansen", todayDate=testDate)
    expected= 'this control has not been finished in time or has been incorrectly made.'

    self.assertEqual(expected, message)

  def testForMissingResponsibility(self):
    testDate=19042021
    contactInfoFunc()
    datetimestr = datetime.strptime("29.04.2021", "%d.%m.%Y")
    message =checkForDue(1,2,datetimestr,"","Jens Tester",todayDate=testDate)
    expected= "Missing Contact Information"
    self.assertEqual(expected,message)

  def tearDown(self):
    pass

class contactInfoFuncUnittest(unittest.TestCase):
  def setUp(self):
    testDate = 19042021
    contactInfoFunc()
    datetimestr = datetime.strptime("29.04.2021", "%d.%m.%Y")
    Ctrls.ctrlsList.clear()
    Ctrls(
      "1",
      'Verify Screening processes',
      datetimestr,
      "X",
      'Knud Bentsen')
    Ctrls(
      "2",
      'Verify terms and conditions',
      datetimestr,
      "",
      'Knud Bentsen')
    Ctrls(
      "3",
      'Verify Screening processes',
      datetimestr,
      "",
      'Knud Bentsen')

  def testClassmaker(self):
    message=mailingList
    expected=[[7, 'testControlFailed', datetime.date(datetime.strptime("16.04.2021", "%d.%m.%Y")), '', 'jensh6247@gmail.com', 'this control has not been finished in time or has been incorrectly made.'],
              [2, 'testControlFive', datetime.date(datetime.strptime("24.04.2021", "%d.%m.%Y")), '', 'jensh6247@gmail.com', 'Send a reminder!'],
              [5, 'testControlLateOne', datetime.date(datetime.strptime("18.04.2021", "%d.%m.%Y")), '', 'jensh6247@gmail.com', 'You are late! Please finish this control before end of date'],
              [6, 'testControlLateTwo', datetime.date(datetime.strptime("17.04.2021", "%d.%m.%Y")), '', 'jensh6247@gmail.com', 'You are late!'],
              [3, 'testControlTen', datetime.date(datetime.strptime("29.04.2021", "%d.%m.%Y")), '', 'jensh6247@gmail.com', 'Send a reminder! He got 10 days left'],
              [1, 'testControlZero', datetime.date(datetime.strptime("19.04.2021", "%d.%m.%Y")), '', 'jensh6247@gmail.com', 'Send The email!']]

    self.assertEqual(expected, message)

  def testCtrlsClassUnittest(self):
    message=""
    for item in Ctrls.ctrlsList:
      message+=str(item.number)+" "+item.control+" "+str(item.due)+ " "+item.verification+ " "+ item.responsible +"\n"

    expected="1 Verify Screening processes 2021-04-29 00:00:00 X Knud Bentsen"+"\n"+\
             "2 Verify terms and conditions 2021-04-29 00:00:00  Knud Bentsen"+"\n"+\
             "3 Verify Screening processes 2021-04-29 00:00:00  Knud Bentsen"+"\n"
    self.assertEqual(expected, message)

if __name__ == '__main__':
  unittest.main()
