from openpyxl import load_workbook
from os import getcwd, walk
from shutil import move
import datetime
from openpyxl.styles import NamedStyle
validatedControls = []


def DateToExcel(day, month, year):

  offset = 693594
  current = datetime.date(year, month, day)
  n = current.toordinal()
  return (n - offset)


def downloadFiles():
  dlPath = getcwd() + '\\' 'Downloaded controls'
  for roots, dirs, files in walk(dlPath):
    for file in files:
      sheet = load_workbook(dlPath + '\\' + file)
      ws = sheet.active
      maxControlRow = len(ws['B'])
      ListofControls = []
      for value in ws.iter_rows(
        min_row=3,
        max_row=maxControlRow,
        min_col=2,
        max_col=10,
        values_only=True):



        if value[0] != (None):
          ListofControls.append(value[3])

      if (validationOfControl(ListofControls) == 100.0):
        if isinstance(value[6], datetime.date) and value[7]!= None:
          validatedControls.append([file,value[6],value[7],value[8]])
        else:
          print("control Went bad")

          print("The Date value \"",value[6],  "\" or responsible value \"",value[7],"\" is not correct")


      sheet.close()

def validationOfControl(Listitem):
  count = 0
  try:
    for item in Listitem:
      if "Yes" == item:
        count += 1
    percentage = count/len(Listitem) * 100
    if (percentage == 100.0):
      print("Control Done")
      return percentage
    else:
      print("Control Failed")
      return percentage
  except ZeroDivisionError:
    print("list is empty. Controller forgot to finish his Control!")



def UpdatingControls():
  emptyString = ""
  sheet = load_workbook(filename="mainControllerDoc\\Kontroller.xlsx")
  ws = sheet.active
  maxCtrlRowKontrol = len(ws['A'])
  print(maxCtrlRowKontrol)

  for item in validatedControls:
    validated = item[0].rsplit('.', 1)[0]
    newDate=str(item[1]).strip()[:-9]
    newDay=int(newDate.split("-")[2])
    newMonth = int(newDate.split("-")[1])
    newYear = int(newDate.split("-")[0])
    newlyDate=DateToExcel(newDay,newMonth,newYear)

    newResponsible=item[2]
    for rows in ws.iter_rows(min_row=0,
                           max_row=maxCtrlRowKontrol,
                           min_col=1,
                           max_col=5, ):
      count = 0
      for cell in rows:
        if cell.value == None:
          print(emptyString.strip()[:-9])
          count += 1

          if (emptyString.strip()[:-9] == validated):
            coordOfInterest = str(rows[3]).split('.')[1][:-1]
            ws[coordOfInterest] = 'X'

            ctrlName=emptyString.strip()[:-20]
            newName=" ".join(ctrlName.split(" ")[1:])
            newCoordA="A"+str(maxCtrlRowKontrol+1)
            newCoordB="B" + str(maxCtrlRowKontrol + 1)
            newCoordC="C" + str(maxCtrlRowKontrol + 1)
            newCoordE="E"+ str(maxCtrlRowKontrol + 1)
            ws[newCoordA] = int(maxCtrlRowKontrol)
            ws[newCoordB]=newName
            ws[newCoordC]=newlyDate
            coordwithDate=ws.cell(maxCtrlRowKontrol+1,3)
            coordwithDate.number_format='DD-MM-YYYY'

            ws[newCoordE]=newResponsible
            sheet.save(r"mainControllerDoc\Kontroller.xlsx")
            move("Downloaded controls\\"+item[0],"Evidence\\"+newName+"\\"+item[0])
            emptyString = ""
          else:
            emptyString = ""

        else:
          if count == 4:
            emptyString = ""
          else:
            emptyString += str(cell.value) + " "
            count += 1




downloadFiles()
UpdatingControls()