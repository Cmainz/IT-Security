from openpyxl import load_workbook
from openpyxl.styles import NamedStyle
from datetime import date

sheet = load_workbook(filename="mainControllerDoc\\mainControls.xlsx")
wsCtrl = sheet.active
maxControlRow = len(wsCtrl['A'])
allMainCtrls=set()

productionSheet = load_workbook(filename="mainControllerDoc\\Kontroller.xlsx")
wsProdCtrl = productionSheet.active
maxControlProdRow=len(wsProdCtrl['A'])
allProdCtrls=set()

ctrlDict={}

def DateToExcel(day, month, year):

  offset = 693594
  current = date(year, month, day)
  n = current.toordinal()
  return (n - offset)

def setCtrl(worksheet,maxRow,finalSet):
  for pages in worksheet:
    for row in pages.iter_rows(min_row=2,
                               max_row=maxRow,
                               min_col=2,
                               max_col=5,
                               values_only=True):

      if row[0]== None:
          continue
      else:
        finalSet.add(row[0])
        ctrlDict[row[0]]=(row[1],row[2])
  return finalSet

def insertNewCtrl(ctrl,ctrldate,responsible):
  newCoordA = "A" + str(maxControlProdRow + 1)
  newCoordB = "B" + str(maxControlProdRow + 1)
  newCoordC = "C" + str(maxControlProdRow + 1)
  newCoordE = "E" + str(maxControlProdRow + 1)


  newDate = str(ctrldate).strip()[:-9]
  newDay = int(newDate.split("-")[2])
  newMonth = int(newDate.split("-")[1])
  newYear = int(newDate.split("-")[0])
  newlyDate=DateToExcel(newDay,newMonth,newYear)

  wsProdCtrl[newCoordA] = int(maxControlProdRow)
  wsProdCtrl[newCoordB] = ctrl
  wsProdCtrl[newCoordC] = newlyDate
  wsProdCtrl[newCoordE] = responsible

  coordwithDate = wsProdCtrl.cell(maxControlProdRow + 1, 3)
  coordwithDate.number_format = 'DD-MM-YYYY'
  productionSheet.save(r"mainControllerDoc\Kontroller.xlsx")

def checkForMatch(aCtrls,bCtrls):
  for controls in aCtrls:
    if controls in bCtrls:
      continue
    else:
      print("\"" + controls + "\" is missing ")
      print(ctrlDict[controls])
      insertNewCtrl(controls, ctrlDict[controls][0],ctrlDict[controls][1])



setCtrl(productionSheet,maxControlProdRow,allProdCtrls)
setCtrl(sheet,maxControlRow,allMainCtrls)

checkForMatch(allMainCtrls,allProdCtrls)

#ws[newCoordC] = newlyDate
#newlyDate = DateToExcel(newDay, newMonth, newYear)
#coordwithDate = ws.cell(maxCtrlRowKontrol + 1, 3)
#coordwithDate.number_format = 'DD-MM-YYYY'
