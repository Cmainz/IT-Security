from openpyxl import load_workbook
from openpyxl.styles import NamedStyle
from datetime import date

sheet = load_workbook(filename="mainControllerDoc\\mainControls.xlsx")
ws_ctrl = sheet.active
max_control_row = len(ws_ctrl['A'])
all_main_ctrls=set()

production_sheet = load_workbook(filename="mainControllerDoc\\Kontroller.xlsx")
ws_prod_ctrl = production_sheet.active
max_prod_ctrl_row=len(ws_prod_ctrl['A'])
all_prod_ctrls=set()

ctrlDict={}

def date_to_excel(day, month, year):

  offset = 693594
  current = date(year, month, day)
  n = current.toordinal()
  return (n - offset)

def set_ctrl(worksheet, maxRow, final_set):
  for pages in worksheet:
    for row in pages.iter_rows(min_row=2,
                               max_row=maxRow,
                               min_col=2,
                               max_col=5,
                               values_only=True):

      if row[0]== None:
          continue
      else:
        final_set.add(row[0])
        ctrlDict[row[0]]=(row[1],row[2])
  return final_set

def insert_new_ctrl(ctrl, ctrl_date, responsible):
  new_coord_a = "A" + str(max_prod_ctrl_row + 1)
  new_coord_b = "B" + str(max_prod_ctrl_row + 1)
  new_coord_c = "C" + str(max_prod_ctrl_row + 1)
  new_coord_e = "E" + str(max_prod_ctrl_row + 1)


  new_date = str(ctrl_date).strip()[:-9]
  new_day = int(new_date.split("-")[2])
  new_month = int(new_date.split("-")[1])
  new_year = int(new_date.split("-")[0])
  new_date_ctrl=date_to_excel(new_day, new_month, new_year)

  ws_prod_ctrl[new_coord_a] = int(max_prod_ctrl_row)
  ws_prod_ctrl[new_coord_b] = ctrl
  ws_prod_ctrl[new_coord_c] = new_date_ctrl
  ws_prod_ctrl[new_coord_e] = responsible

  coordwithDate = ws_prod_ctrl.cell(max_prod_ctrl_row + 1, 3)
  coordwithDate.number_format = 'DD-MM-YYYY'
  production_sheet.save(r"mainControllerDoc\Kontroller.xlsx")

def check_for_match(a_ctrls, b_ctrls):
  for controls in a_ctrls:
    if controls in b_ctrls:
      continue
    else:
      print("\"" + controls + "\" is missing ")
      print(ctrlDict[controls])
      insert_new_ctrl(controls, ctrlDict[controls][0], ctrlDict[controls][1])



set_ctrl(production_sheet, max_prod_ctrl_row, all_prod_ctrls)
set_ctrl(sheet, max_control_row, all_main_ctrls)

check_for_match(all_main_ctrls, all_prod_ctrls)
