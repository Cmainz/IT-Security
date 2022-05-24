"""
This file creates a new control from mainControls.xlsx file.
For the control to function correctly a Template also needs to be written for each new control
Please review "Verify Screening processes.xlsx" for hints of creating a new template

"""

from openpyxl import load_workbook
from openpyxl.styles import NamedStyle
from datetime import date

sheet = load_workbook(filename="mainControllerDoc\\mainControls.xlsx")
ws_ctrl = sheet.active
max_control_row = len(ws_ctrl['A'])
all_main_ctrls=set()

prod_controller_file="mainControllerDoc\\Kontroller.xlsx"
production_sheet = load_workbook(filename=prod_controller_file)
ws_prod_ctrl = production_sheet.active
max_prod_ctrl_row=len(ws_prod_ctrl['A'])
all_prod_ctrls=set()

ctrlDict={}
input_list_excel=[]
def date_to_excel(day, month, year):
  """ Takes a date and make it readable for excel"""
  offset = 693594
  current = date(year, month, day)
  n = current.toordinal()
  return (n - offset)

def set_ctrl(worksheet, max_row, final_set):
  """ Creates a set from an excel sheet"""

  for pages in worksheet:
    for row in pages.iter_rows(min_row=2,
                               max_row=max_row,
                               min_col=2,
                               max_col=5,
                               values_only=True):

      if row[0]== None:
          continue
      else:
        final_set.add(row[0])
        ctrlDict[row[0]]=(row[1],row[2])
  return final_set

def insert_new_ctrl(chosen_ctrls):
  """ Inserts the missing ctrls into production """
  for i in range(len(chosen_ctrls)):
    if (len(chosen_ctrls)) ==0:
      i=1
    input_coord = str(max_prod_ctrl_row + i)
    new_name = chosen_ctrls[i - 1][0]
    new_ctrl_date = chosen_ctrls[i - 1][1]
    new_responsible = chosen_ctrls[i - 1][2]
    new_coord_a = "A" + input_coord
    new_coord_b = "B" + input_coord
    new_coord_c = "C" + input_coord
    new_coord_e = "E" + input_coord
    ws_prod_ctrl[new_coord_a] = int(input_coord) - 1
    ws_prod_ctrl[new_coord_b] = new_name
    ws_prod_ctrl[new_coord_c] = new_ctrl_date
    ws_prod_ctrl[new_coord_e] = new_responsible
    coord_with_date = ws_prod_ctrl.cell(int(input_coord) + 1, 3)
    coord_with_date.number_format = 'DD-MM-YYYY'

    production_sheet.save(prod_controller_file)



def check_for_match(a_ctrls, b_ctrls):
  """ Function that verifies that controls are in production and if not creates one"""
  for controls in a_ctrls:
    if controls in b_ctrls:
      continue
    else:
      print("\"" + controls + "\" is missing ")
      print(ctrlDict[controls])
      input_list_excel.append((controls, ctrlDict[controls][0], ctrlDict[controls][1]))



set_ctrl(production_sheet, max_prod_ctrl_row, all_prod_ctrls)
set_ctrl(sheet, max_control_row, all_main_ctrls)

check_for_match(all_main_ctrls, all_prod_ctrls)
insert_new_ctrl(input_list_excel)