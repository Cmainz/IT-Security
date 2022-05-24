from openpyxl import load_workbook
from os import getcwd, walk
from shutil import move
import datetime
from openpyxl.styles import NamedStyle

#downloaded sheets
dl_path = getcwd() + '\\' + 'Downloaded controls'



# Main excel with controls
main_controller_file="mainControllerDoc\\Kontroller.xlsx"
main_controls_sheet = load_workbook(filename=main_controller_file)
ws = main_controls_sheet.active
max_row_control_doc = len(ws['A'])


found_files=[]
dl_controls = []
validatedControls = []
input_list_excel=[]

def date_to_excel(day, month, year):

  offset = 693594
  current = datetime.date(year, month, day)
  n = current.toordinal()
  return (n - offset)

def finding_files():
  for roots, dirs, files in walk(dl_path):
    for file in files:
      found_files.append(file)

def check_completion(list_item):
  for file in list_item:
    sheet = load_workbook(dl_path + '\\' + file)
    ws = sheet.active
    max_control_row = len(ws['B'])
    for value in ws.iter_rows(
      min_row=3,
      max_row=max_control_row,
      min_col=2,
      max_col=10,
      values_only=True):

      if value[0] != None:
        dl_controls.append(value[3])

    if (validating_control(dl_controls) == 100.0):
      if isinstance(value[6], datetime.date) and value[7]!= None:
        validatedControls.append([file,value[6],value[7],value[8]])
      else:
        print("control Went bad")
        print("The Date value \"",value[6],  "\" or responsible value \"",value[7],"\" is not correct")


    sheet.close()

def validating_control(list_item):
  count = 0
  try:
    for item in list_item:
      if item.lower() == "yes":
        count += 1
    percentage = count / len(list_item) * 100
    if (percentage == 100.0):
      print("Control Done")
      return percentage
    else:
      print("Control Failed")
      return percentage
  except (ZeroDivisionError, AttributeError):
    print("list is empty. Controller forgot to finish his Control!")
    return 0

def validator(chosen_ctrls):
  for i in range(len(chosen_ctrls)):
    input_coord = str(max_row_control_doc + i)
    new_name = chosen_ctrls[i - 1][0]
    new_ctrl_date = chosen_ctrls[i - 1][1]
    new_responsible = chosen_ctrls[i - 1][2]
    new_coord_a = "A" + input_coord
    new_coord_b = "B" + input_coord
    new_coord_c = "C" + input_coord
    new_coord_e = "E" + input_coord
    ws[new_coord_a] = int(input_coord) - 1
    ws[new_coord_b] = new_name
    ws[new_coord_c] = new_ctrl_date
    coord_with_date = ws.cell(int(input_coord) + 1, 3)
    coord_with_date.number_format = 'DD-MM-YYYY'
    ws[new_coord_e] = new_responsible
    main_controls_sheet.save(main_controller_file)

def update_controls(valid_list):
  empty_string = ""

  for item in valid_list:
    validated = item[0].rsplit('.', 1)[0]
    new_ctrl_date=str(item[1]).strip()[:-9]
    new_date=int(new_ctrl_date.split("-")[2])
    new_month = int(new_ctrl_date.split("-")[1])
    new_year = int(new_ctrl_date.split("-")[0])
    new_ctrl_date=date_to_excel(new_date, new_month, new_year)


    new_responsible=item[2]
    for rows in ws.iter_rows(min_row=0,
                           max_row=max_row_control_doc,
                           min_col=1,
                           max_col=5, ):
      count = 0
      for cell in rows:
        if cell.value == None:
          count += 1

          if (empty_string.strip()[:-9] == validated):
            coord_of_interest = str(rows[3]).split('.')[1][:-1]
            ws[coord_of_interest] = 'X'
            ctrl_name=empty_string.strip()[:-20][2:]
            input_list_excel.append((ctrl_name,new_ctrl_date,new_responsible))
            #move("Downloaded controls\\" + item[0], "Evidence\\"+ctrl_name + "\\" + item[0])

            empty_string = ""
          else:
            empty_string = ""

        else:
          if count == 4:
            empty_string = ""
          else:
            empty_string += str(cell.value) + " "
            count += 1

  validator(input_list_excel)



finding_files()
check_completion(found_files)
update_controls(validatedControls)