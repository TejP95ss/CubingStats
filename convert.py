import pandas as pd
import os
import glob
import ast
from datetime import datetime
from openpyxl.styles import Alignment
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
from openpyxl.styles import Border, Side
import openpyxl

downloads_path = r"C:\Users\tejpa\Downloads"
file_pattern = os.path.join(downloads_path, "cstimer*.txt")
matching_files = glob.glob(file_pattern)

if not matching_files:
    raise FileNotFoundError("No file starting with 'cstimer' and ending with '.txt' found in Downloads.")
    
cstimer_file = matching_files[0]
with open(cstimer_file, 'r') as file:
    raw_data = file.read()

session2_index = raw_data.find("session2")
if session2_index != -1:
    raw_data = raw_data[:session2_index - 3]
    raw_data = raw_data + ']}'

data = ast.literal_eval(raw_data)

solves_data = []
i = 1
pllDict = {}
for solve in data["session1"]:
    time_ms = solve[0][1]
    pll = solve[2]
    if pll not in pllDict and len(pll) > 0:
        pllDict[pll] = 1
    elif pll in pllDict:
        pllDict[pll] += 1
    time_sec = int(time_ms) / 1000.0
    solves_data.append([i, time_sec, pll])
    i += 1


print(pllDict)
print("Distinct PLLs: ", len(pllDict))
sum = 0
for key in pllDict:
    sum += pllDict[key]
print("Sum PLLs: ", sum)


excel_path = r"C:\Users\tejpa\OneDrive\tej\OneDrive\Cubing.xlsx"
sheet_name = "MAIN(Start-08-31-2025-Sun)"

wb = openpyxl.load_workbook(excel_path)
ws = wb[sheet_name]
alignment = Alignment(horizontal="center", vertical="center")


row = 2
while ws.cell(row=row, column=1).value is not None:
    row += 1

latest_solve_num = solves_data[-1][0]
latest_solve_num_excel = ws.cell(row=row-1, column=1).value
number_of_values_to_add = latest_solve_num - latest_solve_num_excel
if(number_of_values_to_add <= 0):
    raise Exception("No new values to add!")

#adds the last_n_values to the ws
def updateExcelFile(wb, ws, last_n_values, row):
    df = pd.DataFrame(solves_data[-last_n_values:], columns=["Solve Number", 'Time', "PLL"])
    for r in df.itertuples(index=False):
        ws.cell(row=row, column=1).value = r[0]
        ws.cell(row=row, column=2).value = r[1]
        ws.cell(row=row, column=3).value = r[2]
        ws.cell(row=row, column=1).alignment = alignment
        ws.cell(row=row, column=2).alignment = alignment
        ws.cell(row=row, column=3).alignment = alignment
        row += 1
    wb.save(excel_path)
    print(f"Data appended to {excel_path} in sheet '{sheet_name}'.")

#returns the next cell that is empty so that data can be added there
def findNextCell(ws, top_left, top_right):
    col_inc = findColInc(ws, top_left)
    row_inc = findRowInc(ws, top_left)

    row, col_number = coordinate_to_tuple(top_left)
    while ws[f"{get_column_letter(col_number)}{row}"].value is not None:
        row += row_inc
    
    row -= row_inc
    while ws[f"{get_column_letter(col_number)}{row}"].value is not None:
        col_number += col_inc

    _, tr_col_number = coordinate_to_tuple(top_right)
    if(tr_col_number < col_number):
        _, tl_col_number = coordinate_to_tuple(top_left)
        col_number = tl_col_number
        row += row_inc

    return f"{get_column_letter(col_number)}{row}"

#find the column increment
def findColInc(ws, top_left):
    row, original_col_number = coordinate_to_tuple(top_left)
    col_number = original_col_number + 1
    while(True):
        cv = ws[f"{get_column_letter(col_number)}{row}"].value
        if(cv is not None and isinstance(cv, str) and cv.lower().endswith("day")):
            break
        col_number += 1
    return col_number - original_col_number
    
#find the row increment
def findRowInc(ws, top_left):
    original_row, col_number = coordinate_to_tuple(top_left)
    row_number = original_row + 1
    while(True):
        cv = ws[f"{get_column_letter(col_number)}{row_number}"].value
        if(cv is not None and isinstance(cv, str) and cv.lower().endswith("day")):
            break
        row_number += 1
    return row_number - original_row
   
#easier way to change the value and also format the size, font, and alignment
def changeCellValue(cell, value, size, bold):
    cell.value = value
    cell.font = Font(name="Aptos Narrow", size=size, bold=bold)
    cell.alignment = alignment

#adds the various averages, medians, counts for the given batch of data
def addData(latest_solve_num_excel, number_of_values_to_add, dateCell):
    latest_solve_num_excel += 2
    last_cell = latest_solve_num_excel + number_of_values_to_add - 1
    range_of_values = f"B{latest_solve_num_excel}:B{last_cell}"

    row, start_col = coordinate_to_tuple(dateCell)
    end_col = start_col + 2
    start_col_letter = get_column_letter(start_col)
    end_col_letter = get_column_letter(end_col)

    top_row_range = f"{start_col_letter}{row}:{end_col_letter}{row}"
    second_row_range = f"{start_col_letter}{row + 1}:{end_col_letter}{row + 1}"

    ws.merge_cells(top_row_range)
    ws.merge_cells(second_row_range)

    count_cell = returnRelativeCell(dateCell, 11, 1)

    add_thick_border(ws, f"{dateCell}:{returnRelativeCell(dateCell, 14, 2)}")
    changeCellValue(ws[dateCell], datetime.today().strftime("%m/%d/%Y/%A")  , 12, True)
    changeCellValue(ws[returnRelativeCell(dateCell, 1, 0)], "OVERALL STATS", 12, True)
    for i in range(6):
        value = 15 + i
        changeCellValue(ws[returnRelativeCell(dateCell, 3 + i, 0)], f"Sub {value}", 12, True)
    changeCellValue(ws[returnRelativeCell(dateCell, 9, 0)], "Above 25", 12, True)
    for i in range(6):
        value = 15 + i
        changeCellValue(ws[returnRelativeCell(dateCell, 3 + i, 1)], f'=COUNTIF({range_of_values}, "<={value}")', 11, False)
    changeCellValue(ws[returnRelativeCell(dateCell, 9, 1)], f"=COUNTIF({range_of_values}, \">=25\")", 11, False)
    for i in range(7):
        changeCellValue(ws[returnRelativeCell(dateCell, 3 + i, 2)], f"=({returnRelativeCell(dateCell, 3 + i, 1)}/{count_cell}) * 100", 11, False)

    changeCellValue(ws[returnRelativeCell(dateCell, 2, 1)], "Solves", 12, True)
    changeCellValue(ws[returnRelativeCell(dateCell, 2, 2)], "Percent", 12, True)

    for i in range(4):
        titles = ["Total", "Hours Solving", "Avg Time", "Median Time"]
        formulas = [f"=COUNT({range_of_values})", f"=SUM({range_of_values}) / 60 / 60", f"=SUM({range_of_values})/{count_cell}", f"=MEDIAN({range_of_values})"]
        changeCellValue(ws[returnRelativeCell(dateCell, 11 + i, 0)], titles[i], 12, True)
        changeCellValue(ws[returnRelativeCell(dateCell, 11 + i, 1)], formulas[i], 11, False)

    ws[returnRelativeCell(dateCell, 12, 1)].number_format = '0.000'
    ws[returnRelativeCell(dateCell, 13, 1)].number_format = '0.000'
    for i in range(7):
        ws[returnRelativeCell(dateCell, 3 + i, 2)].number_format = '0.00'

    wb.save(excel_path)
    print(f"Data added to {excel_path} in sheet '{sheet_name}'.")
    return 0

#adds a thick ourside border around a block of cells
def add_thick_border(ws, cell_range):
    thick = Side(border_style="thick", color="000000")
    
    start_cell, end_cell = cell_range.split(":")
    _, start_row = ''.join(filter(str.isalpha, start_cell)), int(''.join(filter(str.isdigit, start_cell)))
    _, end_row = ''.join(filter(str.isalpha, end_cell)), int(''.join(filter(str.isdigit, end_cell)))
    
    _, start_col = coordinate_to_tuple(start_cell)
    _, end_col = coordinate_to_tuple(end_cell)

    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            cell = ws[f"{get_column_letter(col)}{row}"]

            border_sides = {
                "top": thick if row == start_row else None,
                "bottom": thick if row == end_row else None,
                "left": thick if col == start_col else None,
                "right": thick if col == end_col else None,
            }

            cell.border = Border(
                top=border_sides["top"] or cell.border.top,
                bottom=border_sides["bottom"] or cell.border.bottom,
                left=border_sides["left"] or cell.border.left,
                right=border_sides["right"] or cell.border.right
            )

def returnRelativeCell(top_left, row_offset, col_offset):
    row, col = coordinate_to_tuple(top_left)
    row += row_offset
    col += col_offset
    return f"{get_column_letter(col)}{row}"

##running the below two functions will update the excel file. Without it, only basic data like pll, sum plls will show
updateExcelFile(wb, ws, number_of_values_to_add, row)
addData(latest_solve_num_excel, number_of_values_to_add, findNextCell(ws, "AN8", "BD8"))
#venv/scripts/activate
#python convert.py


