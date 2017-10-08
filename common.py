import Tkinter
import tkFileDialog
import tkMessageBox
import os
import openpyxl
import itertools
from openpyxl.utils import get_column_letter, column_index_from_string
from NarmError import *

cur_path = os.path.dirname(__file__)
new_unit_path = os.path.join(cur_path, 'resource','Dict', 'unit.XLSX')

def openDialog():
    root = Tkinter.Tk()
    root.withdraw()
    fileName = tkFileDialog.askopenfilename()
    print(fileName)
    return fileName

def openExcelFile(filePath):
    os.chdir(os.path.dirname(filePath))
    wb = openpyxl.load_workbook(os.path.basename(filePath))
    return wb

unit_wb = openExcelFile(new_unit_path)

def findColumnLetterByColNameAndStartRow(worksheet, value, rowNumber):
    for i in range(1, worksheet.max_column+1):
        text = worksheet[get_column_letter(i)+str(rowNumber)].value
        if (text == value):
            return get_column_letter(i)

def findCellInColumnByValue(worksheet, col, value, headerRow):
    if isinstance(col, str):
        col_n = findColumnLetterByColNameAndStartRow(worksheet, col, headerRow)
    elif isinstance(col, int):
        col_n = col
    else:
        col_n = 1

    if value is None:
        return None
    for i in range(1, worksheet.max_row+1):
        cell_val = worksheet.cell(row=i, column=col_n).value
        # print(cell_val, type(cell_val)," : ",value, type(value))
        if (cell_val is not None) and (cell_val == value):
            return worksheet.cell(row=i, column=col_n)
    return None

def findCellListInColumnByValue(worksheet, col, value, headerRow):
    col_letter = findColumnLetterByColNameAndStartRow(worksheet, col, headerRow)
    if value is None:
        return None
    result = set()
    rowN = worksheet.max_row
    for i in range(1, rowN+1):
        cell_val = worksheet[col_letter+str(i)].value
        if (cell_val is not None) and (cell_val == value):
            result.add(worksheet[col_letter+str(i)])
    return result

def transformUnit(input):
    global unit_wb
    ws = unit_wb.get_sheet_by_name("Unit")
    technical_cell = findCellInColumnByValue(ws, "Technical", input, 1)
    if technical_cell is None:
        # raise UnitConversionError("UnitConversion Error: ", ws, "Technical", input)
        return None
    commercial_col_letter = findColumnLetterByColNameAndStartRow(ws, "Commercial", 1)
    return ws[commercial_col_letter+str(technical_cell.row)]

def transformUnitSampling(input):
    global unit_wb
    ws = unit_wb.get_sheet_by_name("Sampling")
    old_cell = findCellInColumnByValue(ws, "Old", input, 1)
    if old_cell is None:
        return None
    new_cell_letter = findColumnLetterByColNameAndStartRow(ws, "New", 1)
    return ws[new_cell_letter+str(old_cell.row)]

def insert_new_row(ws, row_data):
    n = ws.max_row
    print(n, row_data)
    new_row = n+1
    for i in range(1, len(row_data)+1):
        ws[get_column_letter(i)+str(new_row)].value = row_data[i-1]

def isNumeric(input):
    if input is None:
        return False
    return input.isdigit()

def isNumOnly(input):
    if input is None:
        return False
    data = str(input).strip()
    for i in data:
        if not i.isdigit():
            return False
    return True
    
def find_by_keys(ws, headerRow, dataRowStart, keyDict):
    DATA_ROW_COUNT = dataRowStart # how many row to skip in header
    DATA_HEADER_ROW = headerRow # what row to find by field

    found = []
    for k in keyDict.keys():
        cells = findCellListInColumnByValue(ws, k, keyDict[k], DATA_HEADER_ROW)        
        rows = []
        if cells is not None:
            cells_list = list(cells)
            for i in range(len(cells_list)):
                rows.append(cells_list[i].row)
        found.append(set(rows))
    result = set.intersection(*found)
    return result

def calAfterPoint(x):
    if not '.' in x:
        return 0
    return len(x) - x.index('.') - 1

def checkDecimalPlace(num, data):
    """ checkDecimalPlace(int Number of decimal Place, string data) """
    if data == "":
        return True
    else:
        d = calAfterPoint(data)
        num = int(num)
        if abs(num) == abs(d):
            return True
        else:
            return False
