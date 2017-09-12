import Tkinter
import tkFileDialog
import os
import openpyxl
import itertools
from openpyxl.utils import get_column_letter
from NarmError import *

cur_path = os.path.dirname(__file__)
new_unit_path = os.path.join(cur_path, 'resource', 'unit.XLSX')

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
    col_letter = findColumnLetterByColNameAndStartRow(worksheet, col, headerRow)
    if value is None:
        return ""
    for i in range(headerRow+1, worksheet.max_row+headerRow+1):
        cell_val = worksheet[col_letter+str(i)].value
        if (cell_val is not None) and (cell_val == value):
            return worksheet[col_letter+str(i)].value
    raise UnitConversionError("UnitConversion Error: ", worksheet, col, value)

def transformUnit(input):
    global unit_wb
    ws = unit_wb.get_active_sheet()
    technical_cell = findCellInColumnByValue(ws, "Technical", input, 1)
    commercial_col_letter = findColumnLetterByColNameAndStartRow(ws, "Commercial", 1)
    return ws[commercial_col_letter+str(technical_cell.row)]
