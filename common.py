import Tkinter
import tkFileDialog
import tkMessageBox
import os
import openpyxl
import itertools
from openpyxl.utils import get_column_letter
from NarmError import *

cur_path = os.path.dirname(__file__)
new_unit_path = os.path.join(cur_path, 'resource','Dict', 'unit.XLSX')
new_val_dict_path = os.path.join(cur_path, 'resource','Dict', 'unit.XLSX')

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
val_dict_wb = openExcelFile(new_val_dict_path)

def findColumnLetterByColNameAndStartRow(worksheet, value, rowNumber):
    for i in range(1, worksheet.max_column+1):
        text = worksheet[get_column_letter(i)+str(rowNumber)].value
        if (text == value):
            return get_column_letter(i)

def findCellInColumnByValue(worksheet, col, value, headerRow):
    col_letter = findColumnLetterByColNameAndStartRow(worksheet, col, headerRow)
    if value is None:
        return None
    for i in range(headerRow+1, worksheet.max_row+headerRow+1):
        cell_val = worksheet[col_letter+str(i)].value
        if (cell_val is not None) and (cell_val == value):
            return worksheet[col_letter+str(i)]
    return None

def findCellListInColumnByValue(worksheet, col, value, headerRow):
    col_letter = findColumnLetterByColNameAndStartRow(worksheet, col, headerRow)
    if value is None:
        return None
    result = set()
    for i in range(headerRow+1, worksheet.max_row+headerRow+1):
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