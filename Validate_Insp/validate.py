import Tkinter
import tkFileDialog
import os
import openpyxl
import itertools
import easygui
import time
from openpyxl.utils import get_column_letter
from common import *
import sys, traceback

def run():
    filePath = openDialog()
    # wb = openExcelFile(fileName)
    # filePath = "D:/project/narmnarmz-tools/resource/TMP_QM12_Inspection Plan.xlsx"
    start_time = time.time()
    try:
        file_structure = configFileStructure()
        data = openExcelFile(filePath)        
        output_filename = composeFileName(filePath)
        newValidateInspExcel(file_structure, data, output_filename)
        print("--- %s seconds ---" % (time.time() - start_time))
        easygui.msgbox("Your output is "+output_filename+", which is in the same directory that your selected file. \n\nGood luck, have fun!!\n\nExecutime (s): "+str((time.time() - start_time)), title="Success!")
    except TypeError:
        err = traceback.format_exc()
        easygui.msgbox("TypeError: Maybe this happen because the program can't find field in Excel\n\n"+str(err))
    except UnitConversionError as dic:
        easygui.msgbox("Value Not Found: cannot find this following data in 02_Val_Dictionary.xlsx\n\n"+str(dic))
    except:
        err = traceback.format_exc()
        # print(err)
        easygui.msgbox("Unexpected Error: "+str(err))


def composeFileName(fileFullPath):
    return "ERR_"+os.path.basename(fileFullPath)

def newValidateInspExcel(structure, datamodelwb, fileName):
    wb = openpyxl.Workbook()
    old_sheet_list = wb.get_sheet_names()
    for i in structure:
        wb.create_sheet(title=i)
        sheet = wb.get_sheet_by_name(i)
        for j in range(len(structure[i])):
            sheet[get_column_letter(j+1)+'1'] = structure[i][j]
    for i in old_sheet_list:
        wb.remove_sheet(wb.get_sheet_by_name(i))    

    print("....Start Building....")
    
    print("Output: ", fileName)
    wb.save(fileName)

def configFileStructure():
    output_sheets = [
        "01 - Header", "02 - Operation", "03 - MIC", "04 - Mat. Assign"
        ] # index (order) DOES MATTER!!
    main_header = []
    header_01 = ["Status", "Group", "Group Counter", "Plant", "Task list description", "Error Message"]
    header_02 = ["Status", "Group", "Group Counter", "Operation/Activity", "Operation short text", "Error Message"]
    header_03 = ["Status", "Group", "Group Counter", "Operation/Activity", "Characteristic number", "Error Message"]
    header_04 = ["Status", "Group", "Group Counter", "Assign Plant", "Material", "Error Message"]

    # append order MUST match sheet names in [output_sheets]
    main_header.append(header_01)
    main_header.append(header_02)
    main_header.append(header_03)
    main_header.append(header_04)
    result = dict(itertools.izip(output_sheets, main_header))
    return result

def buildHeaderTab(wb, dataWb):
    ## CONFIG HERE NA N'Narm ##
    DATA_TAB_NAME = "01 - Header" # sheet name to find data
    DATA_ROW_COUNT = 2 # how many row to skip in header
    DATA_HEADER_ROW = 2 # what row to find by field
    ROW_START = 2 # row to start writing data
    IS_FREEZE = True # wanna freeze header ?

    active_ws = wb.get_sheet_by_name("01 - Header")
    if (IS_FREEZE):
        active_ws.freeze_panes = "A"+ str(ROW_START)
    data_ws = dataWb.get_sheet_by_name(DATA_TAB_NAME)
    n_of_data = data_ws.max_row - DATA_ROW_COUNT
    for i in range(ROW_START, n_of_data + ROW_START): # ROW LOOP in TASK WORKSHEET by n(data)
        for j in range(1, active_ws.max_column+1): #COLUMN LOOP in TASK WORKSHEET by column template
            col = get_column_letter(j)