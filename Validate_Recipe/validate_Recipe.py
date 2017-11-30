from common import *
from openpyxl.utils import get_column_letter
import openpyxl
from validateR import ValidateError, find_in_dict, find_multiple_in_dict

def writeHeaderReport(ws, status, data, errorMsg, debug=None, colName=None, isQL=None):
    new_row = []
    new_row.append(status)
    if len(data) != 4:
        raise ValueError("[Recipe] Data size is not correct")
    new_row += data
    new_row.append(errorMsg)
    new_row.append(debug)
    new_row.append(colName)
    new_row.append(isQL)
    insert_new_row(ws, new_row) 

def get_decimal_places(ws, row):
    MIC_HEADER = 3
    col = findColumnLetterByColNameAndStartRow(ws, "DEC_PLACES", MIC_HEADER)
    dec = ws[col + str(row)].value
    if dec is None:
        return 0
    return dec

def validate(wb, dataWb):
    ## CONFIG HERE NA N'Narm ##
    DATA_TAB_NAME = "QM_Recipe" # sheet name to find data
    DATA_ROW_COUNT = 2 # how many row to skip in header
    DATA_FIELD_ROW = 1
    DATA_HEADER_ROW = 2 # what row to find by field
    ROW_START = 2 # row to start writing data
    IS_FREEZE = True # wanna freeze header ?

    active_ws = wb.get_sheet_by_name("QM_Recipe")
    if (IS_FREEZE):
        active_ws.freeze_panes = "A"+ str(ROW_START)
    data_ws = dataWb.get_sheet_by_name(DATA_TAB_NAME)
    n_of_data = data_ws.max_row - DATA_ROW_COUNT
    
    # CHECK Addtional Condition 1-2
    PLNNR_col = findColumnLetterByColNameAndStartRow(data_ws, "PLNNR", DATA_HEADER_ROW)
    PLNAL_col = findColumnLetterByColNameAndStartRow(data_ws, "PLNAL", DATA_HEADER_ROW)
    VORNR_col = findColumnLetterByColNameAndStartRow(data_ws, "VORNR", DATA_HEADER_ROW)
    MERKNR_col = findColumnLetterByColNameAndStartRow(data_ws, "MERKNR", DATA_HEADER_ROW)
    
    for i in range(DATA_ROW_COUNT+1, n_of_data + DATA_ROW_COUNT+1):
        PLNNR = data_ws[PLNNR_col + str(i)].value
        PLNAL = data_ws[PLNAL_col + str(i)].value
        VORNR = data_ws[VORNR_col + str(i)].value
        MERKNR = data_ws[MERKNR_col + str(i)].value
        d = dict()
        d["PLNNR"] = PLNNR
        d["PLNAL"] = PLNAL
        d["VORNR"] = VORNR
        d["MERKNR"] = MERKNR
        match_cond_1 = find_by_keys(data_ws, DATA_HEADER_ROW, DATA_ROW_COUNT, d)
        # print("Cond1", match_cond_1)

        data = [PLNNR, PLNAL, VORNR, MERKNR]
        if len(match_cond_1) > 1:
            writeHeaderReport(active_ws, "ERROR", data, ValidateError.DUPLICATE_KEY[1], "N="+str(len(match_cond_1)))
    
    print("Fin Additional Condition")
	
    # Check By Field
    key = ["PLNNR", "PLNAL", "VORNR", "MERKNR"]
    for i in range(DATA_ROW_COUNT+1, n_of_data + DATA_ROW_COUNT+1):
        report_data = [
                data_ws[PLNNR_col+str(i)].value, 
                data_ws[PLNAL_col+str(i)].value,
                data_ws[VORNR_col+str(i)].value,
                data_ws[MERKNR_col+str(i)].value,
            ]
        key_value = [
                data_ws[PLNNR_col+str(i)].value, 
                data_ws[PLNAL_col+str(i)].value,
                data_ws[VORNR_col+str(i)].value,
                data_ws[MERKNR_col+str(i)].value,
            ]
        key_data_dict = dict(itertools.izip(key,key_value))

        QPMK_WERKS_col = findColumnLetterByColNameAndStartRow(data_ws, "QPMK_WERKS", DATA_HEADER_ROW)
        VERKMERKM_col = findColumnLetterByColNameAndStartRow(data_ws, "VERWMERKM", DATA_HEADER_ROW)
        QPMK_WERKS = data_ws[QPMK_WERKS_col + str(i)].value
        VERWMERKM = data_ws[VERKMERKM_col + str(i)].value
        if QPMK_WERKS is None:
            writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NOT_NULL[1].format("Plant - MIC"), i)
        if VERWMERKM is None:
            writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NOT_NULL[1].format("Master Inspection Characteristic"), i)

        x_dict = dict()
        x_dict[2] = QPMK_WERKS
        x_dict[3] = VERWMERKM
        found_QL = find_multiple_in_dict("06-MICQL", x_dict)
        found_QN = find_multiple_in_dict("06-MICQN", x_dict)
        if len(found_QL) == 0 and len(found_QN) != 0:
            isQL = False
        elif len(found_QL) != 0 and len(found_QN) == 0:
            isQL = True
        elif len(found_QL) == 0 and len(found_QN) == 0:
            writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.UNDEFINED[1].format("MIC Doesn't exist"), i)
            continue
        else:
            writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.UNDEFINED[1].format("Found in both MICQL, MICQN"), i)
            continue

        for j in range(1, data_ws.max_column +1):
            field_descr = data_ws.cell(row=DATA_FIELD_ROW, column=j).value
            real_data = data_ws.cell(row=i, column=j).value
            if isinstance(real_data, int) or isinstance(real_data, long) or isinstance(real_data, float):
                data = str(real_data)
            else:
                data = real_data
            
            if data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "PLNNR": #Group           
                if isNull(data):
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NOT_NULL[1].format(field_descr), i, data_ws.cell(row=DATA_HEADER_ROW, column=j).value, isQL)
                elif not isNull(data) and len(data) > 15:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.LENGTH[1].format(field_descr), i, data_ws.cell(row=DATA_HEADER_ROW, column=j).value, isQL)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "PLNAL": #Group Counter
                if isNull(data):
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NOT_NULL[1].format(field_descr), i, data_ws.cell(row=DATA_HEADER_ROW, column=j).value, isQL)
                elif not isNumeric(data):
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.VALUE_TYPE[1].format(field_descr), i, data_ws.cell(row=DATA_HEADER_ROW, column=j).value, isQL)
                if not isNull(data) and len(data) > 2:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.LENGTH[1].format(field_descr), i, data_ws.cell(row=DATA_HEADER_ROW, column=j).value, isQL)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "WERKS": #Plant
            	if isNull(data):
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NOT_NULL[1].format(field_descr), i, data_ws.cell(row=DATA_HEADER_ROW, column=j).value, isQL)
                elif not isNull(data) and len(data) > 4:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.LENGTH[1].format(field_descr), i, data_ws.cell(row=DATA_HEADER_ROW, column=j).value, isQL)
                elif not isNull(data) and find_in_dict("03-Plant",1, real_data) is None:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.FIXED_VALUE_EMPTY[1].format(field_descr), i, data_ws.cell(row=DATA_HEADER_ROW, column=j).value, isQL)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "DATUV": #Valid date
                if isNull(data):
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NOT_NULL[1].format(field_descr), i, data_ws.cell(row=DATA_HEADER_ROW, column=j).value, isQL)
                elif not isNull(data) and data != "01012017":
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.FIXED_VALUE[1].format(field_descr, "X"), i, data_ws.cell(row=DATA_HEADER_ROW, column=j).value, isQL)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "SLWBEZ": #Inspection Point
                if isNull(data):
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NOT_NULL[1].format(field_descr), i, data_ws.cell(row=DATA_HEADER_ROW, column=j).value, isQL)
                elif not isNull(data) and data != "FH1":
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.FIXED_VALUE[1].format(field_descr, "X"), i, data_ws.cell(row=DATA_HEADER_ROW, column=j).value, isQL)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "PARL": #Partial-lot assign.
                if not isNull(data):
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NULL[1].format(field_descr), i, data_ws.cell(row=DATA_HEADER_ROW, column=j).value, isQL)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "QPRZIEHVER": #Sample-drawing proc.
                if not isNull(data):
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NULL[1].format(field_descr), i, data_ws.cell(row=DATA_HEADER_ROW, column=j).value, isQL)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "DMYM": #Dynamic mod. level
                if not isNull(data):
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NULL[1].format(field_descr), i, data_ws.cell(row=DATA_HEADER_ROW, column=j).value, isQL)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "QDYNREGEL": #Modification rule
                if not isNull(data):
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NULL[1].format(field_descr), i, data_ws.cell(row=DATA_HEADER_ROW, column=j).value, isQL)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "VORNR":
                if isNull(data):
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NOT_NULL[1].format(field_descr), i, data_ws.cell(row=DATA_HEADER_ROW, column=j).value, isQL)
                elif not isNull(data) and len(data) != 4:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.UNDEFINED[1].format("Operation should have 4 digits"), i, data_ws.cell(row=DATA_HEADER_ROW, column=j).value, isQL)
                elif not isNumOnly(data):
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.VALUE_TYPE[1].format(field_descr), i, data_ws.cell(row=DATA_HEADER_ROW, column=j).value, isQL)