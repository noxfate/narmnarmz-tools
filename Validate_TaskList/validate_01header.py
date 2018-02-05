from common import *
from validateTL import ValidateError, find_in_dict

def writeHeaderReport(ws, status, data, errorMsg, debug=None):
    new_row = []
    new_row.append(status)
    if len(data) != 4:
        raise ValueError("[01-Header] Data size is not correct")
    new_row += data
    new_row.append(errorMsg)
    new_row.append(debug)
    insert_new_row(ws, new_row)    

def get_value_by_row_colname(ws, colname, row):
    MIC_HEADER = 7
    col = findColumnLetterByColNameAndStartRow(ws, colname, MIC_HEADER)
    return ws[col + str(row)].value

def validate(wb, dataWb):
    ## CONFIG HERE NA N'Narm ##
    DATA_TAB_NAME = "Task List Header" # sheet name to find data
    DATA_ROW_COUNT = 7 # how many row to skip in header
    DATA_FIELD_ROW = 1
    DATA_HEADER_ROW = 7 # what row to find by field
    ROW_START = 2 # row to start writing data
    IS_FREEZE = True # wanna freeze header ?

    active_ws = wb.get_sheet_by_name("1. Task List Header")
    if (IS_FREEZE):
        active_ws.freeze_panes = "A"+ str(ROW_START)
    data_ws = dataWb.get_sheet_by_name(DATA_TAB_NAME)
    n_of_data = data_ws.max_row - DATA_ROW_COUNT


    # CHECK Addtional Condition 1-2
    PLNNR_col = findColumnLetterByColNameAndStartRow(data_ws, "PLNNR", DATA_HEADER_ROW)
    PLNAL_col = findColumnLetterByColNameAndStartRow(data_ws, "PLNAL", DATA_HEADER_ROW)
    WERKS_col = findColumnLetterByColNameAndStartRow(data_ws, "WERKS", DATA_HEADER_ROW)
    KTEXT_col = findColumnLetterByColNameAndStartRow(data_ws, "KTEXT", DATA_HEADER_ROW)
    for i in range(DATA_ROW_COUNT+1, n_of_data + DATA_ROW_COUNT+1):
        PLNNR = data_ws[PLNNR_col + str(i)].value
        PLNAL = data_ws[PLNAL_col + str(i)].value
        d = dict()
        d["PLNNR"] = PLNNR
        d["PLNAL"] = PLNAL
        match_cond_1 = find_by_keys(data_ws, DATA_HEADER_ROW, DATA_ROW_COUNT, d)
        # print("Cond1", match_cond_1)

        #Check operation
        header_ws = dataWb.get_sheet_by_name("TaskList Operation")
        d = dict()
        d["PLNNR"] = PLNNR
        d["PLNAL"] = PLNAL
        match_cond_2 = find_by_keys(header_ws, DATA_HEADER_ROW, DATA_ROW_COUNT, d)

        WERKS = data_ws[WERKS_col + str(i)].value
        KTEXT = data_ws[KTEXT_col + str(i)].value
    
        if len(match_cond_1) > 1:
            data = [PLNNR, PLNAL, WERKS, KTEXT]
            writeHeaderReport(active_ws, "ERROR", data, ValidateError.DUPLICATE_KEY[1], "N="+str(len(match_cond_1)))

        if len(match_cond_2) < 1:
            data = [PLNNR, PLNAL, WERKS, KTEXT]
            writeHeaderReport(active_ws, "ERROR", data, ValidateError.UNDEFINED[1].format("Group not mapping with 2. Task List Operation"), "N="+str(len(match_cond_2)))

    
    print("Fin Additional Condition")

    # Check By Field
    key = ["PLNNR", "PLNAL"]
    for i in range(DATA_ROW_COUNT+1, n_of_data + DATA_ROW_COUNT+1):
        for j in range(1, data_ws.max_column +1):
            report_data = [
                data_ws[PLNNR_col+str(i)].value, 
                data_ws[PLNAL_col+str(i)].value,
                data_ws[WERKS_col+str(i)].value,
                data_ws[KTEXT_col+str(i)].value
            ]
            key_value = [
                data_ws[PLNNR_col+str(i)].value, 
                data_ws[PLNAL_col+str(i)].value,
            ]
            key_data_dict = dict(itertools.izip(key,key_value))
            field_descr = data_ws.cell(row=DATA_FIELD_ROW, column=j).value

            real_data = data_ws.cell(row=i, column=j).value
            #print(real_data, type(real_data))
            if isinstance(real_data, int) or isinstance(real_data, long) or isinstance(real_data, float):
                data = str(real_data)
            else:
                data = real_data

            PLNNR = get_value_by_row_colname(data_ws, "PLNNR", i) #Group

            if data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "PLNNR":                
                if data is None:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NOT_NULL[1].format(field_descr), i)
                if data is not None and len(data) > 15:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.LENGTH[1].format(field_descr), i)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "DATUV":
                if data != "01012017":
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.FIXED_VALUE[1].format(field_descr, "01012017"), i)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "PLNAL":
                if data is None:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NOT_NULL[1].format(field_descr), i)
                if not isNumeric(data):
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.VALUE_TYPE[1].format(field_descr), i)
                if data is not None and len(data) > 2:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.LENGTH[1].format(field_descr), i)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "KTEXT":
                if data is None:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NOT_NULL[1].format(field_descr), i)
                elif data is not None and len(data) > 40:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.LENGTH[1].format(field_descr), i)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "STRAT":
                if data is not None:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NULL[1].format(field_descr), i)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "WERKS":
                if data is None:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NOT_NULL[1].format(field_descr), i)
                elif data != get_value_by_row_colname(data_ws, "WERKS2", i):
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.UNDEFINED[1].format("Planning Plant not match with Plant for Work Center"), i)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "ARBID":
                if data is None:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NOT_NULL[1].format(field_descr), i)
                elif data is not None and len(data) > 8:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.LENGTH[1].format(field_descr), i)
                elif data is not None and find_in_dict("04-Work Center", 3, real_data.upper()) is None:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.FIXED_VALUE_EMPTY[1].format(field_descr), i)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "WERKS2":
                if data is None:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NOT_NULL[1].format(field_descr), i)
                elif data is not None and len(data) > 4:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.LENGTH[1].format(field_descr), i)
                elif data is not None and find_in_dict("04-Work Center", 2, real_data) is None:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.FIXED_VALUE_EMPTY[1].format(field_descr), i)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "VERWE":
                if data != "4":
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.FIXED_VALUE[1].format(field_descr, "4"), i)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "VAGRP":
                if data is None:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NOT_NULL[1].format(field_descr), i)
                elif data is not None and len(data) > 3:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.LENGTH[1].format(field_descr), i)
                elif data != "Q01" and str(PLNNR)[4:6] != "EC":
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.UNDEFINED[1].format("Planner Group Conflit with Group structure"), i)
                elif data is not None and str(PLNNR)[4:6] == "EC":
                    if data is not None and find_in_dict("01-Planner Group", 1, real_data) is None:
                        writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.UNDEFINED[1].format("Incorrect Planner Group"), i)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "STATU":
                if data != "4":
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.FIXED_VALUE[1].format(field_descr, "4"), i)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "ANLZU":
                if str(data) != "0" and str(data) != "1":
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.FIXED_VALUE[1].format(field_descr, "0 or 1"), i)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "SLWBEZ":
                if data is None:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NOT_NULL[1].format(field_descr), i)
                elif str(PLNNR)[5] == "C" and data != "FHC":
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.UNDEFINED[1].format("Inspection Point Conflit with Group structure"), i)
                elif str(PLNNR)[5] == "V" and data != "FHD":
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.UNDEFINED[1].format("Inspection Point Conflit with Group structure"), i)
                elif str(PLNNR)[5] == "D" and data != "FHD":
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.UNDEFINED[1].format("Inspection Point Conflit with Group structure"), i)
                #elif data is not None and find_in_dict("05-Insp point", 1, real_data) is None:
                    #writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.FIXED_VALUE_EMPTY[1].format(field_descr), i)