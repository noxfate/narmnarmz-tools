from common import *
from validate import ValidateError, find_in_dict

def writeHeaderReport(ws, status, data, errorMsg, debug=None):
    new_row = []
    new_row.append(status)
    if len(data) != 4:
        raise ValueError("[01-Header] Data size is not correct")
    new_row += data
    new_row.append(errorMsg)
    new_row.append(debug)
    insert_new_row(ws, new_row)    

def check_same_MatAssign_by_meinh(dataWb, keyDict, meinh_data):
    mat_ws = dataWb.get_sheet_by_name("04 - Mat. Assign")
    found = find_by_keys(mat_ws, 2, 2, keyDict)
    found = list(found)
    found.sort()
    row = found[0] if len(found) >= 1 else 0
    if row == 0:
        return False
    MEINS_col = findColumnLetterByColNameAndStartRow(mat_ws, "MEINS", 2)
    MEINS = mat_ws[MEINS_col + str(row)].value
    return MEINS == meinh_data


def validate(wb, dataWb):
    ## CONFIG HERE NA N'Narm ##
    DATA_TAB_NAME = "01 - Header" # sheet name to find data
    DATA_ROW_COUNT = 2 # how many row to skip in header
    DATA_FIELD_ROW = 1
    DATA_HEADER_ROW = 2 # what row to find by field
    ROW_START = 2 # row to start writing data
    IS_FREEZE = True # wanna freeze header ?

    active_ws = wb.get_sheet_by_name("01 - Header")
    if (IS_FREEZE):
        active_ws.freeze_panes = "A"+ str(ROW_START)
    data_ws = dataWb.get_sheet_by_name(DATA_TAB_NAME)
    n_of_data = data_ws.max_row - DATA_ROW_COUNT

    # CHECK Addtional Condition 1-2
    PLNNR_col = findColumnLetterByColNameAndStartRow(data_ws, "PLNNR", DATA_HEADER_ROW)
    PLNAL_col = findColumnLetterByColNameAndStartRow(data_ws, "PLNAL", DATA_HEADER_ROW)
    WERKS_col = findColumnLetterByColNameAndStartRow(data_ws, "WERKS", DATA_HEADER_ROW)
    KTEXT_col = findColumnLetterByColNameAndStartRow(data_ws, "KTEXT", DATA_HEADER_ROW)
    VERWE_col = findColumnLetterByColNameAndStartRow(data_ws, "VERWE", DATA_HEADER_ROW)
    for i in range(DATA_ROW_COUNT+1, n_of_data + DATA_ROW_COUNT+1):
        PLNNR = data_ws[PLNNR_col + str(i)].value
        PLNAL = data_ws[PLNAL_col + str(i)].value
        WERKS = data_ws[WERKS_col + str(i)].value
        KTEXT = data_ws[KTEXT_col + str(i)].value

        d = dict()
        d["PLNNR"] = PLNNR
        d["PLNAL"] = PLNAL
        #cond_1 = check_duplicate_key(data_ws, DATA_HEADER_ROW, DATA_ROW_COUNT, d)
        match_cond_1 = find_by_keys(data_ws, DATA_HEADER_ROW, DATA_ROW_COUNT, d)
        # print("Cond1", match_cond_1)
        
        VERWE = data_ws[VERWE_col+str(i)].value
        d = dict()
        d["VERWE"] = VERWE
        d["PLNNR"] = PLNNR
        #cond_2 = check_duplicate_key(data_ws, DATA_HEADER_ROW, DATA_ROW_COUNT, d)
        match_cond_2 = find_by_keys(data_ws, DATA_HEADER_ROW, DATA_ROW_COUNT, d)
        # print("Cond2", match_cond_2)

        
        data = [PLNNR, PLNAL, WERKS, KTEXT]
        if len(match_cond_1) > 1:
            writeHeaderReport(active_ws, "ERROR", data, ValidateError.DUPLICATE_KEY[1], "N="+str(len(match_cond_1)))
        if len(match_cond_2) > 1:
            writeHeaderReport(active_ws, "ERROR", data, ValidateError.UNDEFINED[1].format("Duplicate usage within Group"), "N="+str(len(match_cond_2)))
        #if cond_1:
            #writeHeaderReport(active_ws, "ERROR", data, ValidateError.DUPLICATE_KEY[1], "row="+str(i))
        #if cond_2:
            #writeHeaderReport(active_ws, "ERROR", data, ValidateError.UNDEFINED[1].format("Duplicate usage within Group"), "row="+str(i))

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
                
            if data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "PLNNR":                
                if isNull(data):
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NOT_NULL[1].format(field_descr), i)
                if not isNull(data) and len(data) > 15:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.LENGTH[1].format(field_descr), i)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "PLNAL":
                if isNull(data):
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NOT_NULL[1].format(field_descr), i)
                if not isNumeric(data):
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.VALUE_TYPE[1].format(field_descr), i)
                if not isNull(data) and len(data) > 2:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.LENGTH[1].format(field_descr), i)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "VERWE":
                if isNull(data):
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NOT_NULL[1].format(field_descr), i)
                if not isNull(data) and len(data) > 3:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.LENGTH[1].format(field_descr), i)
                if not isNull(data) and find_in_dict("01-Usage", 1, real_data) is None:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.FIXED_VALUE_EMPTY[1].format(field_descr), i)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "WERKS":
                if isNull(data):
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NOT_NULL[1].format(field_descr), i)
                if not isNull(data) and len(data) > 4:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.LENGTH[1].format(field_descr), i)
                if not isNull(data) and find_in_dict("03-Plant", 1, real_data) is None:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.FIXED_VALUE_EMPTY[1].format(field_descr), i)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "DATUV":
                #if isNull(data):
                    #writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NOT_NULL[1].format(field_descr), i)
                if data != "01012017":
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.FIXED_VALUE[1].format(field_descr, "01012017"), i)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "STATU":
                #if isNull(data):
                    #writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NOT_NULL[1].format(field_descr), i)
                if data != "4":
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.FIXED_VALUE[1].format(field_descr, "4"), i)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "LOSVN":
                #if isNull(data):
                    #writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NOT_NULL[1].format(field_descr), i)
                if data != "1":
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.FIXED_VALUE[1].format(field_descr, "1"), i)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "LOSBS":
                #if isNull(data):
                    #writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NOT_NULL[1].format(field_descr), i)
                if data != "99999999":
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.FIXED_VALUE[1].format(field_descr, "99999999"), i)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "MEINH_H":
                if isNull(data):
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NOT_NULL[1].format(field_descr), i)
                if not isNull(data) and len(data) > 6:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.LENGTH[1].format(field_descr), i)
                if not check_same_MatAssign_by_meinh(dataWb, key_data_dict, real_data):
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.UNDEFINED[1].format("Unit not mapping with material master"), i)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "KTEXT":
                if isNull(data):
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.NOT_NULL[1].format(field_descr), i)
                if not isNull(data) and len(data) > 40:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.LENGTH[1].format(field_descr), i)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "SLWBEZ":
                if isNumeric(data) and not isNull(data):
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.VALUE_TYPE[1].format(field_descr), i)
                if not isNull(data) and len(data) > 3:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.LENGTH[1].format(field_descr), i)
                #if not isNull(data) and data != "FH4":
                    #writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.FIXED_VALUE[1].format(field_descr, "FH4"), i)
                if not isNull(data) and find_in_dict("05-Insp point", 1, real_data) is None:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.FIXED_VALUE_EMPTY[1].format(field_descr), i)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "QPRZIEHVER":
                if not isNumeric(data) and not isNull(data):
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.VALUE_TYPE[1].format(field_descr), i)
                if not isNull(data) and len(data) > 8:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.LENGTH[1].format(field_descr), i)
            elif data_ws.cell(row=DATA_HEADER_ROW, column=j).value == "QDYNREGEL":
                if not isNumeric(data) and not isNull(data):
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.VALUE_TYPE[1].format(field_descr), i)
                if not isNull(data) and len(data) > 3:
                    writeHeaderReport(active_ws, "ERROR", report_data, ValidateError.LENGTH[1].format(field_descr), i)
            # else:
            #     writeHeaderReport(active_ws, "", report_data, "Success")