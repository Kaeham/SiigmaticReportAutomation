def extract_water_data(ws):
    """
    Extracts water data
    :Inputs:
        - Workbook: excel workbook containing the data
    :Outputs:
        - extractedData: (water measured value, comments about test)
    """
    # Initial positions
    waterRow, commentsRow, col = (3, 5, 23)
    sectionRow = commentsRow + 13
    waterData = None
    loopIdx = 0

    # extract first, second and third test values
    # store these values if and only if the test is a valid test
    # store them in val, if they are not default values or empty etc.
    # if the current val is empty, return the previous values
    # otherwise iterate until an empty value is found
    while waterRow <= 1048576:
        # print(waterRow, modRow, commentsRow)
        waterValue = ws.cell(row=waterRow, column=col).value
        waterComments = _gather_comments(ws, commentsRow, col)
        validTest = _check_section_pass(ws, sectionRow, col) or ("passed" in waterComments.lower() if waterComments else False)
        
        if waterValue in [None, ""]:
            return waterData if waterData else ("0", "N/A")
        
        if validTest:
            # print(waterValue, modVal, waterComments, validTest)
            # print(isinstance(waterValue, (int, float)), isinstance(modVal, str), isinstance(waterComments, str))
            waterValue = "{:.0f}".format(waterValue) if isinstance(waterValue, (int, float)) else str(waterValue)
            waterData = (waterValue, waterComments if waterComments else "N/A")

        if loopIdx > 0:
            waterRow += 41
            commentsRow += 41
        else:
            waterRow += 28
            commentsRow += 41
        
        sectionRow += commentsRow + 13
        loopIdx += 1

def _gather_comments(ws, start_row, col):
    lines = []
    for i in range(11):
        val = ws.cell(row=start_row+i, column=col).value
        if val not in [None, ""]:
            lines.append(str(val))
    
    return "\n".join(lines)

def _check_section_pass(ws, start_row, col):
    try:
        sec1a = ws.cell(row=start_row, column=col).value
        sec1b = ws.cell(row=start_row+2, column=col).value
        sec2 = [ws.cell(row=start_row+5+i, column=col).value for i in range(3)]
    except:
        return False
    
    secOnePass = (sec1a == "Y" and sec1b =="Y") or sec1a != "Y"
    secTwoPass = all(val != "N" for val in sec2)

    return secOnePass and secTwoPass