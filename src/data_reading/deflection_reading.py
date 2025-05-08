def extract_deflection_data(ws):
    pos = get_all_members(ws, (10, 7))
    neg = get_all_members(ws, (38, 7))
    result = []

    for name, r, c in pos:
        posVal = validate_sensor(ws, r+8, c+4)
        result.append((name, posVal, "placeholder"))

    for name, r, c in neg:
        idx = match_member_by_name(result, name)
        negVal = validate_sensor(ws, r+8, c+4)
        if idx > -1:
            name, posVal, _ = result[idx]
            result[idx] = (name, posVal, negVal)
        else:
            result.append((name, "N/A", negVal))

    return result

def validate_sensor(ws, row, col):
    """
    """
    val = ws.cell(row=row, column=col).value
    for i in range(1, 4):
        if ws.cell(row=row, column= col -(i*2)).value == 0:
            return "N/A"
    return "{:.0f}".format(val) if isinstance(val, (int, float)) else "N/A"

def get_all_members(ws, start):
    """
    From a start cell, collect all valid sash/member labels.
    """
    row, col = start
    names = []
    defaultNames = {"", None, "x sash"}
    default = False
    while not default:
        name = ws.cell(row=row, column=col).value
        name = str(name).lower() if isinstance(name, str) else ""
        if name in defaultNames:
            default = True
        idx = match_member_by_name(names, name)
        if idx > -1:
            names[idx] = (name, row, col)
        else:
            names.append((name, row, col))
        row += 56
    return names

def match_member_by_name(data, name):
    """
    Find index of a named member in a list of tuples.
    """
    for i, (n, *_rest) in enumerate(data):
        if n == name:
            return i
    return -1