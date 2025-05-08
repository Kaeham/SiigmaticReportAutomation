from src.data_reading.utils import format_decimal

def extract_force_data(ws):
    oFData = [(0, 0), (0, 0)]
    r = 10
    c = 14
    while True:
        vals = [ws.cell(row=r, column=c+i).value for i in range(2)]
        vals2 = [ws.cell(row=r+6, column=c+i).value for i in range(2)]

        if all(v in ["", None, 0, '#DIV/0!'] for v in vals + vals2):
            break

        oFData[0] = (format_decimal(vals[0]), format_decimal(vals[1]))
        oFData[1] = (format_decimal(vals2[0]), format_decimal(vals2[1]))

        r += 15

    return oFData
