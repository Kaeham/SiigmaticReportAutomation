from src.data_reading.utils import format_decimal

def extract_air_data(ws):
    row, col = 17, 20
    latest = ("", "")
    while True:
        pos = ws.cell(row=row, column=col).value
        neg = ws.cell(row=row+1, column=col).value
        if pos in ["", 0, None, '#DIV/0!']:
            break
        
        pos = format_decimal(pos, 2) if isinstance(pos, (float, int)) and pos != 0 else "N/A"
        neg = format_decimal(neg, 2) if isinstance(neg, (float, int)) and neg != 0 else "N/A"
        latest = (pos, neg)
        row += 18
    return latest
