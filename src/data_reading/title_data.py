def extract_title_data(ws):
    row = 2
    col = 2
    data = [str(ws.cell(row=row+i, column=col).value) for i in range(8)]
    data[7] = format_date(data[7])
    return data

def format_date(date_str):
    from datetime import datetime
    try:
        return datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S').strftime('%d-%m-%Y')
    except Exception:
        return "N/A"
