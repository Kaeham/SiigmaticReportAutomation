# src/data/utils.py

def safe_cell(ws, row, col, default=""):
    """
    Safely get a cell value, with fallback.
    """
    val = ws.cell(row=row, column=col).value
    return val if val not in ["", None, "#DIV/0!", 0] else default


def format_decimal(val, dp=1):
    """
    Format a numeric value to N decimal places.
    """
    return f"{val:.{dp}f}" if isinstance(val, (int, float)) else str(val)