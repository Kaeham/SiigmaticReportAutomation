# src/data/utils.py

def format_decimal(val, dp=1):
    """
    Format a numeric value to N decimal places.
    """
    return f"{val:.{dp}f}" if isinstance(val, (int, float)) else str(val)