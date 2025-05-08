import openpyxl

def load_workbook(file):
    return openpyxl.load_workbook(file, data_only=True).active