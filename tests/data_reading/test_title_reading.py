from src.data_reading.title_data import *
import openpyxl
import unittest

class TestTitleReading(unittest.TestCase):
    
    def __init__(self, methodName = "runTest"):
        super().__init__(methodName)
        self.wb = openpyxl.Workbook()
        self.ws = self.wb.active
        self.sampleData = [
            "Type A", "Window X", "RPT-001", "Melbourne",
            "Client Ltd", "Client Info", "Tester", "2024-05-09 14:30:00"
        ]
        for i, val in enumerate(self.sampleData):
            self.ws.cell(row=2+i, column=2, value=val)
    
    def test_extract_valid_title_data(self):
        expected = self.sampleData[:-1] + ["09-05-2024"]
        self.assertEqual(extract_title_data(self.ws), expected)
    
    def test_format_date_valid_string(self):
        res = format_date("2024-05-09 14:30:00")
        self.assertEqual(res, "09-05-2024")
    
    def test_format_date_invalid_string(self):
        res = format_date(1)
        self.assertEqual(res, "N/A")
    
    def test_format_date_none(self):
        self.assertEqual(format_date(None), "N/A")
    
    def test_format_date_partial(self):
        res = format_date("2024-05-09")
        self.assertEqual(res, "N/A")
    

