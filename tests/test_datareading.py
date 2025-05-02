import unittest
import os
from src.DataReading import DataReader
from dotenv import load_dotenv
load_dotenv()

class TestDataReader(unittest.TestCase):
    def __init__(self, methodName = "runTest"):
        super().__init__(methodName)
        file = os.environ.get("EXCEL_FILE")
        secondFile = os.environ.get("EXCEL_FILE_2")
        emptyFile = os.environ.get("EMPTY_EXCEL_FILE")

        self.filledReader = DataReader(file)
        self.filledWorkbook = self.filledReader.get_workbook()
        self.secondReader = DataReader(secondFile)
        self.secondWorkbook = self.secondReader.get_workbook()
        self.emptyReader = DataReader(emptyFile)
        self.emptyWorkbook = self.emptyReader.get_workbook()

    def test_valid_data(self):
        """
        Test whether the data read by Data Reader is not empty
        """
        res = self.filledReader.get_data()
        for item in res:
            self.assertNotEqual(item, None)
    
    def test_title_data_return_value(self):
        """
        Test the return value for title_data
        """
        titles, name = self.filledReader.get_title_data(self.filledWorkbook)
        self.assertEqual(type(titles), list) # check if titles are a list
        self.assertEqual(type(name), str) # check is name is a string

        for title in titles:
            self.assertNotEqual(title, None)
    
    def test_format_date(self):
        """
        Test whether date is formatted properly
        """
        unformattedDate = "2025-05-02 11:11:11"
        self.assertEqual(self.filledReader.format_date(unformattedDate), "02-05-2025")
        unformattedDate = "0"
        self.assertEqual(self.filledReader.format_date(unformattedDate), "N/A")

    def test_serviceability_data_return_value(self):
        """
        Test whether serviceability values are extracted properly
        """
        defaultData = self.emptyReader.get_serviceability_data(self.emptyWorkbook)
        normalData = self.filledReader.get_serviceability_data(self.filledWorkbook)
        self.assertEqual(type(defaultData), str)
        self.assertEqual(type(normalData), str)

    def test_deflection_data_return_value(self):
        """
        Test whether the deflection values are extracted properly
        """
        defaultData = self.emptyReader.get_deflection_data(self.emptyWorkbook)
        filledData = self.secondReader.get_deflection_data(self.secondWorkbook)
        self.assertEqual(defaultData, [])
        self.assertEqual(filledData, [("Vertical Sash on Panel".lower(), '2764', '-15308')])

    def test_get_operation_force_return_value(self):
        """
        Test whether the operation force is extracted properly
        """
        defaultData = self.emptyReader.get_operational_force(self.emptyWorkbook)
        filledData = self.secondReader.get_operational_force(self.secondWorkbook)
        self.assertEqual(defaultData, ((0, 0), (0, 0)))
        self.assertEqual(filledData, (('11.1', '2.1'), ('1.5', '0.3')))
    
    def test_get_aI_data_return_value(self):
        """
        test whether the air infiltration data is extracted properly
        """
        defaultData = self.emptyReader.get_aI_data(self.emptyWorkbook)
        filledData = self.filledReader.get_aI_data(self.filledWorkbook)
        secondData = self.secondReader.get_aI_data(self.secondWorkbook)
        self.assertEqual(defaultData, ("", ''))
        self.assertEqual(filledData, ('2.94', 'N/A'))
        self.assertEqual(secondData, ('1.14', 'N/A'))
    
    def test_get_water_data_return_value(self):
        """
        Test whether get_water_data() returns the correct output
        """
        defaultData = self.emptyReader.get_water_data(self.emptyWorkbook)
        filledData = self.filledReader.get_water_data(self.filledWorkbook)
        secondData = self.secondReader.get_water_data(self.secondWorkbook)
        self.assertEqual(defaultData, ('150', 'N/A'))
        self.assertEqual(filledData, ('200', 'N/A'))
        self.assertEqual(secondData, ('300', 'n5 achieved\n'))
    
    def test_get_ultimate_strength_data_return_value(self):
        """
        Test whether get_ultimate_data() returns the correct output
        """
        defaultData = self.emptyReader.get_ultimate_data(self.emptyWorkbook)
        filledData = self.filledReader.get_ultimate_data(self.filledWorkbook)
        secondData = self.secondReader.get_ultimate_data(self.secondWorkbook)
        self.assertEqual(defaultData, ("N/A", "N/A", ["N/A"]*10))
        self.assertEqual(filledData, ('2000', '2000', ['N']*10))
        self.assertEqual(secondData, ('-3000', '2729', ['N']*10))



if __name__ == '__main__':
    unittest.main()