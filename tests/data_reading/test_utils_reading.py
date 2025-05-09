import unittest
from src.data_reading.utils import *

class TestUtils(unittest.TestCase):
    def __init__(self, methodName = "runTest"):
        super().__init__(methodName)
    
    def test_valid_input(self):
        res = format_decimal(1.21212, 2)
        self.assertEqual(res, "1.21")
    
    def test_invalid_input(self):
        res = format_decimal("221")
        self.assertEqual(res, "221")
    
    def test_none_input(self):
        res = format_decimal(None, 2)
        self.assertEqual(res, 'None')