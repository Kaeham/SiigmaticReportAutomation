import unittest
from src.data_reading.schema import *

class TestReportData(unittest.TestCase):
    def test_full_initialization(self):
        data = ReportData(
            body_info=["Window", "Series X", "TR-123", "Hallam", "ClientCo", "123 Info", "John", "01-01-2024", "Jane"],
            deflections=[("sash 1", 250, 240), ("sash 2", 255, 245)],
            operating_forces=[(12.3, 6.8), (10.5, 4.2)],
            air_data=("0.89", "1.01"),
            water=("450", "No penetration noted."),
            ultimate=("5200", "5000", ["Y", "Y", "Y", "Y", "Y", "N", "N", "N", "N", "N"]),
            deflection_val=800,
            filename="TR-123"
        )

        self.assertEqual(data.body_info[2], "TR-123")
        self.assertEqual(data.deflection_val, 800)
        self.assertEqual(data.ultimate[0], "5200")
        self.assertEqual(len(data.ultimate[2]), 10)
        self.assertEqual(data.air_data[1], "1.01")

    def test_empty_initialization(self):
        data = ReportData(
            body_info=[],
            deflections=[],
            operating_forces=[],
            air_data=("", ""),
            water=(0, ""),
            ultimate=(0, 0, []),
            deflection_val=0,
            filename=""
        )

        self.assertEqual(data.body_info, [])
        self.assertEqual(data.deflection_val, 0)
        self.assertEqual(data.ultimate[2], [])
        self.assertEqual(data.filename, "")