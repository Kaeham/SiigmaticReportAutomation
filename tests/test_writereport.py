import unittest
from unittest.mock import patch, MagicMock
from src.WriteReport import ReportGenerator
from src.DataReading import DataReader

class TestReportGenerator(unittest.TestCase):

    @patch('src.WriteReport.filedialog.askopenfilename', return_value='mock_file.xlsx')
    @patch('src.WriteReport.DataReader')
    def test_initialization_creates_document(self, MockReader, mock_file_dialog):
        mock_reader_instance = MagicMock()
        mock_reader_instance.get_data.return_value = (
            ["type", "name", "123", "loc", "client", "info", "tester", "2025-01-01", "toby"],  # body
            [("sash", "120", "130")],  # deflection
            [(1.1, 1.2), (0.9, 1.0)],  # oF
            ("2.3", "1.4"),            # AI
            ("200", "comment\n"),     # water
            ("2000", "1800", ["N"]*10),  # ultimate
            "800",                    # deflectionVal
            "MockFileName"            # filename
        )
        MockReader.return_value = mock_reader_instance

        gen = ReportGenerator()
        self.assertTrue(gen.document is not None)
        self.assertEqual(gen.filename, "MockFileName")

    @patch('src.WriteReport.filedialog.asksaveasfilename', return_value='mock_output.docx')
    @patch('src.WriteReport.Document.save')
    @patch('src.WriteReport.filedialog.askopenfilename', return_value='mock_file.xlsx')
    @patch('src.WriteReport.DataReader')
    def test_save_report_creates_file_when_confirmed(self, MockReader, mock_open, mock_save, mock_saveas):
        mock_reader_instance = MagicMock()
        mock_reader_instance.get_data.return_value = (...)  # same as previous
        MockReader.return_value = mock_reader_instance

        gen = ReportGenerator()
        gen.save_report()

        mock_save.assert_called_once()
    
    @patch('src.WriteReport.filedialog.asksaveasfilename', return_value='')
    @patch('src.WriteReport.Document.save')
    @patch('src.WriteReport.filedialog.askopenfilename', return_value='mock_file.xlsx')
    @patch('src.WriteReport.DataReader')
    def test_save_report_skips_when_cancelled(self, MockReader, mock_open, mock_save, mock_saveas):
        mock_reader_instance = MagicMock()
        mock_reader_instance.get_data.return_value = (...)  # same as previous
        MockReader.return_value = mock_reader_instance

        gen = ReportGenerator()
        gen.save_report()

        mock_save.assert_not_called()
    
    @patch('src.WriteReport.filedialog.askopenfilename', return_value='mock_file.xlsx')
    @patch('src.WriteReport.DataReader')
    def test_error_logging_detects_bad_input(self, MockReader, mock_open):
        mock_reader_instance = MagicMock()
        mock_reader_instance.get_data.return_value = (
            ["type", "name", "123", "loc", "client", "info", "tester", "2025-01-01", "toby"],
            [("sash", "120", "130")],
            [],  # empty oF triggers error
            ("", ""),  # empty AI
            ("", ""),  # empty water
            None,  # ultimateData
            "800",
            "MockFile"
        )
        MockReader.return_value = mock_reader_instance
        gen = ReportGenerator()
        self.assertGreater(len(gen.errors), 0)