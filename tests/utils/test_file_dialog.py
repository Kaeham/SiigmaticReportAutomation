import unittest
from unittest.mock import patch
from src.utils.FIleDialog import select_input_file, get_save_path  # adjust import if needed

class TestFileDialogs(unittest.TestCase):

    @patch("tkinter.filedialog.askopenfilename")
    @patch("tkinter.Tk")
    def test_select_input_file_returns_path(self, mock_tk, mock_askopenfilename):
        mock_askopenfilename.return_value = "/fake/path/input.xlsx"

        result = select_input_file()
        self.assertEqual(result, "/fake/path/input.xlsx")

    @patch("tkinter.filedialog.asksaveasfilename")
    @patch("tkinter.Tk")
    def test_get_save_path_returns_path(self, mock_tk, mock_asksaveasfilename):
        mock_asksaveasfilename.return_value = "/fake/path/output.docx"

        result = get_save_path("output.docx")
        self.assertEqual(result, "/fake/path/output.docx")

    @patch("tkinter.filedialog.asksaveasfilename", return_value="")
    @patch("tkinter.Tk")
    def test_get_save_path_user_cancel(self, mock_tk, mock_dialog):
        result = get_save_path()
        self.assertEqual(result, "")
