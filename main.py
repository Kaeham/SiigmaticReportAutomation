import sys
from PyQt5.QtWidgets import QApplication, QStackedWidget
from src.gui.SiigmaticReportGenerator import SiigmaticReportGenerator

# from src.data_reading.reader import load_report_data
# from src.builder import DocumentBuilder
# from src.utils.FIleDialog import select_input_file

class App(QStackedWidget):
    def __init__(self):
        super().__init__()

        self.siigmatic = SiigmaticReportGenerator(self)

def main():
    app = QApplication(sys.argv)
    window = App()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()