from PyQt5.QtWidgets import (QWidget, QPushButton, QGridLayout, QFileDialog,
                             QLineEdit, QLabel, QFrame, QVBoxLayout)
from ..WriteReport import ReportGenerator
import datetime
# from ..WriteReport import final

class SiigmaticReportGenerator():
    """
    Class to represent the main window
    """
    def __init__(self):
        self.root = QWidget()
        self.root.setWindowTitle("Siigmatic Report Generator")
        self.root.setFixedSize(600, 300)
        
        # ─── Main Layout ──────────────────────────────
        self.mainGrid = QVBoxLayout()
        self.root.setLayout(self.mainGrid)

        # ─── File Section ─────────────────────────────
        self.fileFrame = QFrame()
        self.fileFrame.setStyleSheet("""
            QFrame {
                border: 1px solid #cccccc;
                border-radius: 8px;
                background-color: #f9f9f9;
                padding: 12px;
            }
            QLabel {
                font-weight: bold;
                font-size: 11pt;
            }
        """)
        self.fileGrid = QGridLayout()
        self.fileFrame.setLayout(self.fileGrid)

        self.label = QLabel("File to create a report for")
        self.fileText = QLineEdit()
        self.fileButton = QPushButton("Select File")
        self.fileButton.setFixedSize(100, 32)
        self.style_button(self.fileButton)

        self.fileGrid.addWidget(self.label, 0, 0, 1, 2)
        self.fileGrid.addWidget(self.fileText, 1, 0)
        self.fileGrid.addWidget(self.fileButton, 1, 1)
        self.fileGrid.setColumnStretch(0, 3)
        self.fileGrid.setColumnStretch(1, 1)

        # ─── Features Section ─────────────────────────
        self.featuresFrame = QFrame()
        self.featuresFrame.setStyleSheet("""
            QFrame {
                border: 1px solid #cccccc;
                border-radius: 8px;
                background-color: #f9f9f9;
                padding: 12px;
            }
        """)
        self.featuresGrid = QGridLayout()
        self.featuresFrame.setLayout(self.featuresGrid)

        self.aiButton = QPushButton("AI Analysis")
        self.reportButton = QPushButton("Generate Report")
        self.saveButton = QPushButton("Save")

        for btn in [self.aiButton, self.reportButton, self.saveButton]:
            btn.setFixedSize(140, 32)
            self.style_button(btn)

        self.featuresGrid.addWidget(self.aiButton, 0, 0)
        self.featuresGrid.addWidget(self.reportButton, 1, 0)
        self.featuresGrid.addWidget(self.saveButton, 1, 1)
        self.add_functions()
        self.featuresGrid.setColumnStretch(0, 1)
        self.featuresGrid.setColumnStretch(1, 1)

        # ─── Add Sections to Main Grid ────────────────
        self.mainGrid.addWidget(self.fileFrame)
        self.mainGrid.addWidget(self.featuresFrame)

    def style_button(self, button):
        """
        Apply consistent styling to all QPushButton elements
        """
        button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                padding: 6px 12px;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)

    def add_functions(self):
        """
        Add functionality to each button in the GUI
        """
        self.saveButton.clicked.connect(self.save_report)
        self.reportButton.clicked.connect()
    
    def get_filename(self):
        """
        Get the default filename for this report
        """
        filepath = self.fileText.text()
        
        final = 0
        for idx in range(len(filepath)):
            char = filepath[idx]
            if char == "/" or char == "\\":
                final = idx
        return filepath[final+1:]
    
    def generate_report(self):
        self.reportGenerator = ReportGenerator(self.fileText.text())
        
    def save_report(self):
        print("help")
        try:
            print("Save triggered")
            filename = self.get_filename()
            self.report = self.reportGenerator.get_report()
            savePath, _ = QFileDialog.getSaveFileName(
                self.root,
                "Save Report As: ",
                f"{filename}",
                "Word Documents (*.docx)"
            )

            if savePath:
                self.report.save(savePath)
                print(f"File saved at: {savePath}")
            else:
                print("No file location selected. File not saved.")
            
            if self.report.errors:
                timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                errorFileName = f"error_log_{timestamp}.txt"

                with open(errorFileName, "w", encoding="utf-8") as errorFile:
                    errorFile.write("ERROR LOG\n\n")
                    for error in self.report.errors:
                        errorFile.write(error + "\n")
                
                print(f"Errors detected! See log: {errorFileName}")
            print("Save DOne")
        except Exception as e:
            print(f"error during save {e}")

    def get_window(self):
        """
        Returns the top level widget containing the contents of the class
        """
        return self.root

    def run(self):
        self.root.show()
    