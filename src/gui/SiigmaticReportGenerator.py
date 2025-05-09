from PyQt5.QtWidgets import QWidget, QVBoxLayout, QStackedWidget
from src.gui.components.create_sections import *
from src.builder import DocumentBuilder
from src.data_reading import reader
from src.utils.FIleDialog import select_input_file
from src.ai_checking import ai_air_infiltration_checks, ai_deflection_check, ai_operating_forces_checks, ai_ultimate_strength_checks, ai_water_checks
from src.gui.ai_analysis import AIAnalysis

class SiigmaticReportGenerator:
    def __init__(self, parent:QStackedWidget):
        self.root = QWidget()
        self.parent = parent
        parent.addWidget(self.root)
        parent.setCurrentIndex(0)
        self.root.setWindowTitle("Siigmatic Report Generator")
        self.root.setFixedSize(600, 300)
        self.mainLayout = QVBoxLayout()
        self.root.setLayout(self.mainLayout)
        self.builder = False
        self.ai = False
        self.report = False

        # Sections
        self.fileFrame, self.fileText, self.fileButton = create_file_section()
        self.featuresFrame, self.buttons = create_feature_section()

        self.mainLayout.addWidget(self.fileFrame)
        self.mainLayout.addWidget(self.featuresFrame)

        self._connect_signals()

    def _connect_signals(self):
        self.buttons["save"].clicked.connect(self.save_report)
        self.buttons["ai"].clicked.connect(self._go_to_ai)
        self.buttons["report"].clicked.connect(self.generate_report)
        self.fileButton.clicked.connect(self._select_file)
    
    def _select_file(self):
        file = select_input_file()
        self.fileText.setText(file)

    def _go_to_ai(self):
        if not self.ai and self.report:
            ai = [ai_deflection_check.deflection_checks(self.reportData.deflections), 
                    ai_operating_forces_checks.operating_force_checks(self.reportData.operating_forces),
                    ai_air_infiltration_checks.air_infiltration_checks(self.reportData.air_data),
                    ai_water_checks.water_penetration_checks(self.reportData.water),
                    ai_ultimate_strength_checks.ultimate_strength_checks(self.reportData.ultimate)]
            self.aiFrame = AIAnalysis(ai, self.parent, self)
            self.ai = True
        if self.report:
            self.parent.setCurrentIndex(1)
    
    def generate_report(self):
        file = self.fileText.text()
        if file != "":
            self.reportData = reader.load_report_data(file)
            self.builder = DocumentBuilder(self.reportData)
            self.report = True
        else:
            # add console output to window
            print("Enter Valid File !!")

    def save_report(self):
        # handle save logic, or delegate to a utility module
        if self.builder:
            self.reportData = reader.load_report_data(self.fileText.text())
            self.builder.build_all_sections(self.aiFrame.selectedItems)
            self.builder.save()
            print("Saving report...")
        else:
            print("Generate Report First!!")
            # Include logic to retrieve report data and write to file

    def get_window(self):
        return self.root

    def run(self):
        self.root.show()
        self.app.exec()
