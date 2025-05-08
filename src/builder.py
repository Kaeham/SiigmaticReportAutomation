from docx import Document
from src.sections import TitlePage, Summary, Deflection, OperationForce, AirInfiltration, Water, UltimateStrength, Appendix
from src.utils.FIleDialog import get_save_path

class DocumentBuilder:
    """
    Class which is responsible for building the report in the form of a word document
    """
    def __init__(self, data):
        self.data = data
        self.doc = Document()
        self.errors = []
    
    def build_all_sections(self):
        """
        Builds all the different sections of the reports using data
        """
        TitlePage.write(self.doc, self.data)
        Summary.write(self.doc, self.data, self.errors)
        Deflection.write(self.doc, self.data)
        OperationForce.write(self.doc, self.data)
        AirInfiltration.write(self.doc, self.data)
        Water.write(self.doc, self.data)
        UltimateStrength.write(self.doc, self.data)
        Appendix.start_all(self.doc)
    
    def save(self):
        path = get_save_path(self.data.filename)
        self.doc.save(path)
        # log_errors(self.errors)
        