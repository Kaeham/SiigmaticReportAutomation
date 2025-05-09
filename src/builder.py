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
    
    def build_all_sections(self, ai=[]):
        """
        Builds all the different sections of the reports using data
        """
        defAi = oFAi = aIAi = waterAi = ultimateAi = []
        if ai != []:
            defAi = ai[0]
            oFAi = ai[1]
            aIAi = ai[2]
            waterAi = ai[3]
            ultimateAi = ai[4]
        TitlePage.write(self.doc, self.data)
        Summary.write(self.doc, self.data, self.errors)
        Deflection.write(self.doc, self.data, defAi)
        OperationForce.write(self.doc, self.data, oFAi)
        AirInfiltration.write(self.doc, self.data, aIAi)
        Water.write(self.doc, self.data, waterAi)
        UltimateStrength.write(self.doc, self.data, ultimateAi)
        Appendix.start_all(self.doc)
    
    def save(self):
        path = get_save_path(self.data.filename)
        self.doc.save(path)
        # log_errors(self.errors)
        