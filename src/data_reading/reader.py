from src.data_reading import title_data, deflection_reading, operation_force_reading, air_infiltration_reading, water_reading, ultimate_reading, serviceability_reading
from .schema import ReportData
import openpyxl
import os

def load_report_data(path:str) -> ReportData:
    # rawData = DataReader(path).get_data()
    workbook = openpyxl.load_workbook(path, data_only=True).active
    bodyInfo = title_data.extract_title_data(workbook)
    deflections = deflection_reading.extract_deflection_data(workbook)
    operatingForces = operation_force_reading.extract_force_data(workbook)
    airData = air_infiltration_reading.extract_air_data(workbook)
    water = water_reading.extract_water_data(workbook)
    ultimate = ultimate_reading.extract_ultimate_data(workbook)
    deflectionVal = serviceability_reading.get_serviceability_data(workbook)
    filename = os.path.splitext(os.path.basename(path))[0]
    
    return ReportData(
        bodyInfo,
        deflections,
        operatingForces,
        airData,
        water,
        ultimate,
        deflectionVal,
        filename
    )