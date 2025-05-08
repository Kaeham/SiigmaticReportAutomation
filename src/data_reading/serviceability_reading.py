def get_serviceability_data(ws):
    """
    extract the serviceability design wind pressure value
    Inputs:
        - ws: the excel file where the data is stored
    Outputs:
        - Serviceability Wind Pressure: if the value is a number, returns the value otherwise returns N/A
    """
    serviceability = ws.cell(row=14, column=3).value
    try:
        serviceability = int(serviceability)
        return str(serviceability)
    except:
        return "N/A"