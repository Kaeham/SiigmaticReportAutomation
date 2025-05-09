# File to place all n rating related functions
def n_rating_serviceability(deflection_value):
    """
    Get the n rating for the serviceabiity tests
    """
    print(deflection_value, type(deflection_value))
    if not isinstance(deflection_value, (int, float)):
        return ("", "")
    gen = [(1600, "N6, C4"), (1200, "N5, C3"), (800, "N4, C2"), (600, "N3, C1"), (400, "N2")]
    corner = [(2500, "N6, C4"), (1800, "N5, C3"), (1200, "N4, C2"), (800, "N3, C1"), (600, "N2")]
    
    finalGenRating = None
    finalCornerRating= None
    
    for idx in range(len(gen)):
        genValue, genRating = gen[idx]
        corValue, corRating = corner[idx]
        
        if deflection_value >= genValue and finalGenRating == None:
            finalGenRating = genRating + " (General)"
        
        if deflection_value >= corValue and finalCornerRating == None:
            finalCornerRating = corRating + " (Corner)"
        
    return (finalGenRating, finalCornerRating)

def n_rating_air_results(posAir, negAir):
    """
    get the air infiltration levels of the specimen
    """
    try: posAir = float(posAir)
    except: return ("", "")
    try: negAir = float(negAir) 
    except: return ("", "")

    if posAir <= 1:
        if negAir <= 1:
            return "Low Air Infiltration"
    elif posAir <= 5:
        return "High Air Infiltration"
    
    return ("", "")

def n_rating_water(water_val):
    """
    Get the n rating for the water penetration
    """    
    nonExposed = [(450, "N6, C4"), (300, "N5, C3"), (200, "N4, C2"), (150, "N3, C1")]
    exposed = [(600, "N6, C4"),  (450, "N5, C3"), (300, "N4, C2"), (200, "N1, N2")]
    nonExposedRating = None
    exposedRating = None
    try:
        water_val = float(water_val)
    except:
        return ("", "")

    for idx in range(len(nonExposed)):
        exposedNValue, exposedNRating = exposed[idx]
        if water_val >= exposedNValue and exposedRating == None:
            exposedRating = exposedNRating + " (Exposed)"
        
        nonExposedNValue, nonExposedNRating = nonExposed[idx]
        if water_val >= nonExposedNValue and nonExposedRating == None:
            nonExposedRating = nonExposedNRating + " (Non-Exposed)"
    
    return (exposedRating, nonExposedRating)

def n_rating_ust(ultimate_value):
    """
    Get the n rating for the ultimate strength test
    Inputs:
        ultimate_value: the weakest pressure rating from UST
    Outputs:
        tuple containing the general and corner window rating respectively
    """
    general = [(5300, "N6, C4"), (4000, "N6, C3"), (3000, "N5, C2") , (2700, "N4, C2"),
               (2000, "N4, C1"), (1800, "N3, C1"), (1400, "N3"), (900, "N2"), (600, "N1")]
    corner = [(8000, "N6, C4"), (6000, "N6, C3"), (5900, "N5, C3"),  (4500, "N5, C2"),
              (4000, "N4, C2"), (3000, "N4, C1"), (2700, "N3, C1"), (2000, "N3"), (1300, "N2"), (900, "N1")]
    genRating = None
    cornRating = None
    try:
        ultimate_value = float(ultimate_value)
    except:
        return ("", "")

    for idx in range(len(general)):
        genNValue, genNRating = general[idx]
        if ultimate_value >= genNValue and genRating == None:
            genRating = genNRating + " (General)"
    
    for idx in range(len(corner)):
        cornNValue, cornNRating = corner[idx]
        if ultimate_value >= cornNValue and cornRating == None:
            cornRating = cornNRating + " (Corner)"
    
    return (genRating, cornRating)