def extract_ultimate_data(ws):
    """
    Extract Ultimate Strength Test pressure and observations.
    
    Returns:
        tuple: (positive_pressure: str, negative_pressure: str, observations: list[str])
    """
    row, col = 3, 25          # Positive/negative pressure start
    obsRow, obsCol = 6, 26  # Observations start
    ultimateData = ("N/A", "N/A", ["N/A"]*10)

    while row <= 100000:
        posVal = ws.cell(row=row, column=col).value
        negVal = ws.cell(row=row+1, column=col).value

        if posVal in ["", 0, None] or negVal in ["", 0, None]:
            break
        elif any(not isinstance(val, (float, int)) for val in [posVal, negVal]):
            break

        # Collect 10 observation cells (5 positive, 5 negative)
        observations = []
        for i in range(5):
            val = ws.cell(row=obsRow + i, column=obsCol).value
            observations.append(val if val is not None else "")
        obsRow += 9
        for i in range(5):
            val = ws.cell(row=obsRow + i, column=obsCol).value
            observations.append(val if val is not None else "")
        
        # Check for empty data — early exit

        ultimateData =  (
            f"{posVal:.0f}" if isinstance(posVal, (int, float)) else str(posVal),
            f"{negVal:.0f}" if isinstance(negVal, (int, float)) else str(negVal),
            observations
        )

        # Continue to next block
        row += 19
        obsRow += 19

    return ultimateData
