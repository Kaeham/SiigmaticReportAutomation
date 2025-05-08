# sections/summary.py
from src.utils.Formatter import format_text, format_summary_page, insert_line_break, format_table_width
from src.utils.NRatingFinder import *  # For rating functions

def write(doc, data, errors):
    summary = doc.add_paragraph(style="Heading 1").add_run("Summary")
    format_text(summary, 20, bold_status=True, underline_status=True)

    sashCount, oFMembers, rowIndices = _row_count(data)
    table = doc.add_table(rows=rowIndices["ultimate"]+1, cols=5, style="Table Grid")
    _setup_table_layout(table, (sashCount, oFMembers))
    _fill_table(table, data, errors, rowIndices)

    para = doc.add_paragraph()
    format_text(para.add_run("NR: measurement not required, hereinafter.\n\n"), 12)

    format_summary_page(table)
    insert_line_break(para)
    doc.add_page_break()

def _row_count(data):
    sash = max(2, len(data.deflections))
    op = max(2, len(data.operating_forces))
    return (sash, op, {
        "ser": 1,
        "def": 2,
        "of": 2 + sash * 2,
        "air": 2 + sash * 2 + op,
        "water": 2 + sash * 2 + op + 2,
        "ultimate": 2 + sash * 2 + op + 3,
    })

def _setup_table_layout(table, counts):
    sashCount, oFSashCount = counts
    titleRowIdx = 0
    serRowIdx = titleRowIdx + 1
    defRowIdx = serRowIdx + 1
    oFRowIdx = defRowIdx + sashCount*2
    aIRowIdx = oFRowIdx + oFSashCount
    waterRowIdx = aIRowIdx + 2
    ultimateRowIdx = waterRowIdx + 1

    # Merges and headers for the first column and top row
    table.cell(defRowIdx, 0).merge(table.cell(oFRowIdx - 1, 0)) # column 1 merging
    table.cell(oFRowIdx, 0).merge(table.cell(aIRowIdx - 1, 0))
    table.cell(aIRowIdx, 0).merge(table.cell(waterRowIdx - 1, 0)) 
    table.cell(titleRowIdx, 0).paragraphs[0].add_run("Test Description").bold = True # col 1 title 
    table.cell(serRowIdx, 0).text = "Serviceability Design Wind Pressure"
    table.cell(defRowIdx, 0).text = "Deflection/Span Ratio Structural member 1"
    table.cell(oFRowIdx, 0).text = "Operational Force"
    table.cell(aIRowIdx, 0).text = "Air infiltration at +- 75Pa"
    table.cell(waterRowIdx, 0).text = "Water Penetration"
    table.cell(ultimateRowIdx, 0).text = "Ultimate Strength Test Pressure"
    for row in table.rows:
        format_table_width(row.cells[0], 1.5)
        format_table_width(row.cells[1], 2)
        # format_table_width(row.cells[2], 1)
        format_table_width(row.cells[3], 1.5)
        format_table_width(row.cells[4], 1)

    table.cell(titleRowIdx, 1).merge(table.cell(titleRowIdx, 2)) # column 2 & 3
    table.cell(serRowIdx, 1).merge(table.cell(serRowIdx, 2))
    table.cell(waterRowIdx, 1).merge(table.cell(waterRowIdx, 2))
    table.cell(ultimateRowIdx, 1).merge(table.cell(ultimateRowIdx, 2))
    table.cell(titleRowIdx, 1).paragraphs[0].add_run("Test Results").bold = True
    table.cell(aIRowIdx, 1).paragraphs[0].add_run("+ 75")
    table.cell(aIRowIdx + 1, 1).paragraphs[0].add_run("- 75")
    table.cell(serRowIdx, 3).merge(table.cell(oFRowIdx - 1, 3)) # column 4
    table.cell(oFRowIdx, 3).merge(table.cell(aIRowIdx - 1, 3))
    table.cell(aIRowIdx, 3).merge(table.cell(waterRowIdx - 1, 3))
    table.cell(titleRowIdx, 3).paragraphs[0].add_run("Rating").bold = True
    table.cell(serRowIdx, 4).merge(table.cell(oFRowIdx - 1, 4)) # column 5
    table.cell(oFRowIdx, 4).merge(table.cell(aIRowIdx - 1, 4))
    table.cell(aIRowIdx, 4).merge(table.cell(waterRowIdx-1, 4))
    table.cell(titleRowIdx, 4).paragraphs[0].add_run("Verdict").bold = True

def _fill_table(table, data, errors, row_idxs):
    
    _fill_serviceability(table, data, row_idxs)
    _fill_deflection(table, data, row_idxs)
    _fill_operating_force(table, data, row_idxs, errors)
    _fill_air_infiltration(table, data, row_idxs, errors)
    _fill_water(table, data, row_idxs, errors)
    _fill_ultimate(table, data, row_idxs, errors)

def _fill_serviceability(table, data, idxs):
    gen, cor = n_rating_serviceability(data.deflection_val)
    row = idxs["ser"]
    table.cell(row, 1).text = f"±{data.deflection_val} Pa"
    if gen and cor:
        table.cell(row, 3).text = f"{gen}\n{cor}"
    table.cell(row, 1).paragraphs[0].alignment = 1

def _fill_deflection(table, data, idxs):
    row = idxs["def"]
    sashNum = 1
    plus = True

    for i in range(len(data.deflections) * 2):
        desc = f"Vertical Sash on Panel (V0{sashNum}){'+' if plus else '-'}"
        table.cell(row + i, 1).text = desc
        plus = not plus
        if i % 2: sashNum += 1

    i = row
    for sash in data.deflections:
        for val in sash[1:]:
            table.cell(i, 2).text = f"Span/{val}"
            i += 1

def _fill_operating_force(table, data, idxs, errors):
    row = idxs["of"]
    panel = 1

    for i in range(0, len(data.operating_forces), 2):
        table.cell(row + i, 1).text = f"Panel {panel} Open Initiate / Maintain"
        table.cell(row + i + 1, 1).text = f"Panel {panel} Close Initiate / Maintain"

        try:
            open_vals = data.operating_forces[i]
            close_vals = data.operating_forces[i + 1]
        except IndexError:
            open_vals = close_vals = ("N/A", "N/A")
            errors.append("Error: Operating Force Data not recorded properly")

        table.cell(row + i, 2).text = f"{open_vals[0]} / {open_vals[1]}"
        table.cell(row + i + 1, 2).text = f"{close_vals[0]} / {close_vals[1]}"
        panel += 1

def _fill_air_infiltration(table, data, idxs, errors):
    row = idxs["air"]
    pos, neg = data.air_data
    table.cell(row, 1).text = "At +75Pa"
    table.cell(row + 1, 1).text = "At -75Pa"
    table.cell(row, 2).text = f"{pos}/sm²"
    table.cell(row + 1, 2).text = "NR" if neg == "0" else f"{neg}/sm²"

    rating = n_rating_air_results(pos, neg)
    if rating:
        table.cell(row, 3).text = rating
    else:
        errors.append("Error: Air Infiltration Data not recorded properly")

def _fill_water(table, data, idxs, errors):
    row = idxs["water"]
    try:
        pressure, _ = data.water
        table.cell(row, 1).text = f"{pressure}Pa"
        table.cell(row, 1).paragraphs[0].alignment = 1
        exp, non_exp = n_rating_water(pressure)
        rating = "\n".join(filter(None, [exp, non_exp]))
        table.cell(row, 3).text = rating
    except Exception:
        errors.append("Error: Water Penetration not recorded properly")

def _fill_ultimate(table, data, idxs, errors):
    row = idxs["ultimate"]
    try:
        pos, neg, _ = data.ultimate
        table.cell(row, 1).text = f"+{pos}Pa\n-{neg}Pa"
        table.cell(row, 1).paragraphs[0].alignment = 1
        gen, cor = n_rating_ust(min(pos, neg))
        table.cell(row, 3).text = f"{gen}\n{cor}"
    except Exception:
        errors.append("Error: Ultimate Data not recorded properly")
