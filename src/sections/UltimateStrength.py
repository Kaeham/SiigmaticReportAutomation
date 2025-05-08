# sections/ultimate.py
from src.utils.Formatter import format_text, format_table_width, format_cell_color, insert_line_break
from src.utils.Constants import DEFAULT_CELL_COLOR

def write(doc, data):
    """
    Renders the Ultimate Strength Test section.
    :param doc: python-docx Document
    :param data: ReportData
    """
    _add_heading(doc)
    _add_pressures_table(doc, data.ultimate)
    _add_observations_table(doc, data.ultimate[2])
    _add_compliance(doc)
    doc.add_page_break()


def _add_heading(doc):
    para = doc.add_paragraph()
    format_text(para.add_run("Ultimate Strength Test\n"), 14, bold_status=True, underline_status=True)
    format_text(para.add_run("Method\n"), 12, bold_status=True)
    format_text(para.add_run("The test was conducted in accordance with AS4420.1 Clause 7.\n"), 12)
    format_text(para.add_run("\nResults"), 12, bold_status=True)


def _add_pressures_table(doc, ultimate_data):
    pos, neg, _ = ultimate_data
    table = doc.add_table(rows=2, cols=2, style="Table Grid")
    table.cell(0, 0).text = "Positive Test Pressure Pa (Negative Chamber Pressure)"
    table.cell(0, 1).text = str(pos)
    table.cell(1, 0).text = "Negative Test Pressure Pa (Positive Chamber Pressure)"
    table.cell(1, 1).text = str(neg)

    for row in table.rows:
        format_text(row.cells[0].paragraphs[0].runs[0], 10, bold_status=True)
        format_text(row.cells[1].paragraphs[0].runs[0], 10)
        format_cell_color(row.cells[1], DEFAULT_CELL_COLOR)
    format_table_width(table.cell(0, 0), 3.5)
    format_table_width(table.cell(0, 1), 1.5)


def _add_observations_table(doc, obs_data):
    doc.add_paragraph("\n")
    table = doc.add_table(rows=7, cols=3, style="Table Grid")
    table.cell(0, 0).merge(table.cell(0, 2)).text = "Observations - Pressure Differential"
    format_text(table.cell(0, 0).paragraphs[0].runs[0], 11, bold_status=True)
    table.cell(0, 0).paragraphs[0].alignment = 1

    headers = ["", "Positive Pressure (Y/N)", "Negative Pressure (Y/N)"]
    for i, h in enumerate(headers):
        table.cell(1, i).text = h
        table.cell(1, i).paragraphs[0].alignment = 1
        format_text(table.cell(1, i).paragraphs[0].runs[0], 11, bold_status=True)

    obsLabels = [
        "Failure or dislodgement of any glazing",
        "Dislodgement of a frame or any part of a frame",
        "Removal of a light, either with or without its framing sash, from a frame",
        "Loss of support of a frame, such as when it is unstable in its opening in the building structure",
        "Failure of any sash, locking device, fastener or supporting stay, allowing an opening light to open"
        ]

    for i, label in enumerate(obsLabels):
        table.cell(i+2, 0).paragraphs[0].add_run(label)
        table.cell(i+2, 1).paragraphs[0].add_run(obs_data[i] if i < len(obs_data) else "N/A")
        table.cell(i+2, 2).paragraphs[0].add_run(obs_data[i+5] if i+5 < len(obs_data) else "N/A")

            
        for j in [1, 2]:
            format_text(table.cell(i+2, j).paragraphs[0].runs[0], 9)
            format_cell_color(table.cell(i+2, j), DEFAULT_CELL_COLOR)

    format_table_width(table.cell(1, 0), 4.5)

    for row in table.rows[2:]:
        format_text(row.cells[0].paragraphs[0].runs[0], 9)
    

def _add_compliance(doc):
    para = doc.add_paragraph()
    format_text(para.add_run("\nCompliance\n"), 12, bold_status=True)
    format_text(para.add_run("The specimen satisfied Ultimate Strength requirements per AS2047 Clause 2.3.1.7.\n"), 12)
    format_text(para.add_run("\nComment: Nil"), 12)
    insert_line_break(para)