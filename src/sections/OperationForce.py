# sections/operation_force.py
from src.utils.Formatter import format_text, format_paragraph_spacing, format_cell_color, insert_line_break
from src.utils.Constants import DEFAULT_CELL_COLOR

def write(doc, data):
    """
    Renders the Operating Force section.
    :param doc: python-docx Document
    :param data: ReportData
    """
    _add_heading(doc)
    _add_results(doc, data.operating_forces)
    _add_compliance(doc)
    doc.add_page_break()

def _add_heading(doc):
    para = doc.add_paragraph()
    format_text(para.add_run("Operating Force\n"), 14, bold_status=True, underline_status=True)
    format_text(para.add_run("Method\n"), 12, bold_status=True)
    format_text(para.add_run("The test specimen was tested in accordance with AS4420.1 Clause 4.\n"), 12)
    format_text(para.add_run("\nResults"), 12, bold_status=True)
    # format_paragraph_spacing(para)

def _add_results(doc, forces):
    panelNumber = 1

    for i in range(0, len(forces), 2):
        openVals = forces[i]
        closeVals = forces[i + 1] if i + 1 < len(forces) else ("N/A", "N/A")

        para = doc.add_paragraph()
        format_paragraph_spacing(para)
        format_text(para.add_run(f"Panel {panelNumber}"), 12)

        table = doc.add_table(rows=3, cols=3, style="Table Grid")

        headers = ["", "Initiating Force", "Maintaining Force"]
        for col, text in enumerate(headers):
            format_text(table.cell(0, col).paragraphs[0].add_run(text), 12, bold_status=True)

        # Opening row
        table.cell(1, 0).text = "Opening"
        table.cell(1, 1).text = f"{openVals[0]}N"
        table.cell(1, 2).text = f"{openVals[1]}N"

        # Closing row
        table.cell(2, 0).text = "Closing"
        table.cell(2, 1).text = f"{closeVals[0]}N"
        table.cell(2, 2).text = f"{closeVals[1]}N"

        for row in table.rows[1:]:
            for cell in row.cells[1:]:
                format_text(cell.paragraphs[0].runs[0], 12)
                format_cell_color(cell, DEFAULT_CELL_COLOR)

        panelNumber += 1

def _add_compliance(doc):
    para = doc.add_paragraph()
    format_text(para.add_run("\nCompliance\n"), 12, bold_status=True)
    format_text(para.add_run("The test specimen satisfied the requirements of AS2047 Clause 2.3.1.4\n"), 12)
    format_text(para.add_run("\nComment: Nil\n"), 12)
    insert_line_break(para)
    
