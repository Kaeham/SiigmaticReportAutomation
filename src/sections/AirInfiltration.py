# sections/air.py
from src.utils.Formatter import format_text, format_paragraph_spacing, format_cell_color, insert_line_break
from src.utils.Constants import DEFAULT_CELL_COLOR

def write(doc, data):
    """
    Renders the Air Infiltration section.
    :param doc: python-docx Document
    :param data: ReportData
    """
    _add_heading(doc)
    _add_results_table(doc, data.air_data)
    _add_compliance(doc, data.air_data)
    doc.add_page_break()


def _add_heading(doc):
    para = doc.add_paragraph()
    format_text(para.add_run("Air Infiltration\n"), 14, bold_status=True, underline_status=True)
    format_text(para.add_run("Method\n"), 12, bold_status=True)
    format_text(para.add_run("The test specimen was tested in accordance with AS4420.1 Clause 5.\n"), 12)
    format_text(para.add_run("Results"), 12, bold_status=True)
    format_paragraph_spacing(para)


def _add_results_table(doc, air_data):
    posVal, negVal = air_data
    if negVal == "0": negVal = "N/A"

    table = doc.add_table(rows=4, cols=5, style="Table Grid")
    table.cell(0, 0).merge(table.cell(0, 4)).paragraphs[0].add_run("Measured Flow (Ls⁻¹m⁻²)").bold = True

    headers = ["Level", "Test Pressure", "Allowable", "Positive", "Negative"]
    for i, text in enumerate(headers):
        format_text(table.cell(1, i).paragraphs[0].add_run(text), 12, bold_status=True)

    data_rows = [
        ["Low", "±75", "< 1.0", posVal, negVal],
        ["High", "", "< 5.0", posVal, "N/A"]
    ]

    for r, row in enumerate(data_rows, start=2):
        for c, val in enumerate(row):
            table.cell(r, c).text = str(val)
            if r > 1 and c > 2:
                format_cell_color(table.cell(r, c), DEFAULT_CELL_COLOR)
                format_text(table.cell(r, c).paragraphs[0].runs[0], 12)


def _add_compliance(doc, air_data):
    para = doc.add_paragraph()
    posVal, negVal = map(lambda x: float(x) if x not in ["NR", "N/A"] else 0.0, air_data)

    if posVal < 1.0 and negVal < 1.0:
        rating = "Low"
    elif posVal < 5.0:
        rating = "High"
    else:
        rating = "Unrated"

    format_text(para.add_run("Compliance\n"), 12, bold_status=True)
    format_text(para.add_run(f"The test specimen satisfied the “{rating}” Air Infiltration Level requirements of AS2047 Clause 2.3.1.5\n"), 12)
    format_text(para.add_run("\nComment: Nil\n"), 12)
    insert_line_break(para)
