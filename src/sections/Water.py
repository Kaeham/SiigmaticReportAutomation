# sections/water.py
from src.utils.Formatter import format_text, format_table_width, format_cell_color, insert_line_break
from src.utils.Constants import DEFAULT_CELL_COLOR

def write(doc, data, ai=[]):
    """
    Renders the Water Penetration section.
    :param doc: python-docx Document
    :param data: ReportData
    """
    _add_heading(doc)
    _add_results_table(doc, data.water)
    _add_compliance(doc, data.water[0], ai)
    doc.add_page_break()


def _add_heading(doc):
    para = doc.add_paragraph()
    format_text(para.add_run("Water Penetration\n"), 14, bold_status=True, underline_status=True)
    format_text(para.add_run("Method\n"), 12, bold_status=True)
    format_text(para.add_run("The test specimen was tested in accordance with AS4420.1 Clause 6.\n"), 12)
    format_text(para.add_run("\nResults"), 12, bold_status=True)


def _add_results_table(doc, water_data):
    pressure, comments = water_data
    table = doc.add_table(rows=2, cols=2, style="Table Grid")
    table.cell(0, 0).text = "Test Pressure (Pa)"
    table.cell(0, 1).text = "Observations"
    table.cell(1, 0).text = str(pressure)
    table.cell(1, 1).text = comments or "N/A"

    format_text(table.cell(1, 1).paragraphs[0].runs[0], 12)
    format_cell_color(table.cell(1, 1), DEFAULT_CELL_COLOR)
    format_table_width(table.cell(1, 0), 2)
    format_table_width(table.cell(1, 1), 10)


def _add_compliance(doc, pressure, ai):
    para = doc.add_paragraph()
    format_text(para.add_run("\nCompliance\n"), 12, bold_status=True)
    format_text(para.add_run(f"The specimen satisfied water penetration requirements of AS2047 Clause 2.3.1.6 at {pressure}Pa.\n"), 12)
    format_text(para.add_run("\nComment:"), 12)
    print(ai)
    if ai:
        ai = ai if isinstance(ai, list) else [ai]
        for comment in ai:
            format_text(para.add_run(f"\n{comment}"), 12)
    else:
        format_text(para.add_run(" Nil\n"), 12)
    insert_line_break(para)
