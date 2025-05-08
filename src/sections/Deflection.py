from src.utils.Formatter import format_text, format_paragraph_spacing, insert_line_break, format_cell_color
from src.utils.Constants import DEFAULT_CELL_COLOR

def write(doc, data):
    """
    Renders the deflection section of the report
    """
    _add_heading(doc)

    for name, pos, neg in data.deflections:
        _add_sash_table(doc, name, pos, neg, data.deflection_val)
    
    _add_compliance_text(doc)
    doc.add_page_break()

def _add_heading(doc):
    """
    """
    para = doc.add_paragraph()
    format_text(para.add_run("Deflection\n"), 14, bold_status=True, underline_status=True)
    format_text(para.add_run("Method\n"), 12, bold_status=True)
    format_text(para.add_run("The test specimen was tested in accordance with AS4420.1 Clause 3.\n"), 12)
    format_text(para.add_run("\nResults"), 12, bold_status=True)
    format_paragraph_spacing(para)

def _add_sash_table(doc, sash_name, pos, neg, pressure):
    """
    """
    para = doc.add_paragraph()
    format_paragraph_spacing(para)
    format_text(para.add_run(sash_name.capitalize()), 12)
    
    table = doc.add_table(rows=3, cols=3, style="Table Grid")

    table.cell(0, 0).paragraphs[0].add_run("Pressure Differential").bold = True
    table.cell(0, 1).paragraphs[0].add_run("Pressure Achieved (Pa)").bold = True
    table.cell(0, 2).paragraphs[0].add_run("Resultant Span Ratio").bold = True

    table.cell(1, 0).paragraphs[0].add_run("Positive").bold = True
    table.cell(1, 1).paragraphs[0].add_run(str(pressure))
    table.cell(1, 2).paragraphs[0].add_run(f"Span/{pos}")

    table.cell(2, 0).paragraphs[0].add_run("Negative").bold = True
    table.cell(2, 1).paragraphs[0].add_run(f"-{str(pressure)}")
    table.cell(2, 2).paragraphs[0].add_run(f"Span/{neg}")

    # Formatting
    for row in table.rows[1:]:
        for cell in row.cells[1:]:
            format_text(cell.paragraphs[0].runs[0], 12)
            format_cell_color(cell, DEFAULT_CELL_COLOR)
    
    doc.add_paragraph()

def _add_compliance_text(doc):
    """
    """
    format_text(doc.add_paragraph().add_run("(See sash numbering in Appendix: Elevation drawing)\n"), 12, italicize=True)
    para = doc.add_paragraph()
    format_text(para.add_run("\nCompliance\n"), 12, bold_status=True)
    format_text(para.add_run("The test specimen satisfied the deflection requirements of AS2047 Clause 2.3.1.3 at the specified pressure.\n"), 12)
    format_text(para.add_run("\nComment: Nil"), 12)
    insert_line_break(para)