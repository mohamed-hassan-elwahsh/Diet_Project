import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import copy
import os

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "..", "templates", "weekly_diet_template.docx")

def set_cell_text_preserve_style(cell, text):
    """
    Sets text of a cell while preserving its original style, font, size, direction, and IMAGES.
    Centers the text.
    """
    drawings = []
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            for child in run._r:
                if child.tag.endswith('drawing') or child.tag.endswith('pict'):
                    drawings.append(copy.deepcopy(child))
                    
    font_name = "Segoe UI"
    font_size = Pt(13)
    bold = False
    
    if cell.paragraphs and cell.paragraphs[0].runs:
        first_run = cell.paragraphs[0].runs[0]
        if first_run.font.name:
            font_name = first_run.font.name
        if first_run.font.size:
            font_size = first_run.font.size
        bold = bool(first_run.font.bold)
        
    cell.text = ""
    
    lines = text.split("\n")
    for i, line in enumerate(lines):
        if i == 0:
            p = cell.paragraphs[0]
        else:
            p = cell.add_paragraph()
            
        run = p.add_run(line)
        run.font.name = font_name
        run.font.size = font_size
        run.font.bold = bold
        
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        pPr = p._p.get_or_add_pPr()
        bidi = OxmlElement('w:bidi')
        bidi.set(qn('w:val'), '1')
        pPr.append(bidi)
        
    if drawings:
        p = cell.paragraphs[-1]
        img_run = p.add_run('\n')
        for d in drawings:
            img_run._r.append(d)


def generate_patient_diet_docx(patient_name, plan_date, weekly_plan, output_path):
    doc = docx.Document(TEMPLATE_PATH)
    
    for day_idx, day_data in enumerate(weekly_plan[:7]):
        if day_idx < len(doc.tables):
            table = doc.tables[day_idx]
            for meal_idx, meal in enumerate(day_data.get("meals", [])):
                row_idx = meal_idx + 1
                if row_idx < len(table.rows):
                    row = table.rows[row_idx]
                    if len(row.cells) > 1 and meal.get("meal_type"):
                        set_cell_text_preserve_style(row.cells[1], meal["meal_type"])
                    if len(row.cells) > 2 and meal.get("timing"):
                        set_cell_text_preserve_style(row.cells[2], meal["timing"])
                    if len(row.cells) > 3 and meal.get("content"):
                        set_cell_text_preserve_style(row.cells[3], meal["content"])

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)
    return output_path
