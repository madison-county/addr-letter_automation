from docx import Document
from docx.shared import Inches

document = Document()

document.add_heading('Madison County GIS Office', 0)

p = document.add_paragraph('Lorem ipsum')
p.add_run('bold').bold = True
p.add_run(' and some ')
p.add_run('italic').italic = True

document.add_heading('Heading, level 1', level=1)

document.add_page_break()