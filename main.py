from docx import Document
from docx.shared import Inches

def main():
    document = Document()

    section = document.sections[0]
    header = section.header

    paragraph = header.paragraphs[0]
    paragraph.text = 'Madison County GIS Office'
    paragraph.style = document.styles['Header']

    #document.add_heading('Madison County GIS Office', 0)

    p = document.add_paragraph('Lorem ipsum')
    p.add_run('bold').bold = True
    p.add_run(' and some ')
    p.add_run('italic').italic = True

    #document.add_heading('Heading, level 1', level=0)

    document.add_page_break()

    document.save('demo.docx')

if __name__ == '__main__':
    main()