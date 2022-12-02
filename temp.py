from docxtpl import DocxTemplate
from docx import Document
file='Doc1.docx'

doc = Document(file)
doc.add_paragraph('')
doc.save(file)


tpl = DocxTemplate(file)
context = {}
tpl.render(context)
tpl.save(file)