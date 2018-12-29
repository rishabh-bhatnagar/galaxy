from docx import Document

document = Document('OPF.docx')
table = document.tables[1]
for i in table.rows:
    for j in i.cells:
        print(j.text, end='  ##  ')
    print("\n\n")
