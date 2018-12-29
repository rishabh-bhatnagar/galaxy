from docx import Document
'''
document = Document('OPF.docx')
table = document.tables[1]
for i in table.rows:
    for j in i.cells:
        print(j.text, end='  ##  ')
    print("\n\n")
'''
from glob import glob
import re
import os
from os import listdir
from os.path import abspath
import win32com.client as win32
from win32com.client import constants

# Create list of paths to .doc files
paths = glob('.\\*.doc')

def save_as_docx(path):
    # Opening MS Word
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(path)
    doc.Activate ()

    # Rename path with .docx
    new_file_abs = os.path.abspath(path)
    new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)

    # Save and Close
    word.ActiveDocument.SaveAs(
        new_file_abs, FileFormat=constants.wdFormatXMLDocument
    )
    doc.Close(False)

for path in listdir():
    if '.doc' in path:
        print(abspath(path))
        save_as_docx(abspath(path))