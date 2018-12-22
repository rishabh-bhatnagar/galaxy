from pyPdf import PdfFileWriter, PdfFileReader

output = PdfFileWriter()
input1 = PdfFileReader(file("document1.pdf", "rb"))

# print the title of document1.pdf
print "title = %s" % (input1.getDocumentInfo().title)

# add page 1 from input1 to output document, unchanged
output.addPage(input1.getPage(0))

# add page 2 from input1, but rotated clockwise 90 degrees
output.addPage(input1.getPage(1).rotateClockwise(90))

# add page 3 from input1, rotated the other way:
output.addPage(input1.getPage(2).rotateCounterClockwise(90))
# alt: output.addPage(input1.getPage(2).rotateClockwise(270))

# add page 4 from input1, but first add a watermark from another pdf:
page4 = input1.getPage(3)
watermark = PdfFileReader(file("watermark.pdf", "rb"))
page4.mergePage(watermark.getPage(0))

# add page 5 from input1, but crop it to half size:
page5 = input1.getPage(4)
page5.mediaBox.upperRight = (
    page5.mediaBox.getUpperRight_x() / 2,
    page5.mediaBox.getUpperRight_y() / 2
)
output.addPage(page5)

# print how many pages input1 has:
print "document1.pdf has %s pages." % input1.getNumPages()

# finally, write "output" to document-output.pdf
outputStream = file("document-output.pdf", "wb")
output.write(outputStream)
outputStream.close()

'''
# csv_exporter.py
 
import csv
import os
 
from miner_text_generator import extract_text_by_page
 
 
def export_as_csv(pdf_path, csv_path):
    filename = os.path.splitext(os.path.basename(pdf_path))[0]
 
    counter = 1
    with open(csv_path, 'w') as csv_file:
        writer = csv.writer(csv_file)
        for page in extract_text_by_page(pdf_path):
            text = page[0:100]
            words = text.split()
            writer.writerow(words)
 
 
if __name__ == '__main__':
    pdf_path = 'pdf.pdf'
    csv_path = 'pdf.csv'
    export_as_csv(pdf_path, csv_path)
'''
