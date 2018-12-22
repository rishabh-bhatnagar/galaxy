import io
import re

from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfpage import PDFPage


def extract_text_by_page(pdf_path):
    with open(pdf_path, 'rb') as fh:
        for page in PDFPage.get_pages(fh,
                                      caching=True,
                                      check_extractable=True):
            resource_manager = PDFResourceManager()
            fake_file_handle = io.StringIO()
            converter = TextConverter(resource_manager, fake_file_handle)
            page_interpreter = PDFPageInterpreter(resource_manager, converter)
            page_interpreter.process_page(page)

            text = fake_file_handle.getvalue()
            yield text

            # close open handles
            converter.close()
            fake_file_handle.close()


def extract_text(pdf_path):
    for page in extract_text_by_page(pdf_path):
        yield page


def does_exist(regex, haystack):
    if re.search(regex, haystack, re.IGNORECASE) is not None:
        return True
    return False


def strip(string, weeds):
    i1 = 0
    i2 = 0
    while string[i1] in weeds:
        i1 += 1
    while string[i2] in weeds[::-1]:
        i2 += 1
    i2 = len(string) - i2 + 2
    return string[i1:i2]


def gen_to_list(f):
    def wrapper(*args):
        return list(f(*args))

    return wrapper


@gen_to_list
def split_list_str(l, string):
    for i in l:
        for j in i.split(string):
            if j:
                yield j


def self_re(present_fields, filtered_text):
    filtered_text = filtered_text.lower()
    present_fields = [i.lower() for i in present_fields]
    result = [filtered_text]
    for ff in present_fields:
        result = split_list_str(result, ff)
    # res = [strip(i, ':,. ').strip() for i in re.split("|".join(found_fields), ttext, flags=re.IGNORECASE) if i]
    return result

def main(pdf_name):
    fields = ['sales person', 'pot id', 'goapl opf no', 'opf date', 'customer name', 'Galaxy Billing from (Location)',
              'purchase order no', 'purchase date']
    text = " ".join(list(extract_text(pdf_name)))
    t_text = " ".join((re.split('Billing address', text, flags=re.IGNORECASE)[0]).split()[3:])
    found_fields = []
    for f in fields:
        if f.lower() in t_text.lower():
            found_fields.append(f)
    field_values = self_re(found_fields, t_text)
    field_values = [strip(fv, ':, .') for fv in field_values]
    for f, fv in zip(fields, field_values):
        print(f, ':', fv)


if __name__ == '__main__':
    main('pdf (2).pdf')