import io
import re
from tabula import read_pdf
from pandas import DataFrame as df
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


def merge_compliment_tables(t1, t2):
    """
    :param t1: table 1
    :param t2: complement of table1
    :return: merged table

    Example of compliment table:

    0 ColumnA     0 ColumnB
    1     NaN     1       q
    2     NaN     2       w
    3     NaN     3       e
    4     NaN     4       r
    5       a     5     NaN
    6       b     6     NaN
    7       c     7     NaN
    """
    table = []
    for i, j in zip(t1, t2):
        if str(i) != 'nan':
            table.append(i)
        else:
            table.append(j)
    return df(table)


def bypass_occurence(string, l):
    for i in l:
        string = string.replace(i, '')
    return string

def get_attr(string, field_name, weeds):
    if len(re.split(field_name, string, flags=re.IGNORECASE)) < 2:
        return None
    return bypass_occurence(re.split(field_name, string, flags=re.IGNORECASE)[1], weeds).strip()

def extract_address(vtable):
    result = {}
    table = vtable.T.values.tolist()[0]
    result['which_address'] = table.pop(0)
    address = []
    for idx, i in enumerate(table):
        if i.strip().lower().startswith('state'):
            break
        address.append(i.strip(','))
    result['state'] = strip(re.split("state", i, flags=re.IGNORECASE)[1], ': ')
    address = ', '.join(address)
    result[result['which_address']] = address
    idx += 1
    result['contact_person'] = get_attr(table[idx], 'contact person', ':')
    result['tel'] = get_attr(table[idx+1], "tel", ': #')
    result['email'] = get_attr(table[idx + 2], "email", ': #-')
    result['gstn no'] = get_attr(table[idx+3], 'gstn no', ' #:-')
    result['pan no'] = get_attr(table[idx+4], 'pan no', ' #:-')
    return result

if __name__ == '__main__':
    pdf_name = 'pdf (2).pdf'
    main(pdf_name)
    tables = read_pdf(pdf_name, multiple_tables=True)
    final_tables = []

    table_0 = tables[0]
    temp_table_0 = list()
    temp_table_0.append(merge_compliment_tables(table_0[0], table_0[1]))
    temp_table_0.append(merge_compliment_tables(table_0[2], table_0[3]))
    tables[0] = temp_table_0
    if len(list(tables[0][0].T)) < 7:
        print("OPF Tables cannot be parsed")
    else:
        address = extract_address(tables[0][1])
        for key, value in address.items():
            print("{} : {}".format(key, value))
