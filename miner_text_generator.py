import io
import re
from tabula import read_pdf
from pandas import DataFrame as DF
from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfpage import PDFPage


class ParseError(Exception):
    pass


def extract_text_by_page(pdf_path):
    """
    returns the extracted text for a single pdf page.
    :param pdf_path: path to the pdf from which text is to be extracted.
    :return: list of text for a single page.
    """
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
    """
    Returns the text of all pages from pdf path.
    :param pdf_path: the path to pdf user wants to extract
    :return: the text of all the pages in the pdf
    """

    for page in extract_text_by_page(pdf_path):
        yield page


def does_exist(regex, haystack) -> bool:
    """
    Checks if pattern defined by regex is present in the haystack.
    :param regex: The pattern which is to be searched
    :param haystack: The string in which the given pattern is to be searched.
    :return: T/F to represent whether or not pattern is present in string
    """
    if re.search(regex, haystack, re.IGNORECASE) is not None:
        return True
    return False


def strip(string, weeds) -> list:
    """
    Removes all occurences of weeds from start and end of the string.
    :param string: The victim string which may have weeds in it.
    :param weeds: the unwanted occurence of strings in given string
    :return:
    """
    i1 = 0
    i2 = 0
    while string[i1] in weeds:
        i1 += 1
    while string[i2] in weeds[::-1]:
        i2 += 1
    i2 = len(string) - i2 + 2
    return string[i1:i2]


def gen_to_list(f) -> list:
    """
    Decorator to convert iterable to list.
    :param f: function to be decorated.
    :return: wrapper function.
    """

    def wrapper(*args):
        """
        :param args: All the args of function being wrapped.
        :return: list of the output after running f with given args.
        """
        return list(f(*args))

    return wrapper


@gen_to_list
def split_list_str(l, string) -> list:
    """
    Splits each string in the list l with string as the delimiter
    and unpacking it in the same list.
    :param l: list of strings for which split has to be performed recursively
    :param string: The string wrt to which elements in l will be splitted.
    :return: 1D list having all the splitted strings.
    """
    for i in l:
        for j in i.split(string):
            if j:
                yield j


def self_re_split(present_fields, filtered_text) -> list:
    """
    Re module's split was not behaving properly for unknown reasons compelling me to write this funciton
    to split a text based on the delimiters presented by present_fields
    :param present_fields: the fields based on which text is to be splitted.
    :param filtered_text: the text which is to be splitted.
    :return: List of splitted text based on the given list of delimiters.
    """
    filtered_text = filtered_text.lower()
    present_fields = [i.lower() for i in present_fields]
    result = [filtered_text]
    for ff in present_fields:
        result = split_list_str(result, ff)
    return result


def get_loose_data(name_of_pdf, all_fields) -> dict:
    """
    :param all_fields:The fields for which the values are to be found out
    :param name_of_pdf: pdf name from which the pdf will be read.
    :return: The dictionaries of the key:value.
    """
    text = " ".join(list(extract_text(name_of_pdf)))
    t_text = " ".join((re.split('Billing address', text, flags=re.IGNORECASE)[0]).split()[3:])
    found_fields = []
    for f in all_fields:
        if f.lower() in t_text.lower():
            found_fields.append(f)
    field_values = self_re_split(found_fields, t_text)
    field_values = [strip(fv, ':, .') for fv in field_values]
    return dict(zip(fields, field_values))


def merge_compliment_tables(t1, t2) -> DF:
    """
    Merges two complimented tables based on the presence or absence of data.
    :param t1: table 1
    :param t2: complement of table1
    :return: merged table

    Example of compliment table:

    0      0      0      0
    1     NaN     1       q
    2     NaN     2       w
    3     NaN     3       e
    4     NaN     4       r
    5       a     5     NaN
    6       b     6     NaN
    7       c     7     NaN

    Returns
    0       0
    1       q
    2       w
    3       e
    4       r
    5       a
    6       b
    7       c

    """
    table = []
    for i, j in zip(t1, t2):
        if str(i) != 'nan':
            table.append(i)
        else:
            table.append(j)
    return DF(table)


def bypass_occurence(string, l) -> str:
    """
    Removes all the occurences of given elements in l from string .
    :param string: The victim string which will be operated on.
    :param l: list having all strings to be bypassed.
    :return: The bypassed string.
    """
    "This function gets its name from push down automata\
     that bypassed the input character."

    for i in l:
        string = string.replace(i, '')
    return string


def get_attr(string, field_name, weeds) -> str:
    """
    Based on this pattern :
        String <= 'field_name: value'
    This function returns value removing all the weeds
    from the value obtained from splitting the string.
    :param string: The victim string having field_name:value
    :param field_name: Name of the field which is to be removed.
    :param weeds: unwanted occurences of strings which
                  hinders the understanding the value based on field_name.
    :return: The value based on the field name.
    """
    if len(
            re.split(
                field_name,
                string,
                flags=re.IGNORECASE
            )
    ) < 2:
        return None
    else:
        return \
            bypass_occurence(
                re.split(
                    field_name,
                    string,
                    flags=re.IGNORECASE
                )[1],  # [1] to get the value based on field_name.
                weeds
            ).strip()  # to get rid of any spaces around the string.


def transpose_table_list(table) -> list:
    """
    Transposes a dataframe(table) and returns the list of values.
    :param table: The dataframe.
    :return: list of values of dataframe.
    """
    transpose_table = table.T
    table_np_array = transpose_table.values
    table_list = table_np_array.tolist()
    return table_list


def extract_address(vtable):
    result = {}
    table = transpose_table_list(vtable)[0]
    result['which_address'] = table.pop(0)
    result['company'] = table.pop(0)
    address = []

    # This for loop appends string to address
    # until an element having string state is found.

    i: str
    for index, i in enumerate(table):
        if i.strip().lower().startswith('state'):
            break
        address.append(i.strip(','))

    # Hard coded list of rules to extract data from the data obtained.
    result['state'] = strip(re.split("state", i, flags=re.IGNORECASE)[1], ': ')
    result[result['which_address']] = ', '.join(address)
    result['contact_person'] = get_attr(table[idx + 1], 'contact person', ':')
    result['tel'] = get_attr(table[idx + 2], "tel", ': #')
    result['email'] = get_attr(table[idx + 3], "email", ': #-')
    result['gstn no'] = get_attr(table[idx + 4], 'gstn no', ' #:-')
    result['pan no'] = get_attr(table[idx + 5], 'pan no', ' #:-')

    return result


def print_dict(dicti) -> None:
    """
    Prints all the key value pair in  {Key : Value} pattern.
    :param dicti: The dictionary which is to be printed.
    :return: None
    """
    for key, value in dicti.items():
        print("{} : {}".format(key, value))


def get_first_table_data(name_of_pdf):
    """
    :param name_of_pdf: Name of pdf present in pwd.
    :return: The extracted data from the fist tables of the pdf.
    """
    tables = read_pdf(name_of_pdf, multiple_tables=True)
    table_0 = tables[0]
    temp_table_0 = list()
    temp_table_0.append(merge_compliment_tables(table_0[0], table_0[1]))
    temp_table_0.append(merge_compliment_tables(table_0[2], table_0[3]))
    tables[0] = temp_table_0

    if len(list(tables[0][0].T)) < 7:
        raise ParseError("OPF Tables cannot be parsed")
    else:
        return [
            extract_address(tables[0][1]),
            extract_address(tables[0][0])
        ]


def combine_dicts(dictionaries) -> dict:
    """
    :param dictionaries: list of dictionaries to combine
    :return: Single combined dictionary
    """
    d = {}
    for dicti in dictionaries:
        d.update(dicti)
    return d


if __name__ == '__main__':
    pdf_name = 'pdf (3).pdf'
    fields = ['sales person', 'pot id', 'goapl opf no', 'opf date', 'customer name', 'Galaxy Billing from (Location)',
              'purchase order no', 'purchase date']

    print('Loose Data:')
    print_dict(
        get_loose_data(
            name_of_pdf=pdf_name,
            all_fields=fields
        )
    )
    print("\n", end='')

    for idx, dictionary in enumerate(
            get_first_table_data(
                name_of_pdf=pdf_name
            )
    ):
        print("Table {}:".format(idx))
        if dictionary is not None:
            print_dict(dictionary)
        else:
            print(dictionary)
        print('\n', end='\n')
