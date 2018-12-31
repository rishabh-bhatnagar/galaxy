from docx import Document
import shutil
import os.path
from pandas import DataFrame as DF
import re
import os
from os import listdir
from os.path import abspath
import win32com.client as win32
from win32com.client import constants

state_mapping = dict({
    "bangalore": "Karnataka",
    'tamil nadu': 'tamilnadu',
    'hyderabad': 'andhra pradesh'
})
state_mapping.setdefault(None, '')


class OPF():
    def __init__(self, path):
        self.path = path
        if '/' in self.path:
            self.file_name = self.path.split('/')[-1]
        else:
            self.file_name = self.path.split('\\')[-1]

    def seperate_doc(self, docx_folder_name='Docxs', docs_folder_name='Docs'):
        if '.docx' not in self.path and '.doc' in self.path:
            # file is a doc file but not a docx file
            prev_path = self.path
            self.path = self.save_as_docx()
            self.move_file(file_path=prev_path, to_path=docs_folder_name)
        self.move_file(to_path=docx_folder_name)

    def save_as_docx(self):
        # Code credits: https://stackoverflow.com/questions/38468442/multiple-doc-to-docx-file-conversion-using-python
        # Opening Microsoft word application.
        word = win32.gencache.EnsureDispatch('Word.Application')
        doc = word.Documents.Open(self.path)
        doc.Activate()

        # Rename path with .docx
        new_file_abs = os.path.abspath(path)
        new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)

        # Save and Close
        word.ActiveDocument.SaveAs(
            new_file_abs, FileFormat=constants.wdFormatXMLDocument
        )
        doc.Close(False)
        return new_file_abs

    def move_file(self, to_path, file_path=None):
        if file_path is None:
            file_path = self.path
        if not os.path.exists(to_path):
            os.makedirs(to_path)
        shutil.move(file_path, to_path)

    def get_tables(self):
        self.document = Document(self.path)
        return self.document.tables

    @staticmethod
    def create_table(docx_table):
        table_primitive = []
        for row in docx_table.rows:
            table_row = []
            for cell in row.cells:
                table_row.append(cell.text.strip('\n').strip(' ').replace('\n', ' '))
            table_primitive.append(table_row)
        return table_primitive

    @staticmethod
    def print_table(table_data):
        widths = [max(map(len, col)) for col in zip(*table_data)]
        for row in table_data:
            print(" # ".join((val.ljust(width) for val, width in zip(row, widths))))

    def extract_data(self):
        tables = self.get_tables()
        address_table = self.create_table(tables[0])
        sales_table = self.create_table(tables[1])

        address_data = self.parse_address_tables(address_table)
        sales_data = self.parse_sales_data(sales_table)
        loose_fields = self.get_loose_fields()

        final_result = address_data.copy()
        final_result.update(sales_data)
        final_result.update(loose_fields)
        state = final_result.get('badstate')
        state = state if state is not None else ''
        if state:
            state = state.lower()
        final_result.update(dict(state_to_dc_state=state_mapping.get(state)))
        if not final_result.get('state_to_dc_state'):
            final_result['state_to_dc_state'] = state.capitalize()

        return final_result

    def get_element_from_block(self, block: list, identifier: str, split_by: str) -> str:
        if identifier == 'Galaxy Billing from (Location)':
            identifier = identifier.replace('(', '\(').replace(')', '\)')
        identifier = identifier.replace(" ", '\s*')+'\s*'
        for i in block:
            if re.search(identifier, i, flags=re.IGNORECASE):
                probable_result = split_by.join(i.split(split_by)[1:])
                if not probable_result.strip():
                    probable_result = ":".join(i.split(':')[1:])
                return probable_result.strip("-").strip(":").strip(":").strip("-").strip('_').strip()

    def parse_address_tables(self, address_table):
        def parse_address_block(block):
            which_address = block.pop(0)
            fields = ['State:', 'Contact Person:', "Tel#", "Email#", "GSTN NO:"]
            a = block[0]

            # index before which address is present.
            first_field_index = None
            for index, element in enumerate(block):
                for field in fields:
                    if field[:-1] in element:
                        first_field_index = index
                        break
                if first_field_index is not None:
                    break
            if first_field_index is None:
                return {}  # an empty dict suggesting failure in parsing all fields.
            name = block[0]
            address = "\n".join(block[:first_field_index])
            block = block[first_field_index:]
            state = self.get_element_from_block(block, 'State', ':')
            if state:
                if re.search('Mumbai', state, flags=re.IGNORECASE):
                    # print('state is Mumbai??')
                    state = 'Maharashtra'
            if not state:
                if re.search('Mumbai', address, flags=re.IGNORECASE):
                    state = 'Maharashtra'
                    # print('setting state as Maharashtra')
                # print('state not found')
            if state:
                print(state)
                res_state = state_mapping.get(state.lower())
                state = res_state if res_state else state
            contact_person = self.get_element_from_block(block, 'Contact Person', ':')
            email = self.get_element_from_block(block, 'Email', '#')
            tel = self.get_element_from_block(block, 'tel', '#')
            if self.get_element_from_block(block, 'GST', ':') is not None:
                for i in block:
                    if 'GST' in i:
                        break
                pan = None
                if 'pan no' in i.lower():
                    gstn, pan = re.split('PAN NO', i, flags=re.IGNORECASE)
                else:
                    gstn, pan = i, ''
                gstn = gstn.split(':')[-1]
                pan = pan.split(":-")[-1]
            else:
                gstn = pan = ''
            return dict(
                name=name,
                address=address,
                pan=pan,
                gstn=gstn,
                state=state,
                contact_person=contact_person,
                email=email,
                tel=tel
            )

        address_table = [list(i) for i in zip(*address_table)]
        billing_address = {'bad' + k: v for k, v in parse_address_block(address_table[0]).items()}
        delivery_address = {'dad' + k: v for k, v in parse_address_block(address_table[1]).items()}
        billing_address.update(delivery_address)
        return billing_address

    def parse_sales_data(self, sales_table):
        header = sales_table.pop(0)
        result = {}
        i = 0
        while True:
            i += 1
            sr_no = sales_table[0][0]
            if not [i for i in sr_no if i.isdigit()]: break
            result.update({'desc_' + sr_no: sales_table[0][1]})
            result.update({'qty_' + sr_no: sales_table[0][2]})
            result.update({'unit_price_' + sr_no: sales_table[0][3]})
            result.update({'total_price_' + sr_no: sales_table[0][1]})
            sales_table.pop(0)
        i -= 1
        i = max([i, 1])
        result.update(dict(number_of_products=i))
        for i in sales_table:
            for j in i:
                if 'CGST' in j:
                    result.update(cgst_percentage="".join([k for k in j if k.isdigit()]))
                elif 'SGST' in j:
                    result.update(sgst_percentage="".join([k for k in j if k.isdigit()]))
                elif 'IGST' in j:
                    result.update(igst_percentage="".join([k for k in j if k.isdigit()]))
        return result

    def get_loose_fields(self):
        try:
            self.document
        except:
            self.document = Document(self.path)
        paragraphs = self.document.paragraphs
        lines = []
        for paragraph in paragraphs:
            for run in paragraph.runs:
                if '\t' in run.text or ' '*4 in run.text:
                    lines.append(paragraph)
                    break
        line_texts = list()  # strings of all the
        texts = list()
        for line in lines:
            line_texts.append(line.text.replace('\t', '    ').strip())
        line_texts = [i for i in line_texts if i]
        for text in line_texts:
            count_space = 1
            while ' '*count_space in text:
                count_space += 1
            else:
                count_space -= 1
            count_space = max([count_space, 1])
            texts.extend(text.split(count_space*' '))
        '''
        texts = [element for nest in
                 [k for k in
                  [
                      [
                          i for i in "".join([i.text for i in j.runs]).split('\t') if i.strip()
                      ] for j in lines
                  ] if k
                  ] for element in nest
                 ]
        '''
        payment_terms = ''
        for i in paragraphs:
            if re.search('PAYMENT TERMS', i.text, flags=re.IGNORECASE):
                payment_terms = self.get_element_from_block([i.text], 'PAYMENT TERMS', ":")

        return dict(
            sales_person=self.get_element_from_block(texts, 'Sales Person', ":"),
            pot_id=self.get_element_from_block(texts, 'POT ID', ":"),
            opf_no=self.get_element_from_block(texts, 'OPF No.', "."),
            customer_name=self.get_element_from_block(texts, 'Customer Name', ":"),
            opf_date=self.get_element_from_block(texts, 'OPF Date', ":"),
            opf_location=self.get_element_from_block(texts, 'Galaxy Billing from (Location)', ":"),
            purch_order_no=self.get_element_from_block(texts, 'Purchase Order', "."),
            payment_terms=payment_terms
        )


if __name__ == '__main__':
    docx_folder_name = 'Docxs'
    for path in listdir():
        if '.doc' in path:
            OPF(abspath(path)).seperate_doc()
    os.chdir(abspath(docx_folder_name))
    result_dict_list = []
    opf = OPF('OPF- TK-024.docx').extract_data()
    for i, file_name in enumerate(listdir()):
        if '.docx' in file_name:
            print(file_name)
            opf = OPF(file_name)
            data_dict = opf.extract_data()
            data_dict.update({'opf link': file_name.split('.')[0]})
            result_dict_list.append(data_dict)
    result_dict_list = [i for i in result_dict_list if i['badstate']]
    # result_dict_list = sorted(result_dict_list, key=lambda x:x['badstate'] if x['badstate'] is not None else '')
    df = DF(result_dict_list)
    all_keys = list(df.keys())
    header = [
                 all_keys.pop(all_keys.index('state_to_dc_state')),
                 all_keys.pop(all_keys.index('dadstate')),
                 all_keys.pop(all_keys.index('opf_no')),
                 all_keys.pop(all_keys.index('customer_name')),
                 all_keys.pop(all_keys.index('purch_order_no')),
                 all_keys.pop(all_keys.index('opf_date')),
                 all_keys.pop(all_keys.index('payment_terms')),
                 all_keys.pop(all_keys.index('dadname')),
                 all_keys.pop(all_keys.index('dadaddress')),
                 all_keys.pop(all_keys.index('dadgstn')),
                 all_keys.pop(all_keys.index('badname')),
                 all_keys.pop(all_keys.index('badaddress')),
                 all_keys.pop(all_keys.index('badgstn')),
                 all_keys.pop(all_keys.index('number_of_products')),
             ]
    desc = sorted([i for i in all_keys if 'desc' in i], key=lambda x: int("".join([i for i in x[5:] if i.isdigit()])))
    qty = sorted([i for i in all_keys if 'qty' in i], key=lambda x: int("".join([i for i in x[4:] if i.isdigit()])))
    unit_price = sorted([i for i in all_keys if 'unit_price' in i], key=lambda x: int("".join([i for i in x[11:] if i.isdigit()])))
    all_keys = [i for i in all_keys if i not in desc+unit_price+qty]
    rest_fields = all_keys.pop()
    unpacked_tuples = []
    for i in zip(desc, qty, unit_price):
        unpacked_tuples += i
    df = df[header+unpacked_tuples]
    df.to_excel('../final_output.xlsx')  # saving out of docx folder.
