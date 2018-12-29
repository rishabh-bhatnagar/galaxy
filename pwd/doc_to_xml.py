from docx import Document
import shutil
import os.path
from pandas import DataFrame as DF
from glob import glob
import re
import os
from os import listdir
from os.path import abspath
import win32com.client as win32
from win32com.client import constants

# Create list of paths to .doc files
paths = glob('.\\*.doc')


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
        document = Document(self.path)
        return document.tables
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
        final_result = address_data.copy()
        final_result.update(sales_data)
        return final_result
    def get_element_from_block(self, block: list, identifier: str, split_by: str) -> str:
        for i in block:
            if identifier in i:
                return i.split(split_by)[-1]
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

            address = "\n".join(block[:first_field_index])
            block = block[first_field_index:]
            state = self.get_element_from_block(block, 'State', ':')
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
            return dict(
                address=address,
                pan=pan,
                gstn=gstn,
                state=state,
                contact_person=contact_person,
                email=email,
                tel=tel
            )

        address_table = [list(i) for i in zip(*address_table)]
        billing_address = {'bad'+k:v for k, v in parse_address_block(address_table[0]).items()}
        delivery_address = {'dad' + k: v for k, v in parse_address_block(address_table[0]).items()}
        billing_address.update(delivery_address)
        return billing_address
    def parse_sales_data(self, sales_table):
        header = sales_table.pop(0)
        result = {}
        while True:
            sr_no = sales_table[0][0]
            if not [i for i in sr_no if i.isdigit()]: break
            result.update({'desc_'+sr_no:sales_table[0][1]})
            result.update({'qty_'+sr_no:sales_table[0][2]})
            result.update({'unit_price_'+sr_no:sales_table[0][3]})
            result.update({'total_price_'+sr_no:sales_table[0][1]})
            sales_table.pop(0)
        for i in sales_table:
            for j in i:
                if 'CGST' in j:
                    result.update(cgst_percentage="".join([k for k in j if k.isdigit()]))
                elif 'SGST' in j:
                    result.update(sgst_percentage="".join([k for k in j if k.isdigit()]))
                elif 'IGST' in j:
                    result.update(igst_percentage="".join([k for k in j if k.isdigit()]))
        return result



if __name__ == '__main__':
    docx_folder_name = 'Docxs'
    for path in listdir():
        if '.doc' in path:
            OPF(abspath(path)).seperate_doc()
    os.chdir(abspath(docx_folder_name))
    result_dict_list = []
    opf = OPF('OPF-March\'18.docx').extract_data()
    for i, docx_path in enumerate(listdir()):
        print(docx_path)
        opf = OPF(docx_path)
        result_dict_list.append(opf.extract_data())
    df = DF(result_dict_list)
    df = df['state']
    df.to_excel('final_output.xlsx')


