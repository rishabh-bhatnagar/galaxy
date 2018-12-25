import xml.etree.ElementTree
import re
from itertools import groupby
import csv

e = xml.etree.ElementTree.parse('file.xml')
namespaces = {'w': "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

wt_elements = []

result = {}

for element in e.iter():
    if element.tag.split('}')[-1] == 't':
        wt_elements.append(element)


def get_node_value(e, l):
    return eval('e.getroot(){}.text'.format(''.join(['[{}]'.format(i) for i in l])))


res = []


def recursive_iterate(root, history):
    if root is not None:
        if root.text is not None:
            res.append(history)
        a = list(root)
        for i in range(len(a)):
            recursive_iterate(root[i], '{}, {}'.format(history, i))
    return res


def merge_lists(lt):
    if len(lt) == 1:
        return lt
    a = lt[0]
    a[0][-2] = sum([i[0][-2] for i in lt]) / len(lt)
    a[1] = ' '.join([i[1] for i in lt])
    return a


def re_replace(replacee, replacer, string):
    import re
    insensitive_hippo = re.compile(re.escape(replacee), re.IGNORECASE)
    return insensitive_hippo.sub(replacer, string)


def parse_address_table(res1212):
    def create_address_table(res1212):
        table2 = []
        table_2 = [[i[0][1:-1], i[1]] for i in res1212 if i and i[0] and i[0][0] == 9]
        for k, v in groupby(table_2, lambda x: x[0][:-1]):
            groups = []
            for i in v:
                groups.append(i)
            groups[0][-1] = "".join([i[-1] for i in groups])
            groups[0][0] = groups[0][0][:-1]
            table2.append(groups[0])

        table21 = []
        for k, v in groupby(table2, lambda x: x[0][:-1]):
            groups = []
            for i in v:
                groups.append(i)
            groups[0][-1] = ", ".join([i[-1] for i in groups])
            groups[0][0] = groups[0][0][:-1]
            table21.append(groups[0])
        address_table = create_table(table21)
        address_table = list(zip(*address_table))
        return address_table

    def parse_address_block(block, dict_prefix):
        f1212 = {}
        state_idx = -1
        for i, ele in enumerate(block):
            if 'state' in ele.lower():
                state_idx = i
        address = '; '.join(block[1:state_idx])
        state = re.split('state\s:', block[state_idx], flags=re.IGNORECASE)[1]
        contact_person = re_replace('contact person', '', block[state_idx + 1])
        tel = re_replace('tel', '', block[state_idx + 2])
        email = re_replace('email', '', block[state_idx + 3])
        gstn, pan = block[state_idx + 4].split(', ')
        gstn = re_replace('gstn no', '', gstn)
        pan = re_replace('pan no', '', pan)
        f1212 = {
            dict_prefix + 'address': address,
            dict_prefix + 'state': state,
            dict_prefix + 'contact_person': contact_person,
            dict_prefix + "tel": tel,
            dict_prefix + "email": email,
            dict_prefix + "gstn": gstn,
            dict_prefix + 'pan': pan
        }
        return f1212

    address_table = create_address_table(res1212)
    a = parse_address_block(address_table[0], 'bad')
    b = parse_address_block(address_table[1], 'dad')
    a.update(b)
    return a


def get_index_by_substr(block, substr):
    for i in block:
        if substr in " ".join([u for u in i[1].split(' ') if u]).lower():
            return block.index(i)


def print_table(table_data):
    widths = [max(map(len, col)) for col in zip(*table_data)]
    for row in table_data:
        print("  ".join((val.ljust(width) for val, width in zip(row, widths))))


def create_table(elements):
    row_min = min(elements, key=lambda x: x[0][0])[0][0]
    col_min = min(elements, key=lambda x: x[0][1])[0][1]
    elements = [[i[0] - row_min, i[1] - col_min, string] for i, string in elements]
    row_max = max(elements, key=lambda x: x[0])[0] + 1
    col_max = max(elements, key=lambda x: x[1])[1] + 1
    table = [['' for i in range(col_max + 1)] for j in range(row_max + 1)]
    for i in elements:
        table[i[0]][i[1]] = i[2]
    # table = [[str(i) for i in range(col_max)]] + table
    # print_table(table)
    return table


def parse_sales_table(res):
    def get_sales_table(res):
        i1 = get_index_by_substr(res, 'sales detail') + 1
        i2 = get_index_by_substr(res, 'grand total') + 2
        table9 = [[list(i[0][1:-1]), i[1]] for i in res[i1:i2]]
        table8 = []
        for k, v in groupby(table9, lambda x: x[0][:-1]):
            groups = []
            for i in v:
                groups.append(i)
            groups[0][-1] = "".join([i[-1] for i in groups])
            groups[0][0] = groups[0][0][:-1]
            table8.append(groups[0])
        table7 = []
        for k, v in groupby(table8, lambda x: x[0][:-1]):
            groups = []
            for i in v:
                groups.append(i)
            groups[0][-1] = "; ".join([i[-1] for i in groups])
            groups[0][0] = groups[0][0][:-1]
            table7.append(groups[0])
        return create_table(table7)

    table = get_sales_table(res)
    number_of_products = len([i for i in table if i[0]]) - 1
    result_dict = {}
    for i in range(1, number_of_products + 1):
        row = table[i]
        result_dict['desc' + str(i)] = row[1]
        result_dict['qty' + str(i)] = row[2]
        result_dict['unit_price' + str(i)] = row[2]
        result_dict['total_price' + str(i)] = row[2]
    remaining_table = table[(number_of_products + 1):]
    result_dict['sub_total'] = remaining_table[0][4]
    result_dict['cgst'] = remaining_table[1][4]
    result_dict['sgst'] = remaining_table[2][4]
    result_dict['igst'] = remaining_table[3][4]
    result_dict['freight'] = remaining_table[4][4]
    if 'round' in remaining_table[5][3].lower():
        result_dict['round off'] = remaining_table[5][4]
        result_dict['grand total'] = remaining_table[6][4]
    else:
        result_dict['grand total'] = remaining_table[5][4]
    return result_dict


res = [[i[0][4:], i[1]] for i in [[i, get_node_value(e, i)] for i in sorted([[int(i) for i in string[1:].split(', ')] for string in recursive_iterate(e.getroot(), '')])] if i[1]]

address_details = parse_address_table(res)
sales_details = parse_sales_table(res)
all_details = address_details.copy()
all_details.update(sales_details)

header_present = False
try:
    with open('dict.csv', 'r') as csv_file:
        reader = csv.reader(csv_file)
        for data in reader:
            if len(data)>1:
                header_present = True
                break
except FileNotFoundError:
    header_present = False

with open('dict.csv', 'a', newline='') as csv_file:
    writer = csv.DictWriter(csv_file, fieldnames=all_details.keys())
    if not header_present:
        writer.writeheader()
        writer.writerow(all_details)