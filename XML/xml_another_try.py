import csv
import re
import xml.etree.ElementTree
from itertools import groupby
from os import listdir
from pandas import DataFrame

def get_node_value(e, l):
    return eval('e.getroot(){}.text'.format(''.join(['[{}]'.format(i) for i in l])))


res = []


def recursive_iterate(root, history, init):
    if init:
        global res
        res = []
    if root is not None:
        if root.text is not None:
            res.append(history)
        a = list(root)
        for i in range(len(a)):
            recursive_iterate(root[i], '{}, {}'.format(history, i), False)
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
        temp = 8
        for i in res1212:
            if 'billing' in i[1].lower() and 'address' in i[1].lower():
                temp = i[0][0]
        table_2 = [[i[0][1:-1], i[1]] for i in res1212 if i and i[0] and i[0][0] == temp]
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

    def hard_code_field_find(block, field):
        for i, b in enumerate(block):
            if field.lower() in b.lower():
                return i

    def without_state_parse(block, dict_prefix):
        block = list(block)
        which_address = block.pop(0)
        fields = ['contact', 'email', 'gst']
        indexes = [get_index_by_substr([[-1, i] for i in block], f) for f in fields]
        min_index = min([i for i in indexes if i is not None])
        address = ', '.join(block[:min_index])
        block = block[min_index:]
        contact_person_idx = hard_code_field_find(block, 'person')
        return dict(address=address)

    def parse_address_block(block, dict_prefix):
        try:
            state_idx = -1
            for i, ele in enumerate(block):
                if 'state' in ele.lower():
                    state_idx = i
            address = '; '.join(block[1:state_idx])
            state = re.split('state\s:', block[state_idx], flags=re.IGNORECASE)[1]
            contact_person = re_replace('contact person', '', block[state_idx + 1]).replace(':', '').strip()
            tel = re_replace('tel', '', block[state_idx + 2]).replace('#', '')
            email = re_replace('email', '', block[state_idx + 3]).strip("#").strip(":-")
            gstn, pan = block[state_idx + 4].split(', ')
            gstn = re_replace('gst', '', gstn).strip().strip('N').strip('n').strip('NO').strip('no').strip().strip(
                ":").strip().replace('NO', '').replace(':', '')
            pan = re_replace('pan', '', pan).strip().strip('N').strip('n').strip('NO').strip('no').strip().strip(
                ":").strip().replace('NO', '').replace(':', '').strip('=-')
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
        except IndexError:
            return without_state_parse(block, dict_prefix)

    address_table = create_address_table(res1212)
    return dict(
        billing_address="\n".join(address_table[0][1:]),
        delivery_address="\n".join(address_table[1][1:])
    )


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


def get_field(all_fields, identifier, split_by):
    for field in all_fields:
        if identifier in ' '.join([i for i in field[1].split(' ') if i]).lower():
            return field[1].split(split_by)[-1].strip()


def get_max_difference_idx(row):
    diffs = []
    row = sorted(row, key=lambda x: x[0][1])
    for i in range(len(row) - 1):
        diffs.append(row[i + 1][0][1] - row[i][0][1])
    return diffs.index(max(diffs))


def get_closest(block, field, split_by):
    for i, ele in enumerate(block):
        if field in " ".join([i for i in ele.lower().split() if i]):
            block.pop(i)
            return block, ele.split(split_by)[-1]
    return block, None


def get_loose_data(res):
    for i, ele in enumerate(res):
        if len(ele[0]) == 3:
            break
    block = []
    a = res[i]
    while len(a[0]) == 3:
        block.append(a)
        a = res.pop(i + 1)
    block = [i for i in block if i[1]]
    pairs = []
    for k, v in groupby(block, key=lambda x: x[0][0]):
        group = []
        for i in v:
            group.append([i[0][:-1], i[1]])
        group_0_1 = sorted([i[0][1] for i in group])
        for i in range(max(group, key=lambda x: x[0][1])[0][1]):
            if i not in group_0_1:
                group.append([[group[0][0][0], i], ''])
        group = sorted(group, key=lambda x: x[0][1])
        string = " ".join([i[1] for i in group])
        n_spaces = 1
        while " " * n_spaces in string:
            n_spaces += 1
        n_spaces = max(2, n_spaces)
        pairs.extend(string.split(" " * (n_spaces - 1)))

    pairs, sales_person = get_closest(pairs, 'sales person', ':')
    pairs, opf_no = get_closest(pairs, 'opf no', '.')
    pairs, customer_name = get_closest(pairs, 'customer name', ':')
    pairs, date = get_closest(pairs, 'date', ':')
    pairs, purch_order_no = get_closest(pairs, 'order no', ':')
    pairs, pot_id = get_closest(pairs, 'pot id', ':')
    return dict(sales_person=sales_person, opf_no=opf_no, customer_name=customer_name, date=date,
                purch_order_no=purch_order_no, pot_id=pot_id)


def write_to_csv(file_name):
    e = xml.etree.ElementTree.parse(file_name)
    wt_elements = []
    for element in e.iter():
        if element.tag.split('}')[-1] == 't':
            wt_elements.append(element)
    res = [[i[0][4:], i[1].strip()] for i in [[i, get_node_value(e, i)] for i in sorted(
        [[int(i) for i in string[1:].split(', ')] for string in recursive_iterate(e.getroot(), '', True)])] if i[1]]
    loose_data = get_loose_data(res.copy())

    res = [i for i in res if i[1]]
    address_details = parse_address_table(res.copy())
    sales_details = parse_sales_table(res.copy())

    all_details = loose_data.copy()
    all_details.update(address_details)
    all_details.update(sales_details.copy())
    return all_details


import traceback
dicts = []
for file_name in listdir():
    try:
        result_dict = write_to_csv(file_name)
        dicts.append(result_dict)
    except Exception as e:
        print('error in', file_name)
        traceback.print_exc()
        print('\n\n\n')
result = DataFrame(dicts)
print(result)
result.to_excel('output.xlsx')