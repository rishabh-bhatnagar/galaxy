import xml.etree.ElementTree
import csv

e = xml.etree.ElementTree.parse('file.xml')
namespaces = {'w':"http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

wt_elements = []

result = {}

for element in e.iter():
    if element.tag.split('}')[-1] == 't':
        wt_elements.append(element)


def max_match(string):
    probable = []
    for idx, ele in enumerate(wt_elements):
        curr_str = wt_elements[idx].text
        if curr_str.startswith(string) or string.startswith(curr_str):
            probable.append(ele)
    return max(probable, key=lambda x:x.text)


def adj_substr(identifier, split_by):
    start_idx = idx = wt_elements.index(max_match(identifier))
    min_identifier = "".join(identifier.split())
    heap_identifier = ""
    while heap_identifier+wt_elements[idx].text.strip() != wt_elements and len(heap_identifier)<len(min_identifier):
        heap_identifier += wt_elements[idx].text
        idx += 1
    for i in range(idx-start_idx):
        wt_elements.pop(start_idx).text
    probable_field = heap_identifier.split(min_identifier)[-1].strip()
    if heap_identifier.split(min_identifier)[-1].strip().replace(',', '').replace(':', '').replace('.', '') and (set(heap_identifier.replace(' ', ''))-set(min_identifier.replace(' ', ''))):
        return probable_field
    else:
        while not wt_elements[start_idx].text.replace(split_by, "").strip():
            wt_elements.pop(start_idx)
        res = wt_elements.pop(start_idx).text
        return res.replace(split_by, '').strip()


def get_element(identifier, split_by):
    tag_elements = [ele.text for ele in wt_elements if identifier in ele.text]
    if not tag_elements:
        return adj_substr(identifier, split_by)

    tag_ele = tag_elements.pop(0)
    res = tag_ele.split(split_by)[-1].strip()
    if not res:
        idx = wt_elements.index([ele for ele in wt_elements if identifier in ele.text][0])
        a=wt_elements.pop(idx)
        probable = wt_elements.pop(idx)
        while not probable.text.strip():
            probable = wt_elements.pop(idx)
        return probable.text.strip()
    else:
        wt_elements.pop(wt_elements.index([ele for ele in wt_elements if ele.text==tag_ele][0]))
        return res


for i, element in enumerate(wt_elements):
    wt_elements[i].text = element.text.strip()

result['sales_person'] = get_element('Sales Person:', ':')
result['opf_no'] = get_element('GOAPL OPF No', '.')
result['opf_date'] = get_element('OPF Date', ':')
result['billing_location'] = get_element('Galaxy Billing from (Location)', ':')
result['customer_name'] = get_element('Customer Name', ':')
result['pon'] = get_element('Purchase Order No', '.')
result['purch_date'] = get_element('Purchase Date', ':')
result['pot_id'] = get_element('POT ID', ':')

a=''
while a != 'delivery address':
    a = " ".join(wt_elements.pop(0).text.strip().split()).lower()

j = 0
for i, ele in enumerate(wt_elements):
    if 'GSTN NO' in ele.text:
        break
res = []


def get_node_value(e, l):
    return eval('e.getroot(){}.text'.format(''.join(['[{}]'.format(i) for i in l])))


def hamming_diff_list(l1, l2):
    return len(set([i-j for i, j in zip(l1, l2)]))-1


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
    a[0][-2] = sum([i[0][-2] for i in lt])/len(lt)
    a[1] = ' '.join([i[1] for i in lt])
    return a


def merge_similar_fields(ele, k=2):
    i = 0
    temp = [ele[i]]
    eleres = []
    while i<len(ele)-1:
        if len(ele[i][0]) == len(ele[i+1][0]) and hamming_diff_list(ele[i][0], ele[i+1][0]) == 1 and abs(ele[i+1][0][-k]-ele[i][0][-k]) == 1:
            temp.append(ele[i+1])
        else:
            eleres.append(temp)
            temp = [ele[i+1]]
        i += 1
    for i in range(len(eleres)):
        eleres[i] = merge_lists(eleres[i])

    for i, ele in enumerate(eleres):
        if len(ele) == 1:
            eleres[i] = eleres[i][0]

    eleres2 = []
    i = 0
    while i < len(eleres)-1:
        a = eleres[i]
        b = eleres[i+1]
        if a[0][:-2] == b[0][:-2] and b[0][-2]-a[0][-2] == 1 and a[0][-1]-b[0][-1] == 1:
            a[0][-1] = (a[0][-1]+b[0][-1])/2
            a[0][-2] = (b[0][-2] + a[0][-2])/2
            a[1] = a[1]+ ''+b[1]
            eleres2.append(a)
            i += 1
        else:
            eleres2.append(a)
        i += 1
    return eleres2


def get_4_similar_2_dont_care(idx, l):
    a = l[idx]
    for i, elem in enumerate(l[idx:]):
        if i != idx:
            if len(a[0]) == len(elem[0]):
                if a[0][:-4] == elem[0][:-4]:
                    if a[0][-1] == elem[0][-1]:
                        if a[0][-3] == elem[0][-3]:
                            return [i for i in l if i not in [a, elem]], a, elem


def split_list_pair(array):
    one = []
    two = []
    rest = []
    array = [[tuple(tuple(i[0])), i[1]] for i in array]
    while array:
        similar_elements = get_4_similar_2_dont_care(0, array)
        if similar_elements is not None:
            array = similar_elements[0]
            one.append(similar_elements[1])
            two.append(similar_elements[2])
        else:
            rest.append(array.pop(0))
    return one, two, rest


def split_by_word(lis, word):
    for i, ele in enumerate(lis):
        if word in ' '.join(ele[1].replace(":", "").replace(',', '').split(" ")).lower():
            return lis[:(i+1)], lis[(i+1):]


def get_field(all_fields, identifier, split_by):
    for field in all_fields:
        if identifier in ' '.join([i for i in field[1].split(' ') if i]).lower():
            return field[1].split(split_by)[-1].strip()


def parse_address_details(block, dict_prefix):
    if 'address' in [i.lower() for i in block[0][1].split(' ')]:
        block.pop(0)
    for j, i in enumerate(block):
        if 'state' in [i.lower() for i in i[1].split(' ')]:
            break

    address = ', '.join([u[1] for u in block[:j]])
    block.pop(0)
    state = get_field(block, 'state', ':')
    contact_person = get_field(block, 'person', ':')
    block.pop(0)
    tel = get_field(block, 'tel', ':')
    block.pop(0)
    email = get_field(block, 'email', ':')
    block.pop(0)
    gstn = get_field(block, 'gstn', ':')
    block.pop(0)
    pan = get_field(block, 'pan', ':')
    block.pop(0)

    return {
        dict_prefix+'address':address,
        dict_prefix+'state':state,
        dict_prefix+'contact_person':contact_person,
        dict_prefix+'tel':tel,
        dict_prefix+'email':email,
        dict_prefix+'gstn':gstn,
        dict_prefix+'pan':pan
    }


def get_index_by_substr(block, substr):
    for i in block:
        if substr in " ".join([u for u in i[1].split(' ') if u]).lower():
            return block.index(i)


def print_table(table_data):
    widths = [max(map(len, col)) for col in zip(*table_data)]
    for row in table_data:
        print("  ".join((val.ljust(width) for val, width in zip(row, widths))))


def create_table(elements):
    row_min = min(elements, key=lambda x:x[0][0])[0][0]
    col_min = min(elements, key=lambda x:x[0][1])[0][1]
    elements = [[i[0]-row_min, i[1]-col_min, string] for i, string in elements]
    row_max = max(elements, key= lambda x: x[0])[0]+1
    col_max = max(elements, key=lambda x: x[1])[1]+1
    table = [['' for i in range(col_max+1)] for j in range(row_max+1)]
    for i in elements:
        table[i[0]][i[1]] = i[2]
    # table = [[str(i) for i in range(col_max)]] + table
    # print_table(table)
    return table


def parse_sales_table(table):
    number_of_products = len([i for i in table if i[0]])-1
    result_dict = {}
    for i in range(1, number_of_products+1):
        row = table[i]
        result_dict['desc'+str(i)] = row[1]
        result_dict['qty' + str(i)] = row[2]
        result_dict['unit_price' + str(i)] = row[2]
        result_dict['total_price' + str(i)] = row[2]
    remaining_table = table[(number_of_products+1):]
    result_dict['sub_total'] = remaining_table[0][4]
    result_dict['cgst'] = remaining_table[1][4]
    result_dict['sgst'] = remaining_table[2][4]
    result_dict['igst'] = remaining_table[3][4]
    result_dict['freight'] = remaining_table[4][4]
    result_dict['grand total'] = remaining_table[1][4]
    return result_dict


res = recursive_iterate(e.getroot(), '')
res = [[int(i) for i in string[1:].split(', ')] for string in res]
res = sorted(res)
res1 = [[i, get_node_value(e, i)] for i in res]

res2 = merge_similar_fields(res1)
res2 = [i for i in res2 if i[1].strip()!='']

loose_fields, rest = split_by_word(res2, 'billing address')
rest = [loose_fields.pop()]+rest
all_fields = [('sales person', ':'), ('id', ':'), ('opf no', "."), ('customer name', ':'), ('galaxy billing from (location)', ':'), ('purchase order no', '.'), ('purchase date', ':')]
final_result = {field[0]:get_field(loose_fields,  *field) for field in all_fields}

one, two, rest = split_list_pair(rest)
one, _ = split_by_word(one, 'pan no')
rest.extend(_)
two, _ = split_by_word(two, 'pan no')
rest.extend(_)

loose_fields = loose_fields
bad = parse_address_details(one, 'bad')
dad = parse_address_details(two, 'dad')
final_result.update(bad)
final_result.update(dad)

rest = [[list(i[0]), i[1]] for i in sorted(rest)]
rest = merge_similar_fields(rest, 3)
i1 = get_index_by_substr(rest, 'sales detail')+1
i2 = get_index_by_substr(rest, 'grand total')+2
current_block = rest[i1:i2]
rest = rest[i2:]

table1 = create_table([[list(i[0][5:7]), i[1]] for i in current_block])
sales_details = parse_sales_table(table1)
final_result.update(sales_details)

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

with open('dict.csv', 'a') as csv_file:
    writer = csv.DictWriter(csv_file, fieldnames=final_result.keys())
    if not header_present:
        writer.writeheader()
        writer.writerow(final_result)
