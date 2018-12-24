import xml.etree.ElementTree
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
    if identifier != 'Galaxy Billing from (Location)':
        print = lambda *x:None
    else:
        print = __import__('builtins').print

    start_idx = idx = wt_elements.index(max_match(identifier))
    min_identifier = "".join(identifier.split())
    heap_identifier = ""
    while heap_identifier+wt_elements[idx].text.strip() != wt_elements and len(heap_identifier)<len(min_identifier):
        heap_identifier += wt_elements[idx].text
        idx += 1
    for i in range(idx-start_idx):
        print('popping ', wt_elements.pop(start_idx).text)
    probable_field = heap_identifier.split(min_identifier)[-1].strip()
    print(probable_field)
    if heap_identifier.split(min_identifier)[-1].strip().replace(',', '').replace(':', '').replace('.', '') and (set(heap_identifier.replace(' ', ''))-set(min_identifier.replace(' ', ''))):
        print((set(heap_identifier)-set(min_identifier)))
        print('returning from else')
        return probable_field
    else:
        while not wt_elements[start_idx].text.replace(split_by, "").strip():
            wt_elements.pop(start_idx)
        res = wt_elements.pop(start_idx).text
        return res.replace(split_by, '').strip()


def get_element(identifier, split_by):
    if identifier != 'Galaxy Billing from (Location)':
        print = lambda *x:None
    else:
        print = __import__('builtins').print

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

billing = []
delivery = []
j = 0
for i, ele in enumerate(wt_elements):
    if 'GSTN NO' in ele.text:
        print(ele.text, 'breaking')
        break
    else:
        if j%2==0:
            billing.append(wt_elements.pop(0))
            print('b', billing[-1].text)
        else:
            delivery.append(wt_elements.pop(0))
            print('d', delivery[-1].text)
        j += 1

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
    a[1] = '::'.join([i[1] for i in lt])
    return a


def merge_similar_fields(ele):
    i = 0
    temp = [ele[i]]
    eleres = []
    while i<len(ele)-1:
        if len(ele[i][0]) == len(ele[i+1][0]) and hamming_diff_list(ele[i][0], ele[i+1][0]) == 1 and absfrom cv(ele[i+1][0][-2]-ele[i][0][-2]) == 1:
            temp.append(ele[i+1])
        else:
            eleres.append(temp)
            temp = [ele[i+1]]
        i += 1
    for i in range(len(eleres)):
        eleres[i] = merge_lists(eleres[i])
    return eleres


res = recursive_iterate(e.getroot(), '')
res =[[int(i) for i in string[1:].split(', ')] for string in res]
res = sorted(res)

res1 = [[i, get_node_value(e, i)] for i in res]

res2 = merge_similar_fields(res1)
for i, ele in enumerate(res2):
    if len(ele) == 1:
        res2[i] = res2[i][0]
for i in res2:
    print('#', i)