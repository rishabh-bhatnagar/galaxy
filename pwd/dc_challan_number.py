from csv import DictWriter, DictReader
from sys import argv
from datetime import datetime

current_year = __import__('datetime').datetime.now().year
april_date = datetime(current_year, 4, 1)
current_date= datetime.now()
if current_date<april_date:
    current_year = current_year-1
current_year = int(str(current_year)[-2:])

file_name = 'challan_number.csv'

def get_data_from_csv(file_name):
    with open(file_name, 'r') as file:
        reader = DictReader(file)
        data = [dict(i) for i in reader]
        return data
def write_dicts_to_csv(file_name, data):
    # additional res header of column is appended.
    with open(file_name, 'w', newline='') as file:
        writer = DictWriter(file, fieldnames=list(data[0].keys()) + ['res'])
        writer.writeheader()
        writer.writerows(data)


def increment_col(dictionary, current_year):
    prev_number = dictionary['last_no']
    next_number = str(int(prev_number)+1).zfill(len(prev_number))
    res = dictionary.copy()
    next_challan_no = "{}-{}-{}-{}".format(dictionary['abbrevation'], next_number,current_year, current_year+1)
    res['last_no'] = next_number
    res['dc_challan_no'] = next_challan_no
    return res

if len(argv) < 2:
    print("Cannot process request. Expected exactly 1 field.")
    data = get_data_from_csv(file_name)
    data[0]['res'] = ''
    write_dicts_to_csv(file_name, data)
    exit(0)

state = argv[1].lower()

data = get_data_from_csv(file_name)
index_state = [i['state'] for i in data].index(state)
data[index_state] = increment_col(data[index_state], current_year)
data[0]['res'] = data[index_state]['dc_challan_no']
write_dicts_to_csv(file_name, data)
for i in data:
    print(i)


'''
og_data = [
    dict(state='state', abbrevation='abbrevation', last_no='last_no', dc_challan_no='dc_challan_no'),
    dict(state='telangana', abbrevation='hy', last_no='2', dc_challan_no=''),
    dict(state='maharashtra', abbrevation='mh', last_no='1', dc_challan_no=''),
    dict(state='chennai', abbrevation='mh', last_no='1', dc_challan_no='')
]


with open(file_name, 'w', newline='') as file:
    writer = DictWriter(file, fieldnames=og_data[0].keys())
    writer.writerows(og_data)
'''