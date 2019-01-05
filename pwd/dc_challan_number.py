from csv import DictWriter, DictReader
from sys import argv
from datetime import datetime
from os.path import exists

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
        writer = DictWriter(file, fieldnames=list(data[0].keys()))
        writer.writeheader()
        writer.writerows(data)


def increment_col(dictionary, current_year):
    prev_number = dictionary['last_no']
    next_number = str(int(prev_number)+1).zfill(len(prev_number))
    res = dictionary.copy()
    next_challan_no = "{}-{}-{}-{}".format(dictionary['abbrevation'], next_number,current_year, current_year+1)
    res['last_no'] = next_number
    res['dc_challan_no'] = next_challan_no.upper()
    return res


def main():
    if len(argv) < 2:
        print("Cannot process request. Expected exactly 1 field.")
        data = get_data_from_csv(file_name)
        data[0]['res'] = ''
        write_dicts_to_csv(file_name, data)
        exit(0)

    state = argv[1].lower()
    print(state)

    data = get_data_from_csv(file_name)
    index_state = [i['state'] for i in data].index(state)
    data[index_state] = increment_col(data[index_state], current_year)
    data[0]['res'] = data[index_state]['dc_challan_no']
    write_dicts_to_csv(file_name, data)


def populate_csv():
    og_data = [
        dict(state='state',       abbrevation='abbrevation', last_no='last_no', dc_challan_no='dc_challan_no'),
        dict(state='gujrat',      abbrevation='guj',         last_no='054',     dc_challan_no=''),
        dict(state='maharashtra', abbrevation='mh',          last_no='02901',   dc_challan_no=''),
        dict(state='karnataka',   abbrevation='b',           last_no='0434',    dc_challan_no=''),
        dict(state='delhi',       abbrevation='d',           last_no='0148',    dc_challan_no=''),
        dict(state='tamil nadu',  abbrevation='ch',          last_no='0171',    dc_challan_no=''),
        dict(state='tamilnadu',   abbrevation='ch',          last_no='0171',    dc_challan_no=''),
        dict(state='pune',        abbrevation='mh-p',        last_no='353',     dc_challan_no=''),
        dict(state='telangana',   abbrevation='hy',          last_no='0118',    dc_challan_no=''),
        dict(state='west bengal', abbrevation='k',           last_no='0077',    dc_challan_no=''),
        dict(state='westbengal',  abbrevation='k',           last_no='0077',    dc_challan_no=''),
        dict(state='telangana',   abbrevation='hy',          last_no='0118',    dc_challan_no=''),
    ]

    with open(file_name, 'w', newline='') as file:
        writer = DictWriter(file, fieldnames=og_data[0].keys())
        writer.writerows(og_data)


if __name__ == '__main__':
    if not exists(file_name):
        print('FileNotFoundError: "{}" not found, populating hard coded data.'.format(file_name))
        populate_csv()
    main()
