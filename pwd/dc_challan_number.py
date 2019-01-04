from csv import DictWriter, DictReader
from sys import argv
current_year = int(str(__import__('datetime').datetime.now().year)[-2:])


if len(argv) < 2:
    print("Cannot process request. Expected exactly 1 field.")
    exit(0)
state = argv[1].lower()

file_name = 'challan_number.csv'

og_data = [
    dict(state='state', abbrevation='abbrevation', last_no='last_no'),

    dict(state='telangana', abbrevation='hy', last_no='2'),
dict(state='maharashtra', abbrevation='mh', last_no='1'),
    dict(state='chennai', abbrevation='mh', last_no='1')
]


with open(file_name, 'w') as file:
    writer = DictWriter(file, fieldnames=og_data[0].keys())
    writer.writerows(og_data)
with open(file_name, 'r') as file:
    reader = DictReader(file)
    data = [dict(i) for i in reader]

index_state = [i['state'] for i in data].index(state)
print(data[index_state])