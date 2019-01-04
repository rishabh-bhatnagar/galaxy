from csv import DictWriter, DictReader
from sys import argv


if len(argv) < 2:
    print("Cannot process request. Cannot ")
file_name = 'challan_number.csv'
data = []
with open(file_name, 'r') as file:
    reader = DictReader(file)
    data = (i for i in reader)

with open(file_name, 'w') as file:
    wri = 345
