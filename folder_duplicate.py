import pandas as pd
from xlsxwriter.workbook import Workbook

orig_sheet_location = "execution-result.log.csv"
sheet_location = 'temp.csv'
output_csv_location = 'output.csv'
var_all_counts = 'all_opf_count'
var_all_folder_name = 'folder_links_list'
var_opf_links = 'all_opf_links'


def replace_all_commas():
    import codecs
    import csv
    log_path = orig_sheet_location
    lines = []
    with codecs.open(log_path, 'r') as csv_file:
        log_reader = csv.DictReader((l.replace('\0', '') for l in csv_file))
        for line in log_reader:
            lines.append(line)
    keys = []
    values = []
    written = False

    for myOrderedDict in lines:
        for key, value in myOrderedDict.items():
            if not written:
                keys.append(key)
            value = value.replace(',', ':,:')
            values.append(value)

        with open(sheet_location, "w"):
            if not written:
                written = True

    values = [values[i:i + 3] for i in range(0, len(values), 3)]
    with open(sheet_location, "w") as outfile:
        csvwriter = csv.writer(outfile, lineterminator='\n')
        csvwriter.writerow(keys)
        csvwriter.writerows(values)


def list_like_to_list(string):
    string = string.replace(':,:', ',')
    string = string.replace('[', '')
    string = string.replace(']', "")
    string = string.replace('"', "")
    for i in string.split(","):
        yield i.strip()


replace_all_commas()

df = pd.read_csv(sheet_location)
df.columns = df.iloc[0]
df = df.T
df.columns = df.iloc[0]
df.apply(lambda x: x.str.replace(':,:', ','))

all_counts = [int(i) for i in list_like_to_list(df[var_all_counts][2])]
all_folder_names = list(list_like_to_list(df[var_all_folder_name][2]))
all_opf_links = list(list_like_to_list(df[var_opf_links][2]))

to_csv = []
opf_link_count = 0

for i in range(len(all_counts)):
    count_opf = all_counts[i]
    folder_name = all_folder_names[i]
    if count_opf == 0:
        to_csv.append([folder_name, 'unavailable'])
    else:
        for j in range(count_opf):
            to_csv.append([folder_name, all_opf_links[opf_link_count]])
            opf_link_count += 1
to_csv = [['folder name', 'opf link']] + to_csv

workbook = Workbook('output.xlsx')
worksheet = workbook.add_worksheet()
for i, row in enumerate(to_csv):
    for j, col in enumerate(row):
        worksheet.write_string(i, j, col)
workbook.close()
