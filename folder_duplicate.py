import pandas as pd
import csv

sheet_location = "execution-result.log.csv"
output_csv_location = 'output.csv'
var_all_counts ='all_opf_count'
var_all_folder_name = 'folder_links_list'
var_opf_links = 'all_opf_links'

def list_like_to_list(string):
    string = string.replace('[', '')
    string = string.replace(']', "")
    string = string.replace('"', "")
    for i in string.split(","):
        yield i.strip()

df = pd.read_csv(sheet_location)
df.columns = df.iloc[0]
df = df.T
df.columns = df.iloc[0]

print(list(df))
print(df[var_all_counts])

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
try:
    with open(output_csv_location, 'w', newline='') as op_file:
        writer = csv.writer(op_file)
        writer.writerows(to_csv)
    print("File write successful")
except Exception as a:
    print(a)
    print("file  write unsuccessful")
