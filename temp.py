import codecs
import csv
log_path = 'execution-result.log.csv'
lines = []
with codecs.open(log_path, 'r') as csv_file:
    log_reader = csv.DictReader((l.replace('\0', '') for l in csv_file))
    for line in log_reader:
        lines.append(line)
keys = []
values = []
rows = []
written = False

for myOrderedDict in lines:
    for key, value in myOrderedDict.items():
        if not written:
            keys.append(key)
        value = value.replace(',', ':,:')
        values.append(value)

    with open("frequencies.csv", "w") as outfile:
        csvwriter = csv.writer(outfile, lineterminator='\n')
        if not written:
            written = True

values = [values[i:i+3] for i in range(0, len(values), 3)]
with open("frequencies.csv", "w") as outfile:
        csvwriter = csv.writer(outfile, lineterminator='\n')
        csvwriter.writerow(keys)
        csvwriter.writerows(values)
