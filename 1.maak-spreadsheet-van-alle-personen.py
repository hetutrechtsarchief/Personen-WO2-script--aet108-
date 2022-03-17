#!/usr/bin/env python3

import csv,re
from collections import defaultdict
from xlsxwriter.workbook import Workbook

filename = "data/alle-personen-16-maart.csv"
output_csv_filename = "data/resultaat-van-stap1.csv"
output_xls_filename = "data/resultaat-van-stap1.xlsx"
fixed = ["ID","GUID","CODE", "Bestandsnaam (tmp)"]
flex_key = "PROMPT"
flex_value = "WAARDE"

header = fixed.copy()
items = defaultdict(dict)

for row in csv.DictReader(open(filename)):

    row["WAARDE"] = row["WAARDE"].replace("\n"," ") # replace line breaks by spaces

    # create or get item
    item = items[row["ID"]]

    # add fixed fields
    for k,v in row.items():
        if k in fixed:
            item[k] = v

    # add flex fields
    item[row[flex_key]] = row[flex_value]
        
    # update header
    header.append(row[flex_key]) if row[flex_key] not in header else None

##########################################

# output to csv
writer = csv.DictWriter(open(output_csv_filename,"w"), fieldnames=header, delimiter=";")
writer.writeheader()
writer.writerows(items.values())

##########################################

# write to excel spreadsheet
workbook = Workbook(output_xls_filename)
worksheet = workbook.add_worksheet()

for c, col in enumerate(header):
    worksheet.write(0, c, col)

for r, row in enumerate(items.values()):
    for c, col in enumerate(header):
        worksheet.write(r+1, c, row[col] if col in row else None)

workbook.close()

 