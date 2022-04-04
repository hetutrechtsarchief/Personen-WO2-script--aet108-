#!/usr/bin/env python3

import csv,re
from collections import defaultdict
from xlsxwriter.workbook import Workbook

filename = "data/alle-adressen-4-april.csv"
output_csv_filename = "data/resultaat-van-stap5-adressen.csv"
fixed = ["ADRES_ID", "PERSOON_ID"]
flex_key = "PROMPT"
flex_value = "WAARDE"

header = fixed.copy()
items = defaultdict(dict)

for row in csv.DictReader(open(filename)):

    row["WAARDE"] = row["WAARDE"].replace("\n"," ") # replace line breaks by spaces

    # create or get item
    item = items[row["ADRES_ID"]]


    # add fixed fields
    for k,v in row.items():
        if k in fixed:
            item[k] = v

    # add flex fields
    item[row[flex_key]] = row[flex_value]

    # if row["PERSOON_ID"]=='52161163':
    #     print(item)

        
    # update header
    header.append(row[flex_key]) if row[flex_key] not in header else None

# ##########################################

# output to csv
with open(output_csv_filename,"w") as f:
    writer = csv.DictWriter(f, fieldnames=header, delimiter=";")
    writer.writeheader()
    
    writer.writerows(items.values())

