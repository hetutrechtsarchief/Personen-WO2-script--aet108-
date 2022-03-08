#!/usr/bin/env python3

import csv,re
from collections import defaultdict
from xlsxwriter.workbook import Workbook

# select ntni.code, ahd.id, ahd.guid, aew.fvd_id, fvd.prompt, aew.waarde, bes.bestandsnaam
# from archiefeenheden ahd
# join archiefeenheid_waarden aew on ahd.id=aew.ahd_id
# join ahd_relaties rel on rel.ahd_id=ahd.id
# join ahd_bestanden bes on rel.ahd_id2=bes.ahd_id
# join archiefeenheden ntni on ahd.ahd_id_top=ntni.id
# join flexvelden fvd on fvd.id=aew.fvd_id
# where ahd.aet_id=108
# and ahd.dt > to_date('01-JAN-21','DD-MON-YY')
# and aew.waarde is not null
# and fvd_id not in (8415,8417)
# order by ahd.dt;

filename = "alle-personen-7maart2022.csv"
output_filename = "output.csv"
output_xls_filename = "alle-personen-7maart2022.xlsx"
fixed = ["ID","GUID","CODE","BESTANDSNAAM"]
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

# output to csv
writer = csv.DictWriter(open(output_filename,"w"), fieldnames=header) #, delimiter=',', quoting=csv.QUOTE_ALL, dialect='excel')
writer.writeheader()
writer.writerows(items.values())

# write to excel spreadsheet
workbook = Workbook(output_xls_filename)
worksheet = workbook.add_worksheet()

for c, col in enumerate(header):
    worksheet.write(0, c, col)

for r, row in enumerate(items.values()):
    for c, col in enumerate(header):
        worksheet.write(r+1, c, row[col] if col in row else None)

workbook.close()

 