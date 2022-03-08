#!/usr/bin/env python3

import csv,re,sys
from collections import defaultdict
from xlsxwriter.workbook import Workbook

input_csv_filename = "alle-personen-7maart2022-bewerkt-in-excel.csv"

reader = csv.DictReader(open(input_csv_filename, "r", encoding="utf-8-sig"), dialect='excel', delimiter=";")

reader.fieldnames.append("Soort") # soort adres: ingevuld wanneer adres niet leeg is.

datum_woordenboek = { row["fout"]:row["goed"] for row in csv.DictReader(open("datums-woordenboek.txt"), delimiter='\t') }

ntnis = defaultdict(list)
datums = []

for row in reader:
    code = row["CODE"]

    # indien straatnaam of plaats dan Soort instellen op Adres
    if row["Straatnaam"]!="" or row["Plaats"]!="":
        row["Soort"] = "Adres"

    # foutieve geboortedatums fixen op basis van woordenboek
    row["Geboortedatum"] = row["Geboortedatum"].strip() # verwijder witruimte links en rechts van datum
    if row["Geboortedatum"] in datum_woordenboek:
        row["Geboortedatum"] = datum_woordenboek[row["Geboortedatum"]]

    # vul het veld 'Ouder dan 100 jaar' op basis van geboortejaar
    if not row["Geboortedatum"]:
        row["Ouder dan 100 jaar"] = "Onbekend"
    else:
        geboortejaar = int(row["Geboortedatum"][-4:])
        if geboortejaar==0:
            row["Ouder dan 100 jaar"] = "Onbekend"
        elif not (geboortejaar>1800 and geboortejaar<1980):
            print("Vermoedelijke fout in",code,"bij geboortedatum:",row["Geboortedatum"])
        else: # nu gebruiken we het geboortejaar om het veld 'Ouder dan 100 jaar te vullen'
            row["Ouder dan 100 jaar"] = "Ja" if geboortejaar<1923 else "Nee"

    # maak een lijst van alle unieke datums die voorkomen
    if row["Geboortedatum"] not in datums: # vind alle unieke datums
        datums.append(row["Geboortedatum"])

    # foutieve overlijdensdatums fixen op basis van woordenboek
    row["Overlijdensdatum"] = row["Overlijdensdatum"].strip()
    if row["Overlijdensdatum"] in datum_woordenboek:
        row["Overlijdensdatum"] = datum_woordenboek[row["Overlijdensdatum"]]

    # als overlijdensdatum voldoet aan de regex dan kun je aannemen dat deze voorkomt in een overlijdensbron (ookal is het 00-00-0000).
    if row["Overlijdensdatum"]:
        if not re.findall(r"^\d{2}-\d{2}-\d{4}$", row["Overlijdensdatum"]):
            print("Vermoedelijke fout in",code,"bij overlijdensdatum:",row["Overlijdensdatum"])
        else:
            row["Persoon overleden"] = "Ja"
            if row["Bron overlijden"]=="": # als er al iets staat bij Bron overlijden (bijv CBG) dan niet overschrijven.
                row["Bron overlijden"] = "Overlijdensdatum" 

    # als er in de kolom 'Bron overlijden' het volgende staat 'Ouder dan 100 jaar' 
    # dan deze info verwijderen en onderbrengen in de 'kolom Ouder dan 100 jaar dmv 'Ja'.
    # En 'Persoon overleden leegmaken'
    if row["Bron overlijden"]=="Ouder dan 100 jaar":
        row["Ouder dan 100 jaar"] = "Ja"
        row["Persoon overleden"] = ""
        row["Bron overlijden"] = ""

    # als iemands leeftijd in de oorlog >26 nemen we aan dat deze persoon inmiddels ouder zou zijn dan 100 jaar
    if row["Leeftijd"] and row["Leeftijd"].isdigit():
        leeftijd = int(row["Leeftijd"])
        if leeftijd>26:
            row["Ouder dan 100 jaar"] = "Ja"

    # voeg de regel toe aan de juiste ntni
    ntnis[code].append(row)

########################################

# print een overzicht van alle foutieve datums die niet door het woordenboek zijn hersteld.
datums.sort() 
for datum in datums:
    if datum and not re.findall(r"^\d{2}-\d{2}-\d{4}$", datum):
        print("^"+datum+"$")

########################################

# maak nieuwe spreadsheet(s)
for ntni in ntnis.values():
    firstRow = ntni[0]
    code = firstRow["CODE"]

    # voor nu schrijven we maar 1 ntni weg de rest slaan we over
    if code!="713-9.27": #825.549": #650.50": #1202.216": #292-1.601":
        continue 

    output_xls_filename = f"output/{code}.xlsx"

    # write to excel spreadsheet
    workbook = Workbook(output_xls_filename)
    worksheet = workbook.add_worksheet()

    # write header / fieldnames to top of spreadsheet
    for c, col in enumerate(reader.fieldnames):
        worksheet.write(0, c, col)

    # write all other cells
    for r, row in enumerate(ntni):
        for c, col in enumerate(reader.fieldnames):
            if col in row:
                worksheet.write(r+1, c, row[col])

    workbook.close()

