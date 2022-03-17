#!/usr/bin/env python3

import csv,re,sys,datetime
from collections import defaultdict
from xlsxwriter.workbook import Workbook

input_csv_filename = "data/resultaat-van-stap1.csv"

matching_candidates_file = open("data/matching_candidates.txt","w")

reader = csv.DictReader(open(input_csv_filename, "r", encoding="utf-8-sig"), dialect='excel', delimiter=";")

reader.fieldnames.append("Soort") # soort adres: ingevuld wanneer adres niet leeg is.

datum_woordenboek = { row["fout"]:row["goed"] for row in csv.DictReader(open("data/datums-woordenboek.txt"), delimiter='\t') }

# deze lijst met matches met Oorlogsbronnen is/wordt in stap 4 gemaakt. En kan daarna weer in stap 3 gebruikt worden
NOB_matches = { (row[0]+"_"+row[1]):row[2] for row in csv.reader(open("data/NOB_matches.txt"), delimiter='\t') }

# addressen lijst - lookup table
adressen_lijst = { row["PERSOON_ID"]:row for row in csv.DictReader(open("data/resultaat-van-stap5-adressen.csv"), delimiter=';') }

all_rows = []
ntnis = defaultdict(list)
datums = []

for row in reader:
    code = row["CODE"]

    # indien straatnaam of plaats dan Soort instellen op Adres
    if row["Straatnaam"]!="" or row["Plaats"]!="":
        row["Soort"] = "Adres"

    # verwijder witruimte rondom bij datums
    row["Geboortedatum"] = row["Geboortedatum"].strip() # verwijder witruimte links en rechts van datum

    # probeer datums zonder nullen zoals 1-5-1900 te schrijven als 01-05-1900
    if row["Geboortedatum"]:
        try:
            datum = row["Geboortedatum"]
            row["Geboortedatum"] = datetime.datetime.strptime(datum, '%d-%m-%Y').strftime('%d-%m-%Y')
            # print("datum gefixt",datum,"naar",row["Geboortedatum"])
        except ValueError:
            # print("invalid date",datum)
            pass

    # foutieve geboortedatums fixen op basis van woordenboek    
    if row["Geboortedatum"] in datum_woordenboek:
        row["Geboortedatum"] = datum_woordenboek[row["Geboortedatum"]]

    # vul het veld 'Ouder dan 100 jaar' op basis van geboortejaar
    if not row["Geboortedatum"]:
        row["Ouder dan 100 jaar"] = "Onbekend"
    else:
        geboortejaar = 0
        try:
            geboortejaar = int(row["Geboortedatum"][-4:])
        except:
            pass
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

    # maak een lijst van namen+geboortedatum die mogelijk te matchen is via NOB
    # indien stap 4 al is uitgevoerd wordt ook een evt gevonden match in het veld Externe Identifier bewaard
    if row["Geboortedatum"]:
        try:
            # schrijf naar matching candidates bestand zodat in stap 4 gematched kan worden met NOB
            date = row["Geboortedatum"]
            isodate = datetime.datetime.strptime(date, '%d-%m-%Y').strftime('%Y-%m-%d')
            print(isodate + "\t" + row["Achternaam"], file=matching_candidates_file)

            # wanneer stap 4 (matching met NOB) al is uitgevoerd en gecached 
            # kijk dan of er een match is voor deze persoon
            key = isodate+"_"+row["Achternaam"]
            if key in NOB_matches:
                NOB_url = NOB_matches[key]
                
                # in het geval van een match de url van Netwerk Oorlogsbronnen invullen in het veld Externe Identifier
                # maar voorkom dat handmatig ingevulde of verbeterde Externe Identifier's worden overschreven.
                if not row["Externe Identifier"]:
                    row["Externe Identifier"] = NOB_url

                # als Bron overlijden nog niet is ingevuld en we hebben een match dan Bron overlijden instellen
                if row["Bron overlijden"]=="":
                    row["Bron overlijden"] = "Netwerk Oorlogsbronnen"

        except ValueError:
            pass # skip invalid/incomplete dates

    #############################################
    # Overslaan in uitvoer?

    # alleen wanneer er nog niks (al dan niet handmatig) is ingevuld dan willen we de waarde berekenen
    if row["Overslaan in uitvoer"]=="":

        # als de persoon ouder is dan 100 jaar of er is een Bron overlijden dan willen we deze níet overslaan
        if row["Ouder dan 100 jaar"]=="Ja" or row["Bron overlijden"]!="":
            row["Overslaan in uitvoer"] = "Nee"
        else:
            # in dit geval is de persoon misschien jonger dan 100 (want Nee of Onbekend)
            # en/of de Bron van overlijden is niet bekend (Bron overlijden kan zijn: CBG, NOB, Overlijdensdatum, Burgelijke Stand etc.)
            row["Overslaan in uitvoer"] = "Ja"            


    # voeg overal het trefwoord Tweede Wereldoorlog toe
    row["Trefwoord (tmp)"] = "Tweede Wereldoorlog"

    ###############################################

    # uitzondering: van 650.101 mag wél de meta-data online (maar niet de scan)
    if code=="650.101":
        row["Overslaan in uitvoer"]="Nee"
        # we willen per uitzondering "Persoon overleden" niet tonen op de website bij 650.101
        row["Persoon overleden"] = ""

    ###############################################

    # uitzondering: van 650.102-* mag wél de meta-data online
    if code.find("650.102-")>-1:
        row["Overslaan in uitvoer"]="Nee"

    ###############################################

    # uit het resultaat van stap 5 (adressen) nu de kolommen met adres informatie vullen indien leeg...
    # op basis van Persoon ID!
    if row["Straatnaam"]=="":
        # print("Dan opzoek naar adres in resultaat-van-stap5-adressen.csv",row["ID"])

        if row["ID"] in adressen_lijst:
            adres = adressen_lijst[row["ID"]] # adressen_lijst is per PERSOON_ID

            row["Straatnaam"] = adres["Straat (tijdelijk)"]
            row["Huisnummer"] = adres["Huisnummer(s)"]
            row["Huisnummer toev."] = adres["Huisnummer toev."]
            row["Plaats"] = adres["Plaats (tijdelijk)"]
            row["Soort"] = "Adres"

    ##############################################

    # veld 'Scan Zichtbaar' op 'nee' zetten indien: bij veld Persoon overleden 'Nee' of 'Onbekend' staat
    if row['Persoon overleden'] in ["Nee", "Onbekend"]:
        row["Scan Zichtbaar"] = "Nee"
        row["Opmerking bij scan"] = "<a href='https://hetutrechtsarchief.nl'>test</a>"

    ##############################################

    # voeg de regel toe aan de juiste ntni
    all_rows.append(row)
    ntnis[code].append(row)

########################################

# close matching_candidates file
matching_candidates_file.close()


########################################

# print een overzicht van alle foutieve datums die niet door het woordenboek zijn hersteld.
datums.sort() 
for datum in datums:
    if datum and not re.findall(r"^\d{2}-\d{2}-\d{4}$", datum):
        print("foutieve datum, niet hersteld door woordenboek: ^"+datum+"$", "inv="+code)

########################################

# maak een nieuwe versie van de grote sheet
workbook = Workbook("data/all-rows.xlsx")
worksheet = workbook.add_worksheet()

# write header / fieldnames to top of spreadsheet
for c, col in enumerate(reader.fieldnames):
    worksheet.write(0, c, col) 

# write all rows
for r, row in enumerate(all_rows):
    for c, col in enumerate(reader.fieldnames):
        if col in row:
            worksheet.write(r+1, c, row[col])

workbook.close()

########################################

# maak losse nieuwe spreadsheet(s) per ntni
for ntni in ntnis.values():
    firstRow = ntni[0]
    code = firstRow["CODE"]

    # voor nu schrijven we maar 1 ntni weg de rest slaan we over
    # if code!="650.102-a": #713-9.27": #825.549": #650.50": #1202.216": #292-1.601":
    #     continue 

    output_xls_filename = f"data/naar-mais/{code}.xlsx"

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

