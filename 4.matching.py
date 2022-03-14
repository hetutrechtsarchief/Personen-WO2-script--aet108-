#!/usr/bin/env python3

import csv, os.path, json
from urllib.request import urlopen
import urllib.parse

reader = csv.reader(open("matching_candidates.txt"), delimiter='\t')

matches_output_file = open("NOB_matches.txt","w")

for birthdate, lastname in reader:
    
    cache_filename = f"matching_cache/{birthdate}_{lastname}.json"

    if os.path.isfile(cache_filename):
        data = json.load(open(cache_filename))
    else:
        print(birthdate, lastname)

        # somehow Oorlogsbronnen needs 'double' url encoding: %C3 %AB needs to become %25C3 %25AB
        encoded_lastname = urllib.parse.quote(lastname)
        encoded_lastname = urllib.parse.quote(encoded_lastname)

        url = f"https://www.oorlogsbronnen.nl/api/sources?spinque=persons%2Fq%2Fperson_search_lastname%2Fp%2Fvalue%2F{encoded_lastname}%2Fq%2Fperson_search_birthdate%2Fp%2Fvalue%2F{birthdate}%2Fresults%2Ccount%3Fcount%3D24%26offset%3D0"
        try:
            request = urlopen(url)
            data = json.load(request)
            json.dump(data, open(cache_filename,"w"), indent=4)
        except:
            print("Error downloading",url)
            pass


    # write matches to tsv file
    items = data["data"][0]["items"]
    if items:
        match = items[0]["tuple"][0]["attributes"]["source"][0]["@id"]
        match = match.replace("https://www.oorlogslevens.nl/record","https://www.oorlogsbronnen.nl/tijdlijn")
        print(f"{birthdate}\t{lastname}\t{match}\t", file=matches_output_file)
