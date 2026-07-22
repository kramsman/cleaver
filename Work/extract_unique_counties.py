#!/usr/bin/env python3
"""Extract unique State/County pairs from the zip code database.

Reads the master zip code CSV and writes a sorted, de-duplicated list of
State,County pairs to unique_counties.csv (the master county list).
"""

import csv

INPUT_FILE = "/Users/Denise/Library/CloudStorage/Dropbox/Postcard Files/ROVPrograms/ROVCleaver_Production/zip-codes-database-DELUXE-BUSINESS.csv"
OUTPUT_FILE = "/Users/Denise/Dropbox/Postcard Files/InputFiles/unique_counties.csv"


def main():
    unique = set()

    with open(INPUT_FILE, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            state = row["State"].strip()
            county = row["County"].strip()
            if state and county:
                unique.add((state, county))

    with open(OUTPUT_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["State", "County"])
        for state, county in sorted(unique):
            writer.writerow([state, county])

    print(f"Wrote {len(unique)} unique state/county pairs to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()