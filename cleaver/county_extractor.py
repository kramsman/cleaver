#!/usr/bin/env python3
"""Extract county names from manually-entered filenames.

Method: tokenize + longest contiguous-window exact match.
  - The filename is split into uppercase word tokens on any run of
    non-alphanumeric characters (space, -, (, ), _, etc.).
  - Every contiguous token window is compared against the county names
    for one state (loaded from the master unique_counties.csv).
  - The longest match at a position wins, so "ROANOKE CITY" beats
    "ROANOKE" when both are present.
  - Exactly one distinct match -> that county (guaranteed to be in the
    master list). Zero matches or two+ distinct matches -> error.

Reusable API (import from other programs, e.g. rovcleaver):

    from county_extractor import load_counties, extract_county

    counties = load_counties("VA")               # once, up front
    county, error = extract_county("2026 General Roanoke City.csv", counties)
    if error:
        ...  # error is "NO_MATCH" or "AMBIGUOUS: X / Y"
    else:
        ...  # county == "ROANOKE CITY"

    # Batch: log one line per file + summary, write a CSV report, and get
    # back ({filename: county_or_None}, error_count):
    results, errors = extract_counties_with_report(
        filenames, counties, report_file, log=logger.info)

Standalone: reads filenames from the FileList tab of Test_setup.xlsx
(column A, starting row 4), prints one line per file, and writes
county_extraction_report.csv next to the xlsx.

    python3 county_extractor.py VA
"""

import csv
import os
import re
import sys
from pathlib import Path

# Default locations
_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_COUNTY_CSV = Path("/Users/Denise/Dropbox/Postcard Files/InputFiles/unique_counties.csv")
DEFAULT_XLSX = _ROOT / "Work" / "Test_setup.xlsx"
FILELIST_SHEET = "FileList"
FILELIST_FIRST_ROW = 4


def tokenize(text):
    """Split text into uppercase word tokens on any non-alphanumeric run."""
    return [t for t in re.split(r"[^A-Za-z0-9]+", text.upper()) if t]


def load_counties(state, county_csv=DEFAULT_COUNTY_CSV):
    """Return the list of county names (uppercase) for one state.

    Reads the master list (State,County columns). Raises ValueError if the
    state has no counties in the file.
    """
    state = state.strip().upper()
    counties = []
    with open(county_csv, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            if row["State"].strip().upper() == state:
                counties.append(row["County"].strip().upper())
    if not counties:
        raise ValueError(f"No counties found for state {state!r} in {county_csv}")
    return counties


def extract_county(filename, counties):
    """Extract the county name from one filename.

    counties: iterable of county names for the target state (any case).

    Returns (county, error):
      - ("ROANOKE CITY", None)            exactly one county found
      - (None, "NO_MATCH")                no county name in the filename
      - (None, "AMBIGUOUS: ROANOKE / FRANKLIN")  two+ distinct counties
    """
    stem = os.path.splitext(str(filename))[0]
    tokens = tokenize(stem)

    county_tuples = {c.strip().upper(): tuple(tokenize(c)) for c in counties}

    # Collect every (start, end, county) window match in the filename.
    matches = []
    for county, ctoks in county_tuples.items():
        n = len(ctoks)
        for start in range(len(tokens) - n + 1):
            if tuple(tokens[start:start + n]) == ctoks:
                matches.append((start, start + n, county))

    # Longest match wins: drop any match fully contained in a longer one
    # (e.g. ROANOKE inside ROANOKE CITY).
    surviving = [
        (s, e, county)
        for (s, e, county) in matches
        if not any(
            (s2 <= s and e <= e2) and (e2 - s2) > (e - s)
            for (s2, e2, _) in matches
        )
    ]

    found = sorted({county for (_, _, county) in surviving})
    if len(found) == 1:
        return found[0], None
    if not found:
        return None, "NO_MATCH"
    return None, "AMBIGUOUS: " + " / ".join(found)


def extract_counties_with_report(filenames, counties, report_file, log=print):
    """Extract a county from every filename, log each result, and write a CSV report.

    filenames:   iterable of filenames to process.
    counties:    iterable of county names for the target state (see load_counties).
    report_file: path of the CSV report to write (overwritten). Columns:
                 Status, Filename, County_or_Error; a summary row at the end.
    log:         callable for one line per file plus the summary
                 (print for CLI use, logger.info from other programs).

    Returns (results, error_count) where results is {filename: county_or_None}.
    """
    results = {}
    rows = []
    errors = 0
    for filename in filenames:
        county, error = extract_county(filename, counties)
        results[filename] = county
        if error:
            errors += 1
            log(f"ERROR  {filename}  ->  {error}")
            rows.append(("ERROR", filename, error))
        else:
            log(f"OK     {filename}  ->  {county}")
            rows.append(("OK", filename, county))

    summary = (f"{len(results)} files, {len(results) - errors} matched, "
               f"{errors} error(s)")
    log(summary)

    with open(report_file, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Status", "Filename", "County_or_Error"])
        writer.writerows(rows)
        writer.writerow([])
        writer.writerow([summary])

    return results, errors


def read_filelist(xlsx_path=DEFAULT_XLSX, sheet=FILELIST_SHEET,
                  first_row=FILELIST_FIRST_ROW):
    """Read filenames from column A of the FileList tab, stopping at the
    first blank cell."""
    import openpyxl

    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb[sheet]
    filenames = []
    for (value,) in ws.iter_rows(min_row=first_row, max_col=1, values_only=True):
        if value is None or str(value).strip() == "":
            break
        filenames.append(str(value).strip())
    wb.close()
    return filenames


def main(argv):
    if len(argv) >= 2:
        state = argv[1]
    else:
        state = input("State (e.g. VA): ").strip()
        if not state:
            print(f"usage: {os.path.basename(argv[0])} STATE [xlsx_path]")
            return 2
    xlsx_path = argv[2] if len(argv) > 2 else DEFAULT_XLSX

    counties = load_counties(state)
    filenames = read_filelist(xlsx_path)

    report_file = Path(xlsx_path).parent / "county_extraction_report.csv"
    _, errors = extract_counties_with_report(filenames, counties, report_file)
    print(f"Report written to {report_file}")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
