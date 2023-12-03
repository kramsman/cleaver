
#

import numpy as np
import pandas as pd

def create_ZipToCounty_dict():
    """
    Returns a dictionary which uses a tuple of (state,county,zip) to return county; flag key errors to identify bad zp/county combinations.

    Data purchased from https://www.zip-codes.com/.
    Multi county and unique files are merged, multi county does not contain unique zips.
    Unique contains multiple records for the same county for multiple cities in zip.
    """
    main_zip_file  = 'zip-codes-database-DELUXE-BUSINESS.csv'
    multi_county_zip_file = 'zip-codes-database-MULTI-COUNTY.csv'
    zip_rows_to_read = 999999  # for testing
    # xl_rows_to_read_limit = 100

    main_zip_temp = pd.read_csv(main_zip_file, nrows=zip_rows_to_read, keep_default_na=False, usecols=['State', 'County', 'ZipCode'])
    # main_zip_temp.rename(columns={'ZipCode': 'zip', 'State': 'state', 'County': 'county'}, inplace=True) # keep in case field names change between files

    multi_county_temp = pd.read_csv(multi_county_zip_file, nrows=zip_rows_to_read, keep_default_na=False, usecols=['State', 'County', 'ZipCode'])
    # multi_county_temp.rename(columns={'ZipCode': 'zip', 'State': 'state', 'County': 'county'}, inplace=True) # keep in case field names change between files

    combined_temp =  main_zip_temp.append(multi_county_temp, ignore_index=True)
    combined_temp.rename(columns={'ZipCode': 'zip', 'State': 'state', 'County': 'county'}, inplace=True)
    combined_temp['countyclean'] = combined_temp['county'].str.lower().str.replace(' ','').str.replace('-','')

    unique_zip = combined_temp.drop_duplicates(
      subset = ['state', 'countyclean', 'zip'],
      keep = 'last')

    unique_zip.sort_values(['zip', 'state', 'countyclean'], ascending=[True, True, True], inplace=True)

    lstForDict = [[tuple(  [state, countyclean, zip]  ), countyclean ] for zip, state, county, countyclean in unique_zip.values.tolist()]
    return dict(lstForDict)

if __name__ == '__main__':
    dictZipToCounty = create_ZipToCounty_dict()
    print("dictZipToCounty.get(('GA', 'clayton', 30288),'mismatch'", dictZipToCounty.get(('GA', 'clayton', 30288),'mismatch') )
    print("dictZipToCounty.get(('GA', 'dekalb', 30288),'mismatch'", dictZipToCounty.get(('GA', 'dekalb', 30288),'mismatch') )
    print("dictZipToCounty.get(('GA', 'fulton', 30288),'mismatch'", dictZipToCounty.get(('GA', 'fulton', 30288),'mismatch') )

    # c2 = dictZipToCounty.get(('GA', 'dekalb', 30288),'mismatch')
    # c3 = dictZipToCounty.get(('GA', 'fulton', 30288),'mismatch')

    a=1