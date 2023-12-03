# Reads in the purchased zip file and creates a county lookup xlsx

# TODO: write sheet to google then have all BOE sheets reference single google sheet.
#  worry about file format change in sheet causing old BOE sheets to not work

import pandas as pd
import numpy as np

def Counties_to_xlsx(zipcsv, fn):
    """
    Returns an xlsx, fn, of unique state/county form purchased county data  https://www.zip-codes.com/.
    And a dictionary keying state/county to countyFilename, countyToPrint, stateMixedCounty
    """
    zipcsv  = 'zip-codes-database-DELUXE-BUSINESS.csv'
    # multi_county_zip_file = 'zip-codes-database-MULTI-COUNTY.csv'
    zip_rows_to_read = 999999  # for testing
    # zip_rows_to_read = 999  # for testing

    main_zip_file = pd.read_csv(zipcsv, nrows=zip_rows_to_read, keep_default_na=False, usecols=['State', 'County', 'CountyMixedCase'])

    main_zip_file.rename(columns={'ZipCode': 'zip', 'State': 'state', 'County': 'county', 'CountyMixedCase':'countymixedcase'}, inplace=True)
    main_zip_file= main_zip_file.loc[main_zip_file['county'].str.strip() != ""]  # military states like AA and AE have no county
    main_zip_file['countyclean'] = main_zip_file['county'].str.upper().str.replace(' ','',regex=False).str.replace('-','',regex=False).str.replace('.','',regex=False)
    main_zip_file['stateCounty'] = main_zip_file['state'] + "-" + main_zip_file['countyclean']
    main_zip_file['countyFilename'] = main_zip_file['countymixedcase'].str.replace(' ','',regex=False).str.replace('-','',regex=False).str.replace('.','',regex=False)
    main_zip_file['countyToPrint'] = np.where( main_zip_file['countymixedcase'].str[-4:] != "City",main_zip_file['countymixedcase'] + " County", main_zip_file['countymixedcase'] )
    main_zip_file['stateCountyMixed'] = main_zip_file['state'] + "-" + main_zip_file['countyFilename']

    # unique_county = main_zip_file.drop_duplicates(subset = ['state', 'countyFilename'], keep = 'last')
    unique_county = main_zip_file.drop_duplicates(subset = ['stateCounty'], keep = 'last')
    sorted_unique_county = unique_county.sort_values(['stateCounty'], ascending=[True])

    sorted_unique_county.to_excel(fn, index=False, columns=['stateCounty', 'countyFilename', 'countyToPrint', 'stateCountyMixed' ]) # directory defaults to current python work

    # convert df to list so we can set up dictionary tuples easily
    dfList = unique_county[['stateCounty', 'countyFilename', 'countyToPrint', 'stateCountyMixed']].values.tolist()

    # dictList = [ (stateCounty,[countyFilename, countyToPrint, stateMixedCounty] ) for row in unique_county]
    # dictList = [ (row[0],[row[1], row[2], row[3]] ) for row in dfList]
    # set up list of tuples
    dictList = [ (stateCounty,[countyFilename, countyToPrint, stateCountyMixed] ) for stateCounty, countyFilename, countyToPrint, stateCountyMixed in dfList]

    return dict(dictList)

if __name__ == "__main__":

    countydict = Counties_to_xlsx('zip-codes-database-DELUXE-BUSINESS.csv', 'Unique_County_List.xlsx')
    a=1

