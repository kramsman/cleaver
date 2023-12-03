# test numpy array to file using json dump.  index makes it tough.

import pandas as pd
import numpy as np
import csv


main_zip_file = 'zip-codes-database-DELUXE-BUSINESS.csv'
multi_county_zip_file = 'zip-codes-database-MULTI-COUNTY.csv'


def create_ZipToCountyList_dict(op):
    # BEK2/28/22 copied code in rather than import
    """
    Returns a dictionary which uses zip to return a list of counties

    Data purchased from https://www.zip-codes.com/.
    Multi county and unique files are merged, multi county does not contain unique zips.
    Unique contains multiple records for the same county for multiple cities in zip.
    """
    # main_zip_file  = 'zip-codes-database-DELUXE-BUSINESS.csv'  # put in main body BEK 5/3/22
    # multi_county_zip_file = 'zip-codes-database-MULTI-COUNTY.csv'
    zip_rows_to_read = 999999 # for testing
    # zip_rows_to_read = 999 # for testing


    main_zip_temp = pd.read_csv(main_zip_file, nrows=zip_rows_to_read, keep_default_na=False, usecols=['County', 'ZipCode'])
    multi_county_temp = pd.read_csv(multi_county_zip_file, nrows=zip_rows_to_read, keep_default_na=False, usecols=['County', 'ZipCode'])

    combined_temp =  main_zip_temp.append(multi_county_temp, ignore_index=True)
    combined_temp.rename(columns={'ZipCode': 'zip', 'County': 'county'}, inplace=True)
    combined_temp['countyclean'] = combined_temp['county'].str.lower().str.replace(' ','').str.replace('-','').str.replace('.','')

    combined_temp2 = combined_temp[['zip', 'countyclean']]  # keep only two cols
    unique_zip_county = combined_temp2.drop_duplicates(subset = ['zip', 'countyclean' ], keep = 'last')

    # this is very slow cause loop in loop with many zips - use GROUPBY below
    # lstForDict = [ ( zip  , [ countyclean2 for zip2, countyclean2 in unique_zip_countylist if zip2 == zip ]) for zip in unique_ziplist]

    # seriesForDict = unique_zip_county.groupby(["zip"] )["countyclean"].unique() # produced array not list. MUCH faster than list comprehension
    # seriesForDict = unique_zip_county.groupby(["zip"], as_index=False )["countyclean"].unique() # as_index=False supposed to create variables instead of multindex but doesn't work with unique.
    # https://realpython.com/pandas-groupby/?utm_source=notification_summary&utm_medium=email&utm_campaign=2022-05-17

    # or try this https://stackoverflow.com/questions/22219004/how-to-group-dataframe-rows-into-list-in-pandas-groupby
    # like df.groupby('a').agg({'b': lambda x: list(x)})
    # dfForDict = unique_zip_county.groupby(["zip"], as_index=False ).agg({'countyclean': lambda x: list(x)})  # return "normal" list not array so no need for following
    dfForDict = unique_zip_county.groupby(["zip"], as_index=False ).agg({'countyclean': list})  # return "normal" list not array so no need for following
    listForDict = dfForDict.values.tolist()
    # seriesForDict = unique_zip_county.groupby(["zip"])["countyclean"].unique() # produced array not list. MUCH faster than list comprehension


    # TODO: to write this to a file must convert ndarray to list, but zip is contained in index  https://www.geeksforgeeks.org/write-a-dictionary-to-a-file-in-python/
    # https://www.journaldev.com/32797/python-convert-numpy-array-to-list
    # uses this method  https://stackoverflow.com/questions/58858374/pandas-get-list-of-unique-values-in-column-a-for-each-unique-value-in-column-b

    # convert series objects to lists and bring index into a list
    # listForDict = []
    # for index, countyseries in seriesForDict.items():
    #     listForDict.append([int(index), countyseries.tolist()])

    dictZipToCountyList = dict(listForDict)

    with open(op, 'w') as f:
        print(dictZipToCountyList, file=f)

    return dictZipToCountyList

if __name__ == "__main__":

    create_ZipToCountyList_dict('dictZipToCountyList.py')

