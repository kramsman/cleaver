# test import an outide file into a program

import ast
import math
import datetime
import pymsgbox
import os
import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils import column_index_from_string
from random import random
import numpy as np
import pandas as pd

from tkinter import *
from tkinter import Tk  # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
from tkinter import messagebox
from uszipcode import SearchEngine
from math import floor

search = SearchEngine(simple_zipcode=True)

translateCounty = {
"alexandria": "alexandriacity",
"arlington": "arlingtoncounty",
"chesapeake": "chesapeakecity",
"fredericksburg": "fredericksburgcity",
"henrico": "henricocounty",
"loudoun": "loudouncounty",
"lynchburg":"lynchburgcity",
"newportnews": "newportnewscity",
"petersburg": "petersburgcity",
"princewilliam": "princewilliamcounty",
"richmondcity":"richmondcity",
"roanokecity":"roanokecity",
"chesterfieldcounty":"chesterfieldcounty",
"hampton":"hamptoncity",
"norfolk":"norfolkcity",
"virginiabeach":"virginiabeachcity",
"suffolk":"suffolkcity",
"fairfaxcounty": "fairfaxcounty"
}


def is_number(s): #expects parm to be a string to trap all types of data.  use IP = IP.astype({"MZIP":str}) to force
    # qq = np.isnan(s)
    # if qq:
    #     return False
    if s == "nan": return False
    try:
        float(s)
        return True
    except ValueError:
        return False

def padded_zip(zzip):
    if is_number(zzip):
        zipnum = floor(float(zzip))
        zp = format(zipnum, '05')
    else:
        zp = '00000'
    return zp

def county_from_zip(zzip):
    # fill res_county with zip lookup
    zp = padded_zip(zzip)

    # IP.loc[index, 'zip'] = zp

    county_from_zip = str(search.by_zipcode(zp).county).lower().replace(" ", "")
    county_from_zip2 = translate_county(county_from_zip)

    return county_from_zip2

    # IP.at[index, 'res_county'] = county_from_zip

    # cleanCounty = str(IP.at[index, 'county']).lower().replace(" ", "")
    # IP.at[index, 'county'] = cleanCounty  # don't permanently store clean county in lower case with spaces removed

    # use dictionary translateCounty to reformat a files county format into te zip lookup format

def translate_county(cnty):
    cleanCounty = str(cnty).lower().replace(" ", "").replace("'", "").replace(".", "")
    if cleanCounty in translateCounty:
        transcounty = translateCounty[cleanCounty]
    else:
        transcounty = cleanCounty
    return transcounty

def check_mismatch(a,b):
    if a != b: return 1
    else: return 0

# IP = pd.DataFrame({'A':[13.1,"","a", 3345],'B':["x", "y", "z","q"]})
# IP = pd.DataFrame([1,"x"],[2,"y"],[3,"z"], columns=['A','B'])

# **** MAIN

rows_to_read_limit = 999999
# rows_to_read_limit = 999



if False:
    choice = pymsgbox.confirm("Do you want to update the zip dictionary file (it will take time)?", "Update Zip Data", ["OK to continue?", 'Cancel'])
    if choice != 'OK to continue?': exit()

    filename = '/Users/Denise/Dropbox/Postcard Files/Other/VoterLetters/zip_dict_char_file.py'

    zip_list = []

    for numZip in range(99999):
        zip = padded_zip(numZip)
        dirtyCounty = county_from_zip(zip)
        county = translate_county(dirtyCounty)
        zip_list.append(tuple((zip,county)))  # could exclude 'non' but then would need to trap keys not found

    zip_dict_char = dict(zip_list)

    with open(filename, 'w') as data:
        data.write(str(zip_dict_char))

    print(zip_dict_char['00544'])


if True:
    choice = pymsgbox.confirm("Do you want to update the zip dictionary file (it will take time)?", "Update Zip Data", ["OK to continue?", 'Cancel'])
    if choice != 'OK to continue?': exit()

    filename = '/Users/Denise/Dropbox/Postcard Files/Other/VoterLetters/zip_dict_num_file.py'

    zip_list = []

    for numZip in range(99999):
        if numZip % 500 == 0: print(numZip)
        zip = padded_zip(numZip)
        dirtyCounty = county_from_zip(zip)
        county = translate_county(dirtyCounty)
        zip_list.append(tuple((numZip,county)))  # could exclude 'non' but then would need to trap keys not found

    zip_dict_num = dict(zip_list)

    with open(filename, 'w') as data:
        data.write(str(zip_dict_num))

    print(zip_dict_num[544])


if False:
    # choice = pymsgbox.confirm("Do you want to update the zip dictionary file (it will take time)?", "Update Zip Data", ["OK to continue?", 'Cancel'])
    # if choice != 'OK to continue?': exit()

    filename = '/Users/Denise/Dropbox/Postcard Files/Other/VoterLetters/zip_dict_char_file.py'

    file = open(filename, "r")
    lines = file.read()
    file.close()
    zip_dict_char = ast.literal_eval(lines)  # create dictionary object from string

    print(zip_dict['00544'])

if False:
    # choice = pymsgbox.confirm("Do you want to update the zip dictionary file (it will take time)?", "Update Zip Data", ["OK to continue?", 'Cancel'])
    # if choice != 'OK to continue?': exit()

    filename = '/Users/Denise/Dropbox/Postcard Files/Other/VoterLetters/zip_dict_num_file.py'

    file = open(filename, "r")
    lines = file.read()
    file.close()
    zip_dict_num = ast.literal_eval(lines)  # create dictionary object from string

    print(zip_dict[544])

a=1
