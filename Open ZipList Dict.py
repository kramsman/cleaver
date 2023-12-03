import ast
ZipToCountyList_file = 'dictZipToCountyList.py' # file where the numeric zip to county list is stored (ie  1011: ['hampden', 'hampshire'])

dict_file = open(ZipToCountyList_file, "r")
dict_file_data = dict_file.read()
dict_file.close()
dictZipToCountyList = ast.literal_eval(dict_file_data)
# print("Ran create_ZipToCounty_dict()")
print("Imported dictZipToCountyList.py")