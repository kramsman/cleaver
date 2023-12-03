def last_code_func(df, removedict, expectedstate):
    from collections import namedtuple
    Full_Address = namedtuple("FullAddress", ["state", "county", "city", "address"])
    def concentrated_address(full_address) -> bool:
        return (full_address.state.lower(), full_address.county.lower(),full_address.city.lower(),
                   full_address.address.lower()) in removedict
    def address_desc(full_address) -> str:
        return removedict.get((full_address.state.lower(), full_address.county.lower(),
                   full_address.city.lower(), full_address.address.lower()), 'Other')[0]
    def address_remove_desc(full_address) -> str:
        return removedict.get((full_address.state.lower(), full_address.county.lower(),
               full_address.city.lower(), full_address.address.lower()), 'Other')[1]
    df['full_address'] = df.apply(lambda row: Full_Address(row['state'], row['county'], row['city'], row['address']), axis='columns')
    df.loc[df['address'].str.lower() == 'general delivery', 'remove'] = "General Delivery"
    df.loc[df['address'].str.startswith("0"), 'remove'] = "Address begins 0"
    df.loc[df['full_address'].apply(concentrated_address),'remove'] = df['full_address'].apply(address_remove_desc)
    df.loc[df['mismatch_county'] == 1, 'remove'] = "County zip mismatch"
    df.loc[(df['firstname'] == '') | (df['firstname'] == '.'), 'remove'] = "Blank firstname"
    df.loc[(df['lastname'] == '') | (df['lastname'] == '.'), 'remove'] = "Blank lastname"
    df.loc[(df['address'] == '') | (df['address'] == '.'), 'remove'] = "Blank Address"
    df.loc[(df['city'] == '') | (df['city'] == '.'), 'remove'] = "Blank City"
    df.loc[(df['zip'] == '') | (df['zip'] == '.'), 'remove'] = "Blank Zip"
    df.loc[df['state'] != expectedstate, 'remove'] = "State"
    df.loc[df['dupe_id_field'].isin(['L','D','O']), 'remove'] = "Duplicate other than first"  # usually keep 'F', 'X' of choices: (F)irst, (L)ast, other (D)upe, (O)ther [not used], 'X':not dupe
    df.loc[df['carol'] == True, 'remove'] = "Assigned to Carol"
    df.loc[(df['remove'] == '') & (df['pull_group'] != 3), 'remove'] = "Clean but pull group not 3"
