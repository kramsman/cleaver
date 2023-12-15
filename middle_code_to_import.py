def middle_code_func(df):
    import numpy as np
    np.random.seed(42)
    def clean_field(fld, case_convert='lower'):
        """
        returns a string in lower, strip, no space, no -, no ., no '
        can be used with dataframe like IP['clean2'] = IP['B'].apply(clean_field, convert_case='keep')
        1/28/23 added optional paramter convert_case defaulting to lower, as was done before, but allowing 'upper' or 'keep'.
        """
        return_fld = str(fld).strip().replace(" ", "").replace("'", "").replace(".", "").replace("-", "")
        if case_convert == 'lower':
            return_fld = return_fld.lower()
        elif case_convert == 'upper':
            return_fld = return_fld.upper()
        elif case_convert == 'keep':
            pass
        else:
            exit_yes(f"wrong value fed to clean_field parameter case_convert, '{case_convert}' - exiting")
        return return_fld
        """  pass ethnic_group and set simple race to use in group """ 
    df['address'] = df.apply(lambda row: (row['addresstemp'] if type(row['address2'])!=str or row['address2'].strip() == '' else row['addresstemp']  + ", " + row['address2']), axis='columns' )
    df.loc[df['lastname'].str.strip() =='', 'lastname'] = "."
    df.loc[df['firstname'].str.strip() =='', 'firstname'] = "."
    df.loc[df['age'] =='', 'age'] = "."
    df.loc[df['sex'].str.strip() =='', 'sex'] = "."
    df['carol'] = df['pull_group'].map(lambda x: (True if x == 1 else False) )
    df['early_addresses'] = df['custom_field'].map(lambda x: (True if x in ['d','e'] else False) )
    df['age_group'] = df['age'].map(lambda x: ('Young' if x <50 else 'Senior') )
    df['race'] = df['rownum'].map(lambda x: ('Hispanic' if x % 2==0 else 'Black') )
    df.loc[df['filename'] == '2023 HD58 - BIPOC.csv', 'race'] = "Asian"
