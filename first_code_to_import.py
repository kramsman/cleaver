def first_code_func(df):
    df.loc[df['county'].str.strip().str.lower() == 'st croix', 'county'] = "SaintCroix"
