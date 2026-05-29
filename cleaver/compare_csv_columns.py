""" compare columns across csvs in a directory
"""

# TODO: Needs cleaning up of import, logging, select directory


def main_compare_csv_columns(initial_campaign_dir):
    """ compare columns across csvs in a directory
    """

    # from uvbekutils import get_dir_name
    from uvbekutils import select_file
    from loguru import logger
    import pandas as pd
    from pathlib import Path

    # from bekutils import setup_loguru
    # from uvbekutils import get_dir_name

    USE_HARDCODED_SETUP = False

    # setup_loguru("INFO", "DEBUG")

    # log_level = "DEBUG"  # used for log file; screen set to INFO. TRACE, DEBUG, INFO, WARNING, ERROR

    # INITIAL_CAMPAIGN_DIR = Path("~/Dropbox/Postcard Files/InputFiles/Campaigns").expanduser()
    # INITIAL_CAMPAIGN_DIR = Path("~/Dropbox/Postcard Files/TestInputFiles/").expanduser()

    logger.info('compare_csv_columns.py')

    if USE_HARDCODED_SETUP:
        csv_dir = Path("~/Dropbox/Postcard Files/"
                         "/InputFiles/Campaigns/FL General GOTV 6-2024/Update - Odd PDI format/Rawdata").expanduser()
    else:
        # str_csv_dir = get_dir_name("Select a DIRECTORY containing CSVs files to compare columns (usually 'Rawdata')",
        #                            "CSVs Directory",
        #                            initial_campaign_dir)
        str_csv_dir = select_file("Select a DIRECTORY containing CSVs files to compare columns (usually 'Rawdata')",
                                  start_dir=str(initial_campaign_dir),
                                  files_like="*",
                                  mode="dir",
                                  title2="CSVs Directory")
        csv_dir = Path(str_csv_dir)

    # get list of files with xls or csv
    csv_files_w_path = list(csv_dir.glob("*.csv"))

    # create list of dicts, each key a variable name corresponding to a file value,
    # eg xls_files_dict[0]['xls_name'] = 'abc.csv'.
    # xls_files_dict = [{'xls_w_path': xls_w_path, 'xls_stem': xls_w_path.stem, 'xls_name': xls_w_path.name}
    #                   for xls_w_path
    #                   in csv_files_w_path]

    csv_column_lists = []
    for csv in csv_files_w_path:
        dftemp  = pd.read_csv(csv, nrows=0)
        # csv_column_lists.append({[csv.name, dftemp.columns.tolist()])
        csv_dict = {'name': csv.name, 'columns': dftemp.columns.tolist()}
        # blank columns in csv come over as 'Unnamed xx' and should be removed
        # csv_dict['columns'] = [col for col in csv_dict['columns'] if 'Unnamed' not in col]
        csv_column_lists.append(csv_dict)

    csv_column_superset = set()
    for csv_column_list in csv_column_lists:
        csv_column_superset = csv_column_superset | set(csv_column_list['columns'])

    csv_columns_incommon = set(csv_column_lists[0]['columns'])  # need to initialize first
    for csv_column_list in csv_column_lists:
        csv_columns_incommon = csv_columns_incommon.intersection(set(csv_column_list['columns']))

    # missing and extra columns
    for csv_column_list in csv_column_lists:
        missing_columns = csv_column_superset -  set(csv_column_list['columns'])
        csv_column_list['missing_columns'] = missing_columns
        extra_columns = set(csv_column_list['columns']) - csv_columns_incommon
        csv_column_list['extra_columns'] = extra_columns
        a=1

    print()
    # csvs_with_extra_columns = [dict for dict in csv_column_lists if dict['extra_columns'] != set()]
    csvs_with_extra_columns = [dict for dict in csv_column_lists if dict['extra_columns']]
    if csvs_with_extra_columns:
        print('Columns extra over minimum in common')
        for csv_column_list in csvs_with_extra_columns:
            print(f"File:'{csv_column_list['name']}', Extra:'{csv_column_list['extra_columns']}'")
    else:
        print("No csvs have extra columns")
    print()

    csvs_with_missing_columns = [dict for dict in csv_column_lists if dict['missing_columns']]
    if csvs_with_missing_columns:
        print('Columns missing from all columns')
        for csv_column_list in csvs_with_missing_columns:
            print(f"File:'{csv_column_list['name']}', Missing:'{csv_column_list['missing_columns']}'")
    else:
        print("No csvs have missing columns")
    print()

    a=1

    # newline = '\n'  # cause can't use \n in f-string
    # # prompt stating two lists - skip with_csv, process no_csv - ok to continue?
    # exit_yes_no(f"The following files WILL BE converted to csvs: {newline}" +
    #             newline.join([file_info['xls_name'] for file_info in xls_wo_csv]) +
    #             f"{newline}{newline}The following files already HAVE CSVs and will be SKIPPED: {newline}" +
    #             newline.join([file_info['xls_name'] for file_info in xls_w_csv]),
    #             "Continue with convert?"
    #             )

    # read each csv into dataframe with options and write out with same name
    # for file_info in xls_wo_csv:
    #     logger.info(f"Reading '{file_info['xls_name']}'")
    #     df = pd.read_excel(file_info['xls_w_path'], dtype=str)
    #     logger.info(f"Writing '{(file_info['xls_stem'] + '.csv')}'")
    #     df.to_csv(csv_dir / (file_info['xls_stem'] + '.csv'), index=False)
    #     logger.debug('')

if __name__ == '__main__':
    from loguru import logger
    from uvbekutils import setup_loguru
    from uvbekutils import get_dir_name
    from pathlib import Path

    setup_loguru("INFO", "DEBUG")

    INITIAL_CAMPAIGN_DIR = Path("~/Dropbox/Postcard Files/InputFiles/Campaigns").expanduser()

    main_compare_csv_columns(INITIAL_CAMPAIGN_DIR)

    a=1
