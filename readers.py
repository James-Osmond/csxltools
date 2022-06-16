# -*- coding: utf-8 -*-
"""
Created on Thu Jun 16 14:52:33 2022

@author: James-Osmond
@email: james.osmond@pm.me
"""

def import_ONS_time_series(file_path, write_path = None,

                           column_name = None, period = 'quarterly',

                           first_period=None):

    """

    This function takes a time series (downloaded from the ONS website) and

    reformats it to be easier to use.

 

    Parameters

    ----------

    file_path : string

        Path to where the time series data are stored.

    write_path : string, optional

        Path where the time series data should be stored after formatting,

        optional. The default is None.

    column_name : string, optional

        Name of the column of data. The default is None.

    period : string, optional

        String value equal to 'annual', 'quarterly' or 'monthly'. It dictates

        which rows should be in the output data. The default is 'quarterly'.

    first_period : optional

        The first row name (string or int) that should be in the output data.

        The default is None.

 

    Raises

    ------

    ValueError

        If the specified first time period is not consistent with the period

        type specified (e.g. first_period is '1997 Q1', but period is

        '1997 JAN') a ValueError is raised.

 

    Returns

    -------

    time_series : pandas object

        Column data of the desired format.

 

    """

   

    

    year_regex = r'^[0-9]{4}$'

    quarter_regex = r'^[0-9]{4} Q[0-4]$'

    month_regex = r'^[0-9]{4} [A-Z]{3}$'

   

    time_series = pd.read_excel(file_path, header = None)

   

    if period == 'annual':

        period = 'Year'

        regex = year_regex

    elif period == 'monthly':

        period = 'Month'

        regex = month_regex

    else:

        period = 'Quarter'

        regex = quarter_regex

   

    time_series = time_series.rename(columns = {0: period})

    time_series = time_series.set_index(period)

    CDID = str(time_series.loc['CDID', 1])

    column_name = CDID if column_name == None else column_name

    time_series = time_series.rename(columns = {1: column_name})

    time_series = time_series.filter(regex=regex, axis=0)

   

    if period == 'Year':

        time_series.index = time_series.index.map(int)

   

    if first_period != None:

        if not re.match(regex, str(first_period)):

            raise ValueError('The first time period and period type do not '

                             'match up.')

        time_series = time_series.loc[first_period:,:]

    if write_path != None:

        write_to_excel(write_path, time_series, CDID)

    return time_series

 

 

def read_in_OECD(df_file_path, df_write_path = None, header_row = 5,

                 multiplier = 1, **kwargs):

    """

    This function takes a dataset that has been download from the OECD website,

    and reformats it to be easier to use.

 

    Parameters

    ----------

    df_file_path : string

        File path of dataset.

    df_write_path : string, optional

        File path that the newly formatted dataset should be written to,

        optional. The default is None.

    header_row : int, optional

        Row number that the data table starts on in EXCEL. (i.e. unlike

        pandas.read_excel(), row 1 in an Excel document corresponds to row 1

        here) The default is 5.

    multiplier : float, optional

        Multiply every value in the dataset by this number. The default is 1.

    **kwargs

        Extra keyword arguments. If the output excel sheet needs a named sheet,

        use sheet_names.

 

    Returns

    -------

    df : pandas DataFrame

        DataFrame in an easy-to-use format.

 

    """

   

    

    df = pd.read_excel(df_file_path, header = header_row-1)

    df = df.rename(columns = {'Country': 'Year'})

   

    # Get unnamed columns, unnamed because these are formatted differently.

    # The header row is taken up by a merged cell called 'Non-OECD Economies',

    # causing lots of these column names to be empty, and the actual country

    # names are the next row down.

    non_OECD = df.filter(regex = r'(Non-OECD Economies$|Unnamed: [0-9]+$)',

                         axis = 1)

    non_OECD_countries = non_OECD.iloc[0,:].tolist()

    non_OECD_columns = non_OECD.columns

   

    df = df.rename(columns = dict(zip(non_OECD_columns, non_OECD_countries)))

   

    df = df.rename(columns = dict(zip(

        ['United Kingdom', 'United States',

         'European Union – 27 countries (from 01/02/2020)',

         'European Union (28 countries)'],

        ['UK', 'US', 'EU', 'EU']

        )))

   

    # We don't need the mostly empty row below the header row, the unit row or

    # the row consisting mostly of i's.

   

    df = df.set_index('Year')

   

    # Filter out all rows that don't match a year in the index

    df = df.filter(regex = r'^[0-9]{4}$', axis = 0)

   

    # Convoluted way of removing any unnamed columns (if a column is still

    # unnamed at this stage, it means it was unnamed at the start, AND the cell

    # in the first row was also empty, so we will have a NaN value). This line

    # checks if there are any characters in the column name, and only accepts

    # the column if this is the case.

    df = df.loc[:, df.columns[df.columns.str.contains(r'.') == True]]

   

    if 'i' in df.columns:

        df = df.drop(columns = ['i'])

   

    # Empty cells should be empty, not have a '..' in them.

    df = df.replace('..', np.nan)

   

    if multiplier != 1:

        df *= multiplier

   

    df.index = df.index.map(int)

   

    if df_write_path != None:

        write_to_excel(df_write_path, df, **kwargs)

   

    return df