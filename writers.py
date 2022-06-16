# -*- coding: utf-8 -*-
"""
Created on Thu Mar  3 13:48:20 2022

@author: James-Osmond
@email: james.osmond@pm.me
"""

import pandas as pd

import xlsxwriter

 

def write_to_excel(write_path, dataframes, sheet_names = None):

    """

    Writes DataFrames to an Excel workbook.

 

    Parameters

    ----------

    write_path : string

        File path to which we wish to write our Excel spreadsheet.

    dataframes : list

        List of pandas DataFrames. Each DataFrame will be given its own

        worksheet.

        Can also be given as dict object, with sheet names as keys, dataframes

        as values.

    sheet_names : list, optional

        List of sheet names (strings) for the workbook. The default is None.

 

    Returns

    -------

    None.

 

    """

    # Could be broken by giving dictionary object for dataframes AND

    # sheet_names.

    try:

        writer = pd.ExcelWriter(write_path)

        # Sheet names are not specified

        if sheet_names == None:

       

            if isinstance(dataframes, list):

                for df in dataframes:

                    df.to_excel(writer)

                writer.save()

           

            elif isinstance(dataframes, dict):

                for sheet_name, df in dataframes.items():

                    df.to_excel(writer, sheet_name)

                writer.save()

           

            else:

                dataframes.to_excel(writer)

                writer.save()

       

        # dataframes is not a list, implying there is only one dataframe.

        # Having this condition makes it easier to write something to Excel,

        # as it is not necessary to specify lists when calling the function.

        elif not isinstance(dataframes, list):

            if isinstance(sheet_names, list):

                dataframes.to_excel(writer, sheet_names[0])

                writer.save()

           

            else:

                dataframes.to_excel(writer, sheet_names)

                writer.save()

           

        # dataframes is a list and sheet names are specified

        else:

            if isinstance(sheet_names, list):

                for df, sheet_name in zip(dataframes, sheet_names):

                    df.to_excel(writer, sheet_name)

            else:

                dataframes[0].to_excel(writer, sheet_names)

            writer.save()

    except (xlsxwriter.exceptions.FileCreateError, PermissionError):

        print('Could not write to Excel. Check that the worksheet with this '

              f'file path is not in use. {write_path}')

 

 

def produce_chart_download(write_folder, data, figure_number, figure_name,

                           figure_content = None, notes = None, unit = None,

                           decimal_places = 1):

    """

    Produces a chart download for use in a publication. The file name of the

    download will be 'Data download - Figure X.xlsx', where X denotes the

    number of the Figure within the article.

 

    Parameters

    ----------

    write_folder : string

        Desired folder for the chart download to be written to.

    data : pandas DataFrame

        Data to appear in the chart and chart download. It should have one

        index column and one heading row.

    figure_number : int

        Number of the Figure in the article. Will be used in the file name.

    figure_name : string

        Name/ description of the chart.

    figure_content : string, optional

        Additional information about the chart, e.g. which years/industries

        does the chart contain information about? The default is None.

    notes : string, optional

        Any notes about the chart or data. The default is None.

    unit : string, optional

        The unit/s of the data. The default is None.

    decimal_places : int, optional

   

    Returns

    -------

    None.

 

    """

   

    try:

        title = f'Figure {figure_number}: {figure_name}'

        write_path = \
            f'{write_folder}Data download - Figure {figure_number}.xlsx'

        data = data.copy().astype(float).round(decimal_places)

        

        workbook = xlsxwriter.Workbook(write_path)

        if decimal_places > 0:

            # Creates a string with as many zeroes as desired decimal places,

            # for use in the f-string deciding number formats

            zeros = '0' * decimal_places

            number_format = workbook.add_format({'num_format':

                                                 f'#,##0.{zeros}'})

        else:

            number_format = workbook.add_format({'num_format': '#,##0'})

       

        worksheet = workbook.add_worksheet(f'Figure_{str(figure_number)}')

        worksheet.write(0, 0, title)

        worksheet.write(1, 0, figure_content)

        worksheet.write(3, 0, 'Notes')

        worksheet.write(3, 1, notes)

        worksheet.write(4, 0, 'Unit')

        worksheet.write(4, 1, unit)

       

        #worksheet.write(6, 0, data.index.name)

       

        # Write the headers and row names to the worksheet

        for row_num, i in enumerate(data.index):

            worksheet.write(7+row_num, 0, i)

        for col_num, j in enumerate(data.columns):

            worksheet.write(6, 1+col_num, j)

       

        # Writes the data itself to the worksheet, in a format that rounds all

        # entries to the same number of decimal places

        for i in range(data.shape[0]):

            for j in range(data.shape[1]):

                try:

                    worksheet.write(7+i, 1+j, data.iloc[i,j], number_format)

                except:

                    pass

        workbook.close()

    except (xlsxwriter.exceptions.FileCreateError, PermissionError):

        print(f'The data download \'{write_path}\' could not be created. '

              'Check that a file with this file path is not already open.')

       

 

 

