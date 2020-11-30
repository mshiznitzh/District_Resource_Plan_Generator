"""
Module Docstring
"""

__author__ = "MiKe Howard"
__version__ = "0.1.0"
__license__ = "MIT"


import logging
from logzero import logger
import pandas as pd
import glob
import os
import datetime as DT
from typing import Optional
from xlsxwriter.worksheet import (
    Worksheet, cell_number_tuple, cell_string_tuple, xl_rowcol_to_cell
)
from pathlib import Path
import tkinter as tk
from tkinter import filedialog
import numpy as np

#OS Functions
def filesearch(word=""):
    """Returns a list with all files with the word/extension in it"""
    logger.info('Starting filesearch')
    file = []
    for f in glob.glob("*"):
        if word[0] == ".":
            if f.endswith(word):
                file.append(f)

        elif word in f:
            file.append(f)
            #return file
    logger.debug(file)
    return file

def Change_Working_Path(path):
    # Check if New path exists
    if os.path.exists(path):
        # Change the current working Directory
        try:
            os.chdir(path)  # Change the working directory
        except OSError:
            logger.error("Can't change the Current Working Directory", exc_info = True)
    else:
        print("Can't change the Current Working Directory because this path doesn't exits")

#Pandas Functions
def Excel_to_Pandas(filename,check_update=False):
    logger.info('importing file ' + filename)
    df=[]
    if check_update == True:
        timestamp = DT.datetime.fromtimestamp(Path(filename).stat().st_mtime)
        if DT.datetime.today().date() != timestamp.date():
            root = tk.Tk()
            root.withdraw()
            filename = filedialog.askopenfilename(title =' '.join(['Select file for', filename]))

    try:
        df = pd.read_excel(filename, sheet_name=None)
        df = pd.concat(df, axis=0, ignore_index=True)
    except:
        logger.error("Error importing file " + filename, exc_info=True)

    df=Cleanup_Dataframe(df)
    logger.debug(df.info(verbose=True))
    return df

def Cleanup_Dataframe(df):
    logger.info('Starting Cleanup_Dataframe')
    logger.debug(df.info(verbose=True))
    # Remove whitespace on both ends of column headers
    df.columns = df.columns.str.strip()

    # Replace whitespace in column header with _
    df.columns = df.columns.str.replace(' ', '_')

    return df

def get_column_width(worksheet: Worksheet, column: int) -> Optional[int]:
    """Get the max column width in a `Worksheet` column."""
    strings = getattr(worksheet, '_ts_all_strings', None)
    if strings is None:
        strings = worksheet._ts_all_strings = sorted(
            worksheet.str_table.string_table,
            key=worksheet.str_table.string_table.__getitem__)
    lengths = set()
    for row_id, colums_dict in worksheet.table.items():  # type: int, dict
        data = colums_dict.get(column)
        if not data:
            continue
        if type(data) is cell_string_tuple:
            iter_length = len(strings[data.string])
            if not iter_length:
                continue
            lengths.add(iter_length)
            continue
        if type(data) is cell_number_tuple:
            iter_length = len(str(data.number))
            if not iter_length:
                continue
            lengths.add(iter_length)
    if not lengths:
        return None
    return max(lengths)

def set_column_autowidth(worksheet: Worksheet, column: int):
    """
    Set the width automatically on a column in the `Worksheet`.
    !!! Make sure you run this function AFTER having all cells filled in
    the worksheet!
    """
    maxwidth = get_column_width(worksheet=worksheet, column=column)
    if maxwidth is None:
        return
    worksheet.set_column(first_col=column, last_col=column, width=maxwidth)

def Genrate_Resource_Plan(scheduledf, Budget_item_df):

    #scheduledf = scheduledf[scheduledf['Region_Name'] == 'METRO WEST']
    scheduledf.drop_duplicates(subset='PETE_ID', keep='last', inplace=True)
    scheduledf = scheduledf[scheduledf.PROJECTCATEGORY != 'ROW']
    scheduledf.rename(columns={'BUDGETITEMNUMBER': 'Budget_Item_Number'}, inplace=True)
    new_header=Budget_item_df.iloc[0]
    Budget_item_df = Budget_item_df[1:]
    Budget_item_df.columns = new_header

    Budget_item_df.rename(columns={'Budget Item Number': 'Budget_Item_Number'}, inplace=True)
    Budget_item_df.rename(columns={'Description': 'Budget_Item'}, inplace=True)

    scheduledf = pd.merge(scheduledf, Budget_item_df, on='Budget_Item_Number')


    for district in np.sort(scheduledf.WORKCENTERNAME.dropna().unique()):
        writer = pd.ExcelWriter(district + ' District Resource Plan.xlsx', engine='xlsxwriter')
        for type in np.sort(scheduledf.PROJECTTYPE.dropna().unique()):
            filtereddf = scheduledf[(scheduledf['Estimated_In-Service_Date'] >= pd.to_datetime('2021-01-01')) &
                                    (scheduledf['Estimated_In-Service_Date'] <= pd.to_datetime('2021-12-31')) &
                                    (scheduledf['PROJECTTYPE'] == type) &
                                    (scheduledf['WORKCENTERNAME'] == district)]



            outputdf = filtereddf.sort_values(by=['PLANNEDCONSTRUCTIONREADY'], ascending=True)
            outputdf = filtereddf[list(('PETE_ID',
                                      'WA',
                                      'Project_Name',
                                      'Description',
                                      'PLANNEDCONSTRUCTIONREADY',
                                      'Estimated_In-Service_Date',
                                      'Budget_Item',
                                      ))]

            if type == 'Station':
                Work= [ 'P&C Work',
                        'Set Steel',
                        'Weld Bus',
                        'Set Switches',
                        'Install Jumpers',
                        'Dress Transformer',
                        'Above Grade Demo',
                        'Set Breakers',
                        'Remove Old Breakers'
                      ]
            else:
                Work = ['Build Lattice',
                        'FCC',
                        'Install Insulators',
                        'Set Switches',
                        'Replace Arms'
                        ]

            Work.sort()

            for item in Work:
                outputdf[item] = np.nan

            outputdf['Comments'] = np.nan

            outputdf['PLANNEDCONSTRUCTIONREADY'] = outputdf['PLANNEDCONSTRUCTIONREADY'].dt.date
            outputdf['Estimated_In-Service_Date'] = outputdf['Estimated_In-Service_Date'].dt.date

            outputdf['PLANNEDCONSTRUCTIONREADY'] = outputdf['PLANNEDCONSTRUCTIONREADY'].dropna().astype(str)
            outputdf['Estimated_In-Service_Date'] = outputdf['Estimated_In-Service_Date'].dropna().astype(str)
            outputdf['WA'] = outputdf['WA'].dropna().astype(str)
            #outputdf['Earliest_PC_Delivery'] = outputdf['Earliest_PC_Delivery'].dropna().astype(str)
            #outputdf['Estimated_In-Service_Date'] = outputdf['Estimated_In-Service_Date'].dropna().astype(str)

            outputdf.rename(columns={'PLANNEDCONSTRUCTIONREADY': 'Construction Ready'}, inplace= True)
            outputdf.rename(columns={'Project_Name': 'Project Name'}, inplace= True)

            # Create a Pandas Excel writer using XlsxWriter as the engine.
            # Save the unformatted results

            if len(outputdf) >= 1:
                outputdf.to_excel(writer, index=False, sheet_name=district + ' ' + type)

                # Get workbook
                workbook = writer.book
                worksheet = writer.sheets[district + ' ' + type]

                # There is a better way to so this but I am ready to move on
                # note that PETE ID is diffrent from the ID used to take you to a website page
                x = 0
                for row in filtereddf.iterrows():
                    worksheet.write_url('A' + str(2 + x),
                                        'https://pete.corp.oncor.com/pete.web/project-details/' + str(
                                            filtereddf['PROJECTID'].values[x]),
                                        string=str('%05.0f' % filtereddf['PETE_ID'].values[x]))  # Implicit format
                    x = x + 1

                for column in outputdf.columns:
                    index = outputdf.columns.get_loc(column)
                    if column == 'P&C Work':
                        worksheet.data_validation(
                            xl_rowcol_to_cell(1, index) + ':' + xl_rowcol_to_cell(outputdf.shape[0], index),
                            {'validate': 'list', 'source': [
                                'Commissioning Group',
                                'District',
                                'N/A',
                                'Outside District'
                            ], })

                    elif column in Work:
                        worksheet.data_validation(
                            xl_rowcol_to_cell(1, index) + ':' + xl_rowcol_to_cell(outputdf.shape[0], index),
                            {'validate': 'list', 'source': [
                                'District will do all',
                                'District will do some, see comments',
                                'N/A',
                                'Outside District'

                            ], })
            cell_format = workbook.add_format()

            cell_format = workbook.add_format()
            cell_format.set_align('center')
            cell_format.set_align('vcenter')
            worksheet.set_column('A:' + chr(ord('@') + len(outputdf.columns)), None, cell_format)

            for x in range(len(outputdf.columns)):
                set_column_autowidth(worksheet, x)

            wrap_format = workbook.add_format()
            wrap_format.set_text_wrap()
            wrap_format.set_align('vcenter')
            worksheet.set_column('C:D', None, wrap_format)
            worksheet.set_column('C:D', 100)

        writer.save()
        writer.close()

def main():
    Project_Data_Filename ='All Project Data.xlsx'
    Budget_Item_Filename = 'Budget Item.xlsx'

    """" Main entry point of the app """
    logger.info("Starting Pete Maintenance Helper")
    Change_Working_Path('./Data')
    try:
        Project_Data_df=Excel_to_Pandas(Project_Data_Filename, True)
    except:
        logger.error('Can not find Project Data file')
        raise

    try:
        budget_item_df = Excel_to_Pandas(Budget_Item_Filename)
    except:
        logger.error('Can not find Budget Item Data file')

   # Project_Schedules_All_Data_df = pd.merge(Project_Schedules_df, Project_Data_df, on='PETE_ID', sort= False, how='outer')

    Genrate_Resource_Plan(Project_Data_df, budget_item_df)

if __name__ == "__main__":
    """ This is executed when run from the command line """
    # Setup Logging
    logger = logging.getLogger('root')
    FORMAT = "[%(filename)s:%(lineno)s - %(funcName)20s() ] %(message)s"
    logging.basicConfig(format=FORMAT)
    logger.setLevel(logging.DEBUG)

    main()