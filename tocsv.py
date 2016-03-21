#! /usr/bin/env python3
""" Script: tocsv converts an excel file (.xlsx format) to .csv
"""

import argparse
import warnings
import openpyxl
import cmnfns

import cmnfns
__author__ = 'rrehders'


def tocsv(fname, sheetnum=-1, cols=''):
    """
    Perform the conversion of a sheet from an excel file to a csv file
    :param fname: String containing the filename of the input Excel document
    :param sheetnum: Index of the sheet to be converted
    :param cols: Optional string listing the columns to be included
    :return: True if successful, None if incomplete
    """
    # validate file type
    try:
        if cmnfns.getfiletype(fname) != 'excel':
            raise cmnfns.InvalidFileType(fname)
    except cmnfns.InvalidFileType:
        print('ERR: '+fname+' is not a valid filetype to convert to csv')
        print('Valid filetypes are: .xls, .xlsx')
        return

    # load the target workbook
    print('Convert an Excel worksheet to CSV')
    try:
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(fname, data_only=True)
        sheetnms = wb.get_sheet_names()
    except Exception as err:
        print('ERR: '+fname+' '+str(err))
        return

    # Validate the sheetnum, -1 represents all sheets
    if sheetnum != -1:
        if sheetnum not in range(len(sheetnms)):
            # seek user input, provided sheetnum was invalid
            sheetnum = -1
            # Display Sheets in the workbook and ask which sheet to convert
            print('Sheets in '+fname)
            for i in range(len(sheetnms)):
                print(' | '+str(i)+' - '+sheetnms[i], end='')
            print(' |')

            # Get sheet selection
            while sheetnum not in range(len(sheetnms)):
                sheetnum = int(input('Convert which sheet ? '))
            print('')

    # Loop through the workbook's sheets
    for sheetidx in range(len(sheetnms)):
        # test each sheet to see if it should be converted and output
        if sheetnum < 0 or sheetidx == sheetnum:
            # Set the active sheet to the selection
            print('Sheet: '+sheetnms[sheetnum])
            xlsheet = wb.get_sheet_by_name(sheetnms[sheetnum])

            # extract cells from the input excel file and set of columns
            table = cmnfns.extractxltable(xlsheet, cmnfns.cvtcolsstrtoset(cols))

            # output a csv file of table with the hybrid file-sheet name
            cmnfns.outcsvtable(table, fname, sheetnms[sheetidx])

    # Completed all activities, function returns True to mark success
    return True


def create_parser():
    parser = argparse.ArgumentParser(
        description='Utility to convert and Excel (.xlsx) file to .csv'
    )

    parser.add_argument(
        "-c",
        "--cols",
        type=str,
        help='Optional value to specify which columns to use'
    )

    parser.add_argument(
        "-s",
        "--sheet",
        type=int,
        default=-1,
        help='Optional value to specify which sheet to convert (default is all)'
    )

    parser.add_argument(
        'file',
        type=str,
        help='File containing the table'
    )

    return parser


def main():
    parser = create_parser()
    args = parser.parse_args()
    tocsv(args.file, args.sheet, args.cols)

    print('Complete')

if __name__ == '__main__':
    main()
