#! /usr/bin/env python3
"""Module: Command line utility to perform manipulation of table type data files"""

import sys
import argparse
import os
import openpyxl
import csv
import re

__author__ = 'rrehders'


class InvalidFileType(Exception):
    pass


class InvalidExcelColumn(Exception):
    pass


def isValidFile(fname):
    """
    Verifies fname is valid for the current file system
    :param fname: String containing the path and filename
    :return:
    """
    if not os.path.isfile(fname):
        print('ERR: '+fname+' is not a file')
        return False
    else:
        return True


def getFileType(fname):
    """
    Returns a standardized file type based on the extension of the filename provided
    :param fname: Sting containing the path and file
    :return: String of standardized file types {'excel', 'csv'} or False if an invalid file type
    """
    if re.search('\.XLS$', fname, re.IGNORECASE):
        return 'excel'
    elif re.search('\.XLSX$', fname, re.IGNORECASE):
        return 'excel'
    elif re.search('\.CSV$', fname, re.IGNORECASE):
        return 'csv'
    else:
        raise InvalidFileType(fname)

def cvtColsStrToSet(strCols):
    """
    Convert a string of numeric or letter columns to a set of numeric columns
    :param strCols: Original String to convert
    :return: set of integers corresponding to spreadsheet columns
    """
    cols=set()
    params = strCols.split(',')
    for param in params:
        if param.isdecimal():
            cols.add(int(param))
        elif param.isalpha():
            cols.add(openpyxl.utils.column_index_from_string(param))
        else:
            raise InvalidExcelColumn(param)
    return cols


def tocsv(fname, cols=set(), sheetnum=-1):
    # validate file type
    try:
        if getFileType(fname) != 'excel':
            raise InvalidFileType(fname)
    except InvalidFileType as err:
        print('ERR: '+fname+' is not a valid filetype to convert to csv')
        print('Valid filetypes are: .xls, .xlsx')
        sys.exit()

    # load the target workbook
    print('XLS2CSV: Convert an Excel worksheet to CSV')
    try:
        wb = openpyxl.load_workbook(fname, data_only=True)
    except Exception as err:
        print('ERR: '+fname+' '+str(err))
    # Get the sheetnames
    sheetnms = wb.get_sheet_names()

    # Validate the sheetnum (if provided)
    # and seek user input if command line is invalid or missing
    if sheetnum not in range(len(sheetnms)):
        sheetnum = -1
        # Display Sheets in the workbook and ask which sheet to convert
        # A CSV can only contain a single sheet
        print('Sheets in '+fname)
        for i in range(len(sheetnms)):
            print(' | '+str(i)+' - '+sheetnms[i], end='')
        print(' |')

        # Get sheet selection
        while sheetnum not in range(len(sheetnms)):
            sheetnum = int(input('Convert which sheet ? '))
        print('')

    # Set the active sheet to the selection
    print('Sheet: '+sheetnms[sheetnum])
    xlsheet = wb.get_sheet_by_name(sheetnms[sheetnum])

    if len(cols) == 0:
        # Build lists of values for each row
        print('Extracting all columns')
        table = []
        for rowOfCellObjs in xlsheet:
            row = []
            for cellObj in rowOfCellObjs:
                row += [cellObj.value]
                print('.', end='')
            table += [row]
            print('')
    else:
        cols.split(',')
        # Repair loop indexes and investigate openpyxl function for converting strings to indexes
        # Build lists of values for each row for the specified columns
        # Discard columns which are invalid
        cols = cols.intersection(range(xlsheet.get_highest_column()+1))
        print('Extracting columns: ' + str(cols))
        table = []
        for rowOfCellObjs in xlsheet:
            row = []
            col = 1
            for cellObj in rowOfCellObjs:
                if col in cols:
                    row += [cellObj.value]
                col += 1
                print('.', end='')
            table += [row]
            print('')

    # Set the output filename based on sheet name
    ofname = sheetnms[sheetnum]+'.csv'

    # Open output file
    ofile = open(ofname, mode='w', newline='')

    # attach csv_writer to the output file
    ow = csv.writer(ofile)

    # Write out each row of the csv file
    print('Writing file '+ofname+'.')
    for i in range(len(table)):
        print('.', end='')
        ow.writerow(table[i])
    print('')

    # Close output file
    ofile.close()


# TODO Implement additional actions the utility is to perform
def extractcols(fname, strcols):
    """
    Create a new file of the same type as the input file but containing only the columns identified
    :param fname: String containing the input file path, name and extension
    :param strcols: String containg the columns requested for extraction
    :return: Nothing
    """
    if not getFileType(fname):
        print('ERR: '+fname+' is not a valid filetype to extract columns')
        print('Valid filetypes are: .xls, .xlsx, .csv')
        sys.exit()

    return


def create_parser():
    parser = argparse.ArgumentParser(
        description='Perform various table manipulations'
    )

    parser.add_argument(
        '-cols',
        type=str,
        help='Optional value to specify which columns to use'
    )

    parser.add_argument(
        '-sheet',
        type=int,
        help='Optional value to specify which columns to use'
    )

    parser.add_argument(
        'action',
        type=str,
        choices=['tocsv', 'extract'],
        help='Action to perform on the table'
    )

    parser.add_argument(
        'file',
        type=str,
        help='File containing the table'
    )

    return parser


# TODO Implement "MAIN" functionality
def main():
    # parse the command line
    parser = create_parser()
    args = parser.parse_args()
    if args.action is 'tocsv':
        tocsv(args.file, args.sheet)
    elif args.action is 'extract':
        extractcols(args.file)

    # command line arguments filter out any invalid actions
    print('Complete')

if __name__ == '__main__':
    main()
