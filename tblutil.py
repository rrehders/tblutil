#! /usr/bin/env python3
"""Module: Command line utility to perform manipulation of table type data files"""

import sys
import argparse
import os
import openpyxl
import csv
import re

__author__ = 'rrehders'


# Validate the input file
def valid_file(fname):
    if not os.path.isfile(fname):
        print('ERR: '+fname+' is not a file')
        return False
    else:
        return True


def whatfiletype(fname):
    # File extension determines the type of file to process
    if re.match('\.XLS$', sys.argv[1], re.IGNORECASE):
        return 'excel'
    elif re.match('\.XLSX$', sys.argv[1], re.IGNORECASE):
        return 'excel'
    elif re.match('\.CSV$', sys.argv[1], re.IGNORECASE):
        return 'csv'
    else:
        print('ERR: '+fname+' is not a valid file type')
        return False


def tocsv(fname, cols=set(), sheetnum=-1):
    # validate file type
    if whatfiletype(fname) != 'excel':
        print('ERR: '+fname+' is not a valid filetype to convert to csv')
        print('Valid filetypes are: .xls, .xlsx')
        sys.exit()

    # Check for multiple columns to extract
    if len(cols) > 2:
        # Loop through the parameters skipping the first
        for idx in range(2, len(cols)):
            argument = sys.argv[idx].strip().upper()
            if argument.startswith('COL='):
                # Seperate columns by the ','
                val = argument[4:]
                params = val.split(',')
                for param in params:
                    if param.isdecimal():
                        cols.add(int(param))
                    elif param.isalpha():
                        cols.add(openpyxl.utils.column_index_from_string(param))
                    else:
                        print('ERR: '+argument+' is an invalid argument')
                        sys.exit()
            elif argument.startswith('SHEET='):
                val = argument[6:]
                if val.isdecimal():
                    sheetnum = int(val)

    # load the target workbook
    print('XLS2CSV: Convert an Excel worksheet to CSV')
    wb = None
    try:
        wb = openpyxl.load_workbook(fname, data_only=True)
    except Exception as err:
        print('ERR: '+fname+' '+str(err))
    # Get the sheetnames
    sheetnms = wb.get_sheet_names()

    # Validate the sheetnum from the commandline (if provided)
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
def extractcols(fname):
    # validate file type
    if not whatfiletype(fname):
        print('ERR: '+fname+' is not a valid filetype to extract columns')
        print('Valid filetypes are: .xls, .xlsx, .csv')
        sys.exit()


def create_parser():
    parser = argparse.ArgumentParser(
        description='Perform various table manipulations'
    )

    parser.add_argument(
        'action',
        choices=['tocsv', 'extract'],
        help='Action to perform on the table'
    )

    parser.add_argument(
        'file',
        help='File containing the table'
    )

    parser.add_argument(
        '--cols',
        nargs='?',
        action='append',
        help='Optional value to specify which columns to use'
    )
    return parser


# TODO Implement "MAIN" functionality
def main():
    # parse the command line
    parser = create_parser()
    args = parser.parse_args()
    if args.action is 'tocsv':
        tocsv(args.file)
    elif args.action is 'extract':
        extractcols(args.file)

    # command line arguments filter out any invalid actions
    print('Complete')

if __name__ == '__main__':
    main()
