#! /usr/bin/env python3
"""Module: Command line utility to perform manipulation of table type data files"""

import cmnfns
import argparse
import openpyxl
import warnings
import csv
import os

__author__ = 'rrehders'


def cleanpostal(fname, col='', sheet='0'):
    """
    Output a CSV file of the corect postal codes from the input file
    :param fname: String containing the input file path, name and extension
    :param col: String containg the column index for the target postal code information
    :param sheet: specific sheet target within an excel file
    :return: True if successful, None if incomplete
    """
    # Convert columns argument from string to set
    idx = cmnfns.cvtcolsstrtoset(col).pop()
    # Validate that columns were requested
    if not idx:
        print('ERR: no column specified for extraction')
        return
    try:
        ftype = cmnfns.getfiletype(fname)
    except cmnfns.InvalidFileType as err:
        print('ERR: ' + fname + str(err))
        print('Valid filetypes are: .xls, .xlsx')
        return

    if ftype is 'excel':
        # load the target workbook
        print('Extract contents from Excel worksheet')
        try:
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(fname, data_only=True)
            sheetnms = wb.get_sheet_names()
        except Exception as err:
            print('ERR: ' + fname + ' ' + str(err))
            return

        # Validate sheet is valid for the file
        sheetnum = int(sheet)
        if sheetnum not in range(len(sheetnms)):
            print('ERR: invalid sheet specified for extraction')
            return

        # Set the active sheet to the selection
        print('Sheet: '+sheetnms[sheetnum])
        xlsheet = wb.get_sheet_by_name(sheetnms[sheetnum])

        # extract all cells from the input excel file as a list
        table = cmnfns.extractxltable(xlsheet)

        # extract cells from the input excel file identified as the column of postal codes
        column = [item for sublist in cmnfns.extractxltable(xlsheet, {idx}) for item in sublist]

        # Clean the postal codes
        postallist = cmnfns.cleancdnpostallist(column)

        # replace the existing postal items with the cleaned items
        for r in range(len(table)):
            table[r][idx-1] = postallist[r]

        # Set the output filename based on sheet name
        ofname = sheetnms[sheetnum]+'.csv'

        # Open output file
        ofile = open(ofname, mode='w', newline='')

        # attach csv_writer to the output file
        ow = csv.writer(ofile)

        # Write out each row of the csv file
        print('Writing file ' + ofname + '.')
        for i in range(len(table)):
            print('.', end='')
            ow.writerow(table[i])
        print('')

        # Close output file
        ofile.close()

    elif ftype is 'csv':
        print('Extract columns from an CSV File')
        with open(fname, 'r') as filein:
            csvin = csv.reader(filein)
            data = [row for row in csvin]

        table = cmnfns.extractlisttable(data)

        # extract cells from the input excel file identified as the column of postal codes
        column = [item for sublist in cmnfns.extractlisttable(data, {idx}) for item in sublist]

        # Clean the postal codes
        postallist = cmnfns.cleancdnpostallist(column)

        # replace the existing postal items with the cleaned items
        for r in range(len(table)):
            table[r][idx - 1] = postallist[r]

        # Set the output filename based on sheet name
        ofname = os.path.basename(fname)+'_extract.csv'

        # Open output file
        ofile = open(ofname, mode='w', newline='')

        # attach csv_writer to the output file
        ow = csv.writer(ofile)

        # Write out each row of the csv file
        print('Writing file ' + ofname + '.')
        for i in range(len(table)):
            print('.', end='')
            ow.writerow(table[i])
        print('')

        # Close output file
        ofile.close()

    # Completed all activities, function returns True to mark success
    return True


def create_parser():
    parser = argparse.ArgumentParser(
        description='Extract Columns from an excel or .csv file as a .csv file'
    )

    parser.add_argument(
        "-s",
        "--sheet",
        type=int,
        default=0,
        help='Optional value to specify which sheet to use'
    )

    parser.add_argument(
        'file',
        type=str,
        help='File containing the table'
    )

    parser.add_argument(
        'col',
        type=str,
        help='The column index containing the postal code information'
    )

    return parser


def main():
    parser = create_parser()
    args = parser.parse_args()
    cleanpostal(args.file, args.col, args.sheet)
    print('Complete')

if __name__ == '__main__':
    main()
