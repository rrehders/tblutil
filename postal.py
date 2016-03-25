#! /usr/bin/env python3
"""Module: Command line utility to perform manipulation of table type data files"""

import cmnfns
import argparse
import openpyxl
import warnings
import csv
import os

__author__ = 'rrehders'


def extractxltable(xlsheet, cols=set()):
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
        # Discard columns which are invalid
        cols = cols.intersection(range(xlsheet.max_column+1))
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
    return table


def extractlisttable(csvsheet, cols=set()):
    if len(cols) == 0:
        # Build lists of values for each row
        print('Extracting all columns')
        table = []
        for rowOfCells in csvsheet:
            row = []
            for cellObj in rowOfCells:
                row += [cellObj]
                print('.', end='')
            table += [row]
            print('')
    else:
        # Discard columns which are invalid
        cols = cols.intersection(range(len(csvsheet)+1))
        print('Extracting columns: ' + str(cols))
        table = []
        for rowOfCells in csvsheet:
            row = []
            col = 1
            for cellObj in rowOfCells:
                if col in cols:
                    row += [cellObj]
                col += 1
                print('.', end='')
            table += [row]
            print('')
    return table


def cleanpostal(fname, col='', sheet='0'):
    """
    Output a CSV file of the corect postal codes from the input file
    :param fname: String containing the input file path, name and extension
    :param col: String containg the column index for the target postal code information
    :param sheet: specific sheet target within an excel file
    :return: True if successful, None if incomplete
    """
    # Convert columns argument from string to set
    idx = cmnfns.cvtcolsstrtoset(col)
    # Validate that columns were requested
    if not len(idx):
        print('ERR: no column specified for extraction')
        return

    if cmnfns.getfiletype(fname) is 'excel':
        # load the target workbook
        print('Extract columns from an Excel worksheet')
        try:
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(fname, data_only=True)
            sheetnms = wb.get_sheet_names()
        except UserWarning:
            pass
        except Exception as err:
            print('ERR: '+fname+' '+str(err))
            return

        # Validate sheet is valid for the file
        sheetnum = int(sheet)
        if sheetnum not in range(len(sheetnms)):
            print('ERR: invalid sheet specified for extraction')
            return

        # Set the active sheet to the selection
        print('Sheet: '+sheetnms[sheetnum])
        xlsheet = wb.get_sheet_by_name(sheetnms[sheetnum])

        # extract cells from the input excel file and set of columns
        table = extractxltable(xlsheet, idx)

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

    elif cmnfns.getfiletype(fname) is 'csv':
        print('Extract columns from an CSV File')
        with open(fname, 'r') as filein:
            csvin = csv.reader(filein)
            data = [row for row in csvin]

        table = extractlisttable(data, idx)
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
