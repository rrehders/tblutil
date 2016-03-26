#! /usr/bin/env python3
"""Module: Command line utility to perform manipulation of table type data files"""

import openpyxl
import csv
import re
import os

__author__ = 'rrehders'


class InvalidFileType(Exception):
    pass


class InvalidExcelColumn(Exception):
    pass


def getfiletype(fname):
    """
    Returns a standardized file type based on the extension of the filename provided
    :param fname: Sting containing the path and file
    :return: String of standardized file types {'excel', 'csv'} or False if an invalid file type
    """
    if re.search('\.xlsx$', fname, re.IGNORECASE):
        return 'excel'
    elif re.search('\.csv$', fname, re.IGNORECASE):
        return 'csv'
    else:
        raise InvalidFileType(fname)


def cvtcolsstrtoset(strcols):
    """
    Convert a string of numeric or letter columns to a set of integers
    :param strcols: Original String to convert
    :return: A set of integers corresponding to spreadsheet columns or a single integer
    """
    cols = set()
    params = strcols.split(',')
    for param in params:
        if param.isdecimal():
            cols.add(int(param))
        elif param.isalpha():
            cols.add(openpyxl.utils.column_index_from_string(param))
        elif param == '':
            continue
        else:
            raise InvalidExcelColumn(param)
    return cols


def extractxltable(xlsheet, cols=set()):
    """
    Create a list of lists containing the values of the cells in a spreadsheet
    :param xlsheet: an openpyxl worksheet object of input data
    :param cols: a set corresponding to column indexes to be extracted
    :return: a list of lists
    """
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
    """
    Create a list of lists containing the values of the cells in a spreadsheet
    :param csvsheet: a list os lists containing the input data
    :param cols: a set corresponding to column indexes to be extracted
    :return: a list of lists
    """
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


def outcsvtable(table, fnamebase='/output.csv', subname=''):
    """
    :param table: Data to be written
    :param fnamebase: Path and base Filename
    :param subname: Name to be appended to the base
    :return:
    """
    # Set the output filename based on sheet name
    ofname = os.path.splitext(os.path.basename(fnamebase))[0] + '-' + subname + '.csv'

    # Open output file
    ofile = open(ofname, mode='w', newline='')

    # attach csv_writer to the output file
    ow = csv.writer(ofile)

    # Write out each row of the csv file
    print('Writing file ' + ofname)
    for i in range(len(table)):
        ow.writerow(table[i])
    print('')

    # Close output file
    ofile.close()


def cvtjoinstrtoindex(strindex):
    """
    :param strindex: String containing column indecies
    :return: Tuple of 2 column indecies
    """
    if strindex is None or strindex == '':
        return 0, 0
    else:
        strindex = strindex.rstrip()

    if len(strindex.split(',')) == 1:
        return cvtcolsstrtoset(strindex).pop(), cvtcolsstrtoset(strindex).pop()
    else:
        indexes = strindex.rstrip()
        return (idx for idx in cvtcolsstrtoset(indexes))


def cleancdnpostallist(pcodes):
    """
    Returns a list after having scubbed the input list to canadian postal code format.
    If the error rate for postal code prmatting is > 80% returns None
    :param pcodes a list of strings corresponding to postal cods
    """
    # Set up a Canadian postal code regex
    postalregex = re.compile(r'''(?!.*[DFIOQU])      # Eliminate invalid starting letters
        ([A-VXY]             # Valid starting letters
        \d                  # Number
        [A-Z])              # Letter
        .?                  # Optional seperator
        (\d                 # Number
        [A-Z]               # Letter
        \d                  # Number
        )
        ''', re.IGNORECASE | re.VERBOSE)

    # Vewy Pythonic
    return [postalregex.sub(r'\1 \2', postal.strip()) for postal in pcodes]


def main():
    pass


if __name__ == '__main__':
    main()
