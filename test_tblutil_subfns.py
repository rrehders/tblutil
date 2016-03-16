from tblutil import InvalidFileType, InvalidExcelColumn
from tblutil import getfiletype, cvtcolsstrtoset, extractxltable, extractlisttable
from tblutil import cvtstrindextoset, tocsv, extractcols, joinfiles
import unittest
import openpyxl
import os


class SubFunctionTestCases(unittest.TestCase):
    def test_tblutil_getFileType_XLSX(self):
        self.assertTrue(getfiletype('Test.xlsx') == 'excel')

    def test_tblutil_getFileType_CSV(self):
        self.assertTrue(getfiletype('Test.csv') == 'csv')

    def test_tblutil_getFileType_other(self):
        with self.assertRaises(InvalidFileType):
            getfiletype('Test.txt')

    def test_tblutil_cvtcolsstrtoset_err(self):
        with self.assertRaises(InvalidExcelColumn):
            cvtcolsstrtoset('A1')

    def test_tblutil_cvtcolsstrtoset_num(self):
        result = cvtcolsstrtoset('1,3,7')
        self.assertTrue(result-{1, 2, 3})

    def test_tblutil_cvtcolsstrtoset_alpha(self):
        result = cvtcolsstrtoset('A,C,G')
        self.assertTrue(result-{1, 2, 3})
        
    def test_tblutil_cvtstrindextoset_empty(self):
        result = cvtstrindextoset('')

    def test_tblutil_cvtstrindextoset_single_num(self):
        result = cvtstrindextoset('1')

    def test_tblutil_cvtstrindextoset_single_alpha(self):
        result = cvtstrindextoset('A')

    def test_tblutil_cvtstrindextoset_double_alpha(self):
        result = cvtstrindextoset('C,B')

    def test_tblutil_extractxltable_no_columns(self):
        try:
            wb = openpyxl.load_workbook('/users/rrehders/test/test.xlsx', data_only=True)
        except UserWarning as warn:
            pass
        xlsheet = wb.worksheets[0]
        result = extractxltable(xlsheet)
        self.assertTrue(len(result[0]) == xlsheet.max_column)

    def test_tblutil_extractxltable_three_columns(self):
        try:
            wb = openpyxl.load_workbook('/users/rrehders/test/test.xlsx', data_only=True)
        except UserWarning as warn:
            pass
        xlsheet = wb.worksheets[0]
        result = extractxltable(xlsheet, {1, 2, 3})
        self.assertTrue(len(result[0]) == 3)

    def test_tblutil_extractlisttable_no_columns(self):
        table = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
        result = extractlisttable(table)
        self.assertTrue(len(result[0]) == len(table[0]))

    def test_tblutil_extractlisttable_three_columns(self):
        table = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
        result = extractlisttable(table, {1, 2, 3})
        self.assertTrue(len(result[0]) == 3)

    def test_tblutil_outcsvtable_(selfself):
        table = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
