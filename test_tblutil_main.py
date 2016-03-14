from tblutil import InvalidFileType, InvalidExcelColumn
from tblutil import getfiletype, cvtcolsstrtoset, extractxltable, extractlisttable
from tblutil import cvtstrindextoset, tocsv, extractcols, joinfiles
import unittest
import openpyxl
import os


class MainTestCases(unittest.TestCase):
    def test_tblutil_tocsv_fail(self):
        self.assertFalse(tocsv('/users/rrehders/test/test.xls', 0))

    def test_tblutil_tocsv_success_minimum_parameter(self):
        self.assertTrue(tocsv('/users/rrehders/test/test.xlsx', 0))
        os.system('rm ./*.csv')

    def test_tblutil_tocsv_success_parameters(self):
        self.assertTrue(tocsv('/users/rrehders/test/test.xlsx', 0, cvtcolsstrtoset('A,B')))
        os.system('rm ./*.csv')

    def test_tblutils_extractcols_fail_bad_file(self):
        with self.assertRaises(InvalidFileType):
            extractcols('/users/rrehders/test/fail.txt', cvtcolsstrtoset('A'))

    def test_tblutils_extractcols_fail_no_cols(self):
        with self.assertRaises(TypeError):
            extractcols('/users/rrehders/test/test.csv')

    def test_tblutils_extractcols_fail_empty_cols(self):
        self.assertFalse(extractcols('/users/rrehders/test/test.csv', set()))

    def test_tblutils_extractcols_success_csv(self):
        self.assertTrue(extractcols('/users/rrehders/test/test.csv', cvtcolsstrtoset('A,B')))
        os.system('rm ./*.csv')

    def test_tblutils_extractcols_success_xlsx(self):
        self.assertTrue(extractcols('/users/rrehders/test/test.xlsx', cvtcolsstrtoset('A,B'), 0))
        os.system('rm ./*.csv')
