from tblutil import InvalidFileType, InvalidExcelColumn
from tblutil import getfiletype, cvtcolsstr, extractxltable, extractlisttable
from tblutil import cvtjoinstrindex, tocsv, extractcols, joinfiles
import unittest
import os


class MainTestCases(unittest.TestCase):
    def test_tblutil_tocsv_fail(self):
        self.assertFalse(tocsv('/users/rrehders/test/test.xls', 0))

    def test_tblutil_tocsv_success_minimum_parameters(self):
        self.assertTrue(tocsv('/users/rrehders/test/test.xlsx'))
        os.system('rm ./*.csv')

    def test_tblutil_tocsv_success_one_sheet(self):
        self.assertTrue(tocsv('/users/rrehders/test/test.xlsx',0))
        os.system('rm ./*.csv')

    def test_tblutil_tocsv_success_parameters(self):
        self.assertTrue(tocsv('/users/rrehders/test/test.xlsx', 0, cvtcolsstr('A,B')))
        os.system('rm ./*.csv')

    def test_tblutils_extractcols_fail_bad_file(self):
        with self.assertRaises(InvalidFileType):
            extractcols('/users/rrehders/test/fail.txt', {cvtcolsstr('A')})

    def test_tblutils_extractcols_fail_no_cols(self):
        with self.assertRaises(TypeError):
            extractcols('/users/rrehders/test/test.csv')

    def test_tblutils_extractcols_fail_empty_cols(self):
        self.assertFalse(extractcols('/users/rrehders/test/test.csv', set()))

    def test_tblutils_extractcols_success_csv(self):
        self.assertTrue(extractcols('/users/rrehders/test/test.csv', cvtcolsstr('A,B')))
        os.system('rm ./*.csv')

    def test_tblutils_extractcols_success_xlsx(self):
        self.assertTrue(extractcols('/users/rrehders/test/test.xlsx', cvtcolsstr('A,B'), 0))
        os.system('rm ./*.xlsx')
