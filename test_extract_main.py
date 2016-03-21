from cmnfns import InvalidFileType, InvalidExcelColumn, getfiletype, cvtcolsstrtoset
from extract import extractxltable, extractlisttable, extractcols
import unittest
import os


class MainTestCases(unittest.TestCase):
    def test_tblutil_extractcols_fail(self):
        self.assertFalse(extractcols('/users/rrehders/test/test.xls', 0))

    def test_tblutil_extractcols_success_minimum_parameters_excel(self):
        self.assertTrue(extractcols('/users/rrehders/test/test.xlsx'))
        os.system('rm ./*.csv')

    def test_tblutil_extractcols_success_minimum_parameters_csv(self):
        self.assertTrue(extractcols('/users/rrehders/test/test.csv'))
        os.system('rm ./*.csv')

    def test_tblutil_extractcols_success_one_sheet(self):
        self.assertTrue(extractcols('/users/rrehders/test/test.xlsx', 0))
        os.system('rm ./*.csv')

    def test_tblutil_tocsv_success_parameters(self):
        self.assertTrue(extractcols('/users/rrehders/test/test.xlsx', 0, 'A,B'))
        os.system('rm ./*.csv')

    def test_extractcols_fail_bad_file(self):
        with self.assertRaises(InvalidFileType):
            extractcols('/users/rrehders/test/fail.txt', 'A')

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
