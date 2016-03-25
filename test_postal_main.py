from cmnfns import InvalidFileType, InvalidExcelColumn
from postal import cleanpostal
import unittest
import os


class MainTestCases(unittest.TestCase):
    def test_cleanpostal_fail(self):
        self.assertFalse(cleanpostal('/users/rrehders/test/test.xls', '0'))

    def test_cleanpostal_success_minimum_parameters(self):
        self.assertTrue(cleanpostal('/users/rrehders/test/test.xlsx'))
        os.system('rm ./*.csv')

    def test_tblutil_cleanpostal_success_xlsx_sheet_specified(self):
        self.assertTrue(cleanpostal('/users/rrehders/test/test.xlsx', '0', '0'))
        os.system('rm ./*.csv')

    def test_tblutil_cleanpostal_success_xlsx_sheet_not_specified(self):
        self.assertTrue(cleanpostal('/users/rrehders/test/test.xlsx', '0'))
        os.system('rm ./*.csv')

    def test_tblutil_cleanpostal_success_csv(self):
        self.assertTrue(cleanpostal('/users/rrehders/test/test.csv', '0',))
        os.system('rm ./*.csv')

    def test_tblutil_cleanpostal_success_extra_paramater(self):
        self.assertTrue(cleanpostal('/users/rrehders/test/test.csv', '0', '0'))
        os.system('rm ./*.csv')
