from postal import cleanpostal
import unittest
import os


class MainTestCases(unittest.TestCase):
    def test_cleanpostal_fail(self):
        self.assertIsNone(cleanpostal('/users/rrehders/test/test.xls', '0'))

    def test_cleanpostal_success_xlsx_alpha_no_sheet(self):
        self.assertTrue(cleanpostal('/users/rrehders/test/addrtest.xlsx', 'V'))
        os.system('rm ./*.csv')

    def test_tblutil_cleanpostal_success_xlsx_num_sheet_specified(self):
        self.assertTrue(cleanpostal('/users/rrehders/test/addrtest.xlsx', '22', '0'))
        os.system('rm ./*.csv')

    def test_tblutil_cleanpostal_success_csv_alpha(self):
        self.assertTrue(cleanpostal('/users/rrehders/test/addrtest.csv', 'V'))
        # os.system('rm ./*.csv')

    def test_tblutil_cleanpostal_success_csv_num(self):
        self.assertTrue(cleanpostal('/users/rrehders/test/addrtest.csv', '22'))
        os.system('rm ./*.csv')

    def test_tblutil_cleanpostal_success_csv_alpha_extra_paramater(self):
        self.assertTrue(cleanpostal('/users/rrehders/test/addrtest.csv', 'V', '0'))
        os.system('rm ./*.csv')
