from tocsv import tocsv
import unittest
import os


class tocsvTestCases(unittest.TestCase):
    def test_tocsv_fail(self):
        self.assertFalse(tocsv('/users/rrehders/test/test.xls', 0))

    def test_tocsv_success_minimum_parameters(self):
        self.assertTrue(tocsv('/users/rrehders/test/test.xlsx'))
        self.assertTrue(os.path.exists('./test-2016 Test Data.csv'))
        self.assertTrue(os.path.exists('./test-2016 Sample Data.csv'))
        os.system('rm ./*.csv')

    def test_tocsv_success_one_sheet(self):
        self.assertTrue(tocsv('/users/rrehders/test/test.xlsx', 0))
        self.assertTrue(os.path.exists('./test-2016 Test Data.csv'))
        os.system('rm ./*.csv')

    def test_tocsv_success_num_parameters(self):
        self.assertTrue(tocsv('/users/rrehders/test/test.xlsx', 0, '1,2'))
        self.assertTrue(os.path.exists('./test-2016 Test Data.csv'))
        os.system('rm ./*.csv')

    def test_tocsv_success_alpha_parameters(self):
        self.assertTrue(tocsv('/users/rrehders/test/test.xlsx', 0, 'A,B'))
        self.assertTrue(os.path.exists('./test-2016 Test Data.csv'))
        os.system('rm ./*.csv')
