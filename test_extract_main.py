from cmnfns import InvalidFileType
from extract import extractcols
import unittest
import os


class extractTestCases(unittest.TestCase):
    def test_extractcols_fail_bad_file(self):
        with self.assertRaises(InvalidFileType):
            extractcols('/users/rrehders/test/fail.txt', 'A')

    def test_extractcols_fail_no_cols(self):
        self.assertFalse(extractcols('/users/rrehders/test/test.csv', '', '0'))

    def test_extractcols_success_minimum_parameters_num_csv_first_sheet(self):
        self.assertTrue(extractcols('/users/rrehders/test/test.csv', '1,2', '0'))
        os.system('rm ./*.csv')

    def test_extractcols_success_minimum_parameters_alpha_csv_first_sheet(self):
        self.assertTrue(extractcols('/users/rrehders/test/test.csv', 'A,B', '0'))
        os.system('rm ./*.csv')

    def test_extractcols_success_minimum_parameters_num_csv_second_sheet(self):
        self.assertTrue(extractcols('/users/rrehders/test/test.csv', '1,2', '1'))
        os.system('rm ./*.csv')

    def test_extractcols_success_minimum_parameters_alpha_csv_second_sheet(self):
        self.assertTrue(extractcols('/users/rrehders/test/test.csv', 'A,B', '1'))
        os.system('rm ./*.csv')

    def test_extractcols_success_minimum_parameters_num_excel(self):
        self.assertTrue(extractcols('/users/rrehders/test/test.xlsx', '1,2'))
        os.system('rm ./*.csv')

    def test_extractcols_success_minimum_parameters_alpha_excel(self):
        self.assertTrue(extractcols('/users/rrehders/test/test.xlsx', 'A,B'))
        os.system('rm ./*.csv')
