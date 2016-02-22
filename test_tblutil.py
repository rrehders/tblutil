from tblutil import valid_file, whatfiletype, tocsv, extractcols
import unittest


class SubFunctionsTestCase(unittest.TestCase):
    def test_valid_file(self):
        self.assertTrue(valid_file(['tblutil.py']))

    def test_fail_valid_file(self):
        self.assertFalse(valid_file(['tblutel.py']))

    def test_what_file_type_XLS(self):
        self.assertIs(whatfiletype('Test.xls'), 'excel')

    def test_what_file_type_XLSX(self):
        self.assertIs(whatfiletype('Test.xlsx'), 'excel')

    def test_what_file_type_CSV(self):
        self.assertIs(whatfiletype('Test.csv'), 'csv')

    def test_what_file_type_other(self):
        self.assertFalse(whatfiletype('Test.csv'))

    def test_tocsv_invalid_file_type(self):
        with self.assertRaises(SystemExit):
            tocsv('Test.csv')

    def test_extractcols_invalid_file_type(self):
        with self.assertRaises(SystemExit):
            extractcols('Test.xls')
