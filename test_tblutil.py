from tblutil import InvalidFileType, InvalidExcelColumn
from tblutil import isValidFile, getFileType, cvtColsStrToSet
import unittest


class SubFunctionsTestCase(unittest.TestCase):
    def test_tblutil_isValidFile(self):
        self.assertTrue(isValidFile('tblutil.py'))

    def test_fail_tblutil_isValidFile(self):
        self.assertFalse(isValidFile('tblutel.py'))

    def test_tblutil_getFileType_XLS(self):
        self.assertTrue(getFileType('Test.xls') is 'excel')

    def test_tblutil_getFileType_XLSX(self):
        self.assertTrue(getFileType('Test.xlsx')=='excel')

    def test_tblutil_getFileType_CSV(self):
        self.assertTrue(getFileType('Test.csv')=='csv')

    def test_tblutil_getFileType_other(self):
        with self.assertRaises(InvalidFileType):
            getFileType('Test.txt')

    def test_tblutil_cvtColsStrToSet_err(self):
        with self.assertRaises(InvalidExcelColumn):
            cvtColsStrToSet('A1')

    def test_tblutil_cvtColsStrToSet_num(self):
        result= cvtColsStrToSet('1,3,7')
        self.assertTrue(result-{1,2,3})

    def test_tblutil_cvtColsStrToSet_alpha(self):
        result=cvtColsStrToSet('A,C,G')
        self.assertTrue(result-{1,2,3})
