from tocsv import create_parser
import unittest


class CommandLineTestCases(unittest.TestCase):
    def setUp(self):
        self.parser = create_parser()


    def test_no_args(self):
        with self.assertRaises(SystemExit):
            self.parser.parse_args([])

    def test_help_request(self):
        with self.assertRaises(SystemExit):
                self.parser.parse_args(['-h'])

    def test_minimum_paramenters(self):
        # Use the name of the script to guarentee file is present
        args = self.parser.parse_args(['cmnfns.py'])
        self.assertTrue(args.file == 'cmnfns.py')
        self.assertTrue(args.sheet == -1)
        self.assertTrue(args.cols is None)

    def test_valid_tocsv_w_optional_sheet(self):
        # Use the name of the script to guarentee file is present
        args = self.parser.parse_args(['cmnfns.py', '--sheet=1'])
        self.assertTrue(args.file == 'cmnfns.py')
        self.assertTrue(args.sheet == 1)
        self.assertTrue(args.cols is None)

    def test_valid_tocsv_optional_columns_specified_num(self):
        args = self.parser.parse_args(['cmnfns.py', '--cols=1,2'])
        self.assertTrue(args.file == 'cmnfns.py')
        self.assertTrue(args.sheet == -1)
        self.assertTrue(args.cols == '1,2')

    def test_valid_tocsv_sheet_optional_columns_specified_num(self):
        args = self.parser.parse_args(['cmnfns.py', '--sheet=0', '--cols=1,2'])
        self.assertTrue(args.file == 'cmnfns.py')
        self.assertTrue(args.sheet == 0)
        self.assertTrue(args.cols == '1,2')

    def test_valid_tocsv_sheet_optional_columns_specified_alpha(self):
        args = self.parser.parse_args(['cmnfns.py', '--sheet=0', '--cols=A,B'])
        self.assertTrue(args.file == 'cmnfns.py')
        self.assertTrue(args.sheet == 0)
        self.assertTrue(args.cols == 'A,B')
