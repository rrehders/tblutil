from tblutil import create_parser
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

    def test_single_parameter_tocsv(self):
        with self.assertRaises(SystemExit):
            self.parser.parse_args(['tocsv'])

    def test_single_parameter_extract(self):
        with self.assertRaises(SystemExit):
            self.parser.parse_args(['extract'])

    def test_single_parameter_join(self):
        with self.assertRaises(SystemExit):
            self.parser.parse_args(['join'])

    def test_single_parameter_invalid(self):
        with self.assertRaises(SystemExit):
            self.parser.parse_args(['jhkl'])

    def test_invalid_action_but_minimum_paramenters(self):
        with self.assertRaises(SystemExit):
            self.parser.parse_args(['jhkjl', 'tblutil.py'])

    def test_two_valid_parameters_tocsv(self):
        # Use the name of the script to guarentee file is present
        args = self.parser.parse_args(['tocsv', 'tblutil.py'])
        self.assertTrue(args.action == 'tocsv')
        self.assertTrue(args.file == 'tblutil.py')
        self.assertTrue(args.sheet is None)
        self.assertTrue(args.cols is None)

    def test_valid_tocsv_w_optional_sheet(self):
        # Use the name of the script to guarentee file is present
        args = self.parser.parse_args(['tocsv', 'tblutil.py', '--sheet=1'])
        self.assertTrue(args.action == 'tocsv')
        self.assertTrue(args.file == 'tblutil.py')
        self.assertTrue(args.sheet == 1)
        self.assertTrue(args.cols is None)

    def test_valid_tocsv_optional_columns_specified_num(self):
        args = self.parser.parse_args(['tocsv', 'tblutil.py', '--cols=1,2'])
        self.assertTrue(args.action == 'tocsv')
        self.assertTrue(args.file == 'tblutil.py')
        self.assertTrue(args.sheet is None)
        self.assertTrue(args.cols == '1,2')

    def test_valid_tocsv_sheet_optional_columns_specified_num(self):
        args = self.parser.parse_args(['tocsv', 'tblutil.py', '--sheet=0', '--cols=1,2'])
        self.assertTrue(args.action == 'tocsv')
        self.assertTrue(args.file == 'tblutil.py')
        self.assertTrue(args.sheet == 0)
        self.assertTrue(args.cols == '1,2')

    def test_valid_tocsv_sheet_optional_columns_specified_alpha(self):
        args = self.parser.parse_args(['tocsv', 'tblutil.py', '--sheet=0', '--cols=A,B'])
        self.assertTrue(args.action == 'tocsv')
        self.assertTrue(args.file == 'tblutil.py')
        self.assertTrue(args.sheet == 0)
        self.assertTrue(args.cols == 'A,B')

    def test_valid_join(self):
        args = self.parser.parse_args(['join', 'tblutil.py'])
        self.assertTrue(args.action == 'join')
        self.assertTrue(args.file == 'tblutil.py')
        self.assertTrue(args.to is None)
        self.assertTrue(args.index is None)

    def test_valid_join_with_specified(self):
        args = self.parser.parse_args(['join', 'tblutil.py', '--to=tblutil.py'])
        self.assertTrue(args.action == 'join')
        self.assertTrue(args.file == 'tblutil.py')
        self.assertTrue(args.to == 'tblutil.py')
        self.assertTrue(args.index is None)

    def test_valid_join_with_and_numeric_index(self):
        args = self.parser.parse_args(['join', 'tblutil.py', '--to=tblutil.py', '--index=1'])
        self.assertTrue(args.action == 'join')
        self.assertTrue(args.file == 'tblutil.py')
        self.assertTrue(args.to == 'tblutil.py')
        self.assertTrue(args.index == '1')

    def test_valid_join_with_alpa_index(self):
        args = self.parser.parse_args(['join', 'tblutil.py', '--to=tblutil.py', '--index=A'])
        self.assertTrue(args.action == 'join')
        self.assertTrue(args.file == 'tblutil.py')
        self.assertTrue(args.to == 'tblutil.py')
        self.assertTrue(args.index == 'A')

    def test_valid_join_with_2_numeric_index(self):
        args = self.parser.parse_args(['join', 'tblutil.py', '--to=tblutil.py', '--index=3,2'])
        self.assertTrue(args.action == 'join')
        self.assertTrue(args.file == 'tblutil.py')
        self.assertTrue(args.to == 'tblutil.py')
        self.assertTrue(args.index == '3,2')

    def test_valid_join_with_2_alpha_index(self):
        args = self.parser.parse_args(['join', 'tblutil.py', '--to=tblutil.py', '--index=C,B'])
        self.assertTrue(args.action == 'join')
        self.assertTrue(args.file == 'tblutil.py')
        self.assertTrue(args.to == 'tblutil.py')
        self.assertTrue(args.index == 'C,B')
