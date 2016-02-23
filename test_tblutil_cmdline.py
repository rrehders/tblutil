from tblutil import create_parser
import unittest


class CommandLineTestCase(unittest.TestCase):
    def setUp(self):
        self.parser = create_parser()

    def test_no_args(self):
        with self.assertRaises(SystemExit):
            self.parser.parse_args([])

    def test_help_request(self):
        with self.assertRaises(SystemExit):
                self.parser.parse_args(['-h'])

    def test_single_parameter(self):
        with self.assertRaises(SystemExit):
            self.parser.parse_args(['tocsv'])

    def test_one_valid_parameter_but_action_is_invalid(self):
        with self.assertRaises(SystemExit):
            self.parser.parse_args(['jhkjl', 'tblutil.py'])

    def test_two_valid_parameters_tocsv(self):
        # Use the name of the script to guarentee file is present
        args=self.parser.parse_args(['tocsv', 'tblutil.py'])
        self.assertTrue(args.action=='tocsv')
        self.assertTrue(args.file=='tblutil.py')
        self.assertTrue(args.sheet==None)
        self.assertTrue(args.cols==None)

    def test_valid_tocsv_w_optional_sheet(self):
        # Use the name of the script to guarentee file is present
        args=self.parser.parse_args(['tocsv', 'tblutil.py', '-sheet=1'])
        self.assertTrue(args.action=='tocsv')
        self.assertTrue(args.file=='tblutil.py')
        self.assertTrue(args.sheet==1)
        self.assertTrue(args.cols==None)

    def test_optional_columns_specified_num(self):
        args=self.parser.parse_args(['tocsv', 'tblutil.py', '-cols=1,2'])
        self.assertTrue(args.action=='tocsv')
        self.assertTrue(args.file=='tblutil.py')
        self.assertTrue(args.sheet==None)
        self.assertTrue(args.cols=='1,2')

    def test_optional_columns_specified_alpha(self):
        args=self.parser.parse_args(['tocsv', 'tblutil.py', '-cols=A,B'])
        self.assertTrue(args.action=='tocsv')
        self.assertTrue(args.file=='tblutil.py')
        self.assertTrue(args.sheet==None)
        self.assertTrue(args.cols=='A,B')
