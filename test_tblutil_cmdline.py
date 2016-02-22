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

    def test_two_valid_parameters_tocsv(self):
        # Use the name of the script to guarentee file is present
        # calling parse_args will return multiple args without an exception if "happy"
        self.assertTrue(self.parser.parse_args(['tocsv', 'tblutil.py']))

    def test_two_valid_parameters_extract(self):
        # Use the name of the script to guarentee file is present
        # calling parse_args will return multiple args without an exception if "happy"
        self.assertTrue(self.parser.parse_args(['extract', 'tblutil.py']))

    def test_one_valid_parameter_but_action_is_invalid(self):
        with self.assertRaises(SystemExit):
            self.parser.parse_args(['jhkjl', 'tblutil.py'])
    
    def test_optional_columns_specified_num(self):
        self.assertTrue(self.parser.parse_args(['tocsv', 'tblitil.py', '--cols=1,2']))

    def test_optional_columns_specified_alpha(self):
        self.assertTrue(self.parser.parse_args(['tocsv', 'tblitil.py', '--cols=A,B']))
