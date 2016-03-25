from postal import create_parser
import unittest


class CommandLineTestCases(unittest.TestCase):
    def setUp(self):
        self.parser = create_parser()

    def test_postalcmdline_no_args(self):
        with self.assertRaises(SystemExit):
            self.parser.parse_args([])

    def test_postalcmdline_help_request(self):
        with self.assertRaises(SystemExit):
                self.parser.parse_args(['-h'])

    def test_postalcmdline_no_cols(self):
        # Use the name of the script to guarantee file is present
        with self.assertRaises(SystemExit):
            self.parser.parse_args(['cmnfns.py'])

    def test_postalcmdline_valid_colnum(self):
        # Use the name of the script to guarentee file is present
        args = self.parser.parse_args(['cmnfns.py', '1'])
        self.assertTrue(args.file == 'cmnfns.py')
        self.assertTrue(args.col == '1')
        self.assertTrue(args.sheet == 0)

    def test_postalcmdline_valid_colalpha(self):
        # Use the name of the script to guarentee file is present
        args = self.parser.parse_args(['cmnfns.py', 'A'])
        self.assertTrue(args.file == 'cmnfns.py')
        self.assertTrue(args.col == 'A')
        self.assertTrue(args.sheet == 0)

    def test_postalcmdline_valid_colnum_optional_sheet(self):
        args = self.parser.parse_args(['cmnfns.py', '1', '--sheet=1'])
        self.assertTrue(args.file == 'cmnfns.py')
        self.assertTrue(args.col == '1')
        self.assertTrue(args.sheet == 1)

    def test_postalcmdline_valid_colalpha_optional_sheet(self):
        args = self.parser.parse_args(['cmnfns.py', 'A', '--sheet=1'])
        self.assertTrue(args.file == 'cmnfns.py')
        self.assertTrue(args.col == 'A')
        self.assertTrue(args.sheet == 1)
