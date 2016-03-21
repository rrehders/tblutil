from extract import create_parser
import unittest


class CommandLineTestCases(unittest.TestCase):
    def setUp(self):
        self.parser = create_parser()

    def test_extractcmdline_no_args(self):
        with self.assertRaises(SystemExit):
            self.parser.parse_args([])

    def test_extractcmdline_help_request(self):
        with self.assertRaises(SystemExit):
                self.parser.parse_args(['-h'])

    def test_extractcmdline_no_cols(self):
        # Use the name of the script to guarantee file is present
        with self.assertRaises(SystemExit):
            self.parser.parse_args(['cmnfns.py'])

    def test_extractcmdline_valid_colnum(self):
        # Use the name of the script to guarentee file is present
        args = self.parser.parse_args(['cmnfns.py', '1,3'])
        self.assertTrue(args.file == 'cmnfns.py')
        self.assertTrue(args.cols == '1,3')
        self.assertTrue(args.sheet == 0)

    def test_extractcmdline_valid_colalpha(self):
        # Use the name of the script to guarentee file is present
        args = self.parser.parse_args(['cmnfns.py', 'A,C'])
        self.assertTrue(args.file == 'cmnfns.py')
        self.assertTrue(args.cols == 'A,C')
        self.assertTrue(args.sheet == 0)

    def test_extractcmdline_valid_colnum_optional_sheet(self):
        args = self.parser.parse_args(['cmnfns.py', '1,2', '--sheet=1'])
        self.assertTrue(args.file == 'cmnfns.py')
        self.assertTrue(args.cols == '1,2')
        self.assertTrue(args.sheet == 1)

    def test_extractcmdline_valid_colalpha_optional_sheet(self):
        args = self.parser.parse_args(['cmnfns.py', '1,2', '--sheet=1'])
        self.assertTrue(args.file == 'cmnfns.py')
        self.assertTrue(args.cols == '1,2')
        self.assertTrue(args.sheet == 1)

