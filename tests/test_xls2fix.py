from excel2fixture import xls2fix
import sys
import os

class test_xls_to_yaml(object):
    def setUp(self):
        sys.argv = ['', 'xls/Simple.xls', '-y', 'yaml/simple.yaml', '-o', 'output/Simple.json']

    def test_method(self):
        xls2fix.main()

if __name__ == '__main__':
    unittest.main()

