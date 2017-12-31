import unittest
import os.path
from pyxlsb.styles import Styles

class StylesTestCase(unittest.TestCase):
    def setUp(self):
        self.styles = Styles(open(os.path.join('test_files', 'parts', 'styles.bin'), 'rb'))

    def tearDown(self):
        self.styles.close()

    @unittest.skip('TODO')
    def test_get_style(self):
        pass
