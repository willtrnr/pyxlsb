import unittest
from pyxlsb.formula.ptgs import ClassifiedPtg

class ClassifiedPtgTestCase(unittest.TestCase):
    def test_base_ptg(self):
        self.assertEqual(ClassifiedPtg(0x20).base_ptg, 0x20)
        self.assertEqual(ClassifiedPtg(0x30).base_ptg, 0x30)
        self.assertEqual(ClassifiedPtg(0x3F).base_ptg, 0x3F)
        self.assertEqual(ClassifiedPtg(0x40).base_ptg, 0x20)
        self.assertEqual(ClassifiedPtg(0x50).base_ptg, 0x30)
        self.assertEqual(ClassifiedPtg(0x5F).base_ptg, 0x3F)
        self.assertEqual(ClassifiedPtg(0x60).base_ptg, 0x20)
        self.assertEqual(ClassifiedPtg(0x70).base_ptg, 0x30)
        self.assertEqual(ClassifiedPtg(0x7F).base_ptg, 0x3F)

    def test_is_reference(self):
        self.assertTrue(ClassifiedPtg(0x20).is_reference())
        self.assertTrue(ClassifiedPtg(0x30).is_reference())
        self.assertTrue(ClassifiedPtg(0x3F).is_reference())
        self.assertFalse(ClassifiedPtg(0x40).is_reference())
        self.assertFalse(ClassifiedPtg(0x50).is_reference())
        self.assertFalse(ClassifiedPtg(0x5F).is_reference())
        self.assertFalse(ClassifiedPtg(0x60).is_reference())
        self.assertFalse(ClassifiedPtg(0x70).is_reference())
        self.assertFalse(ClassifiedPtg(0x7F).is_reference())

    def test_is_value(self):
        self.assertFalse(ClassifiedPtg(0x20).is_value())
        self.assertFalse(ClassifiedPtg(0x30).is_value())
        self.assertFalse(ClassifiedPtg(0x3F).is_value())
        self.assertTrue(ClassifiedPtg(0x40).is_value())
        self.assertTrue(ClassifiedPtg(0x50).is_value())
        self.assertTrue(ClassifiedPtg(0x5F).is_value())
        self.assertFalse(ClassifiedPtg(0x60).is_value())
        self.assertFalse(ClassifiedPtg(0x70).is_value())
        self.assertFalse(ClassifiedPtg(0x7F).is_value())

    def test_is_array(self):
        self.assertFalse(ClassifiedPtg(0x20).is_array())
        self.assertFalse(ClassifiedPtg(0x30).is_array())
        self.assertFalse(ClassifiedPtg(0x3F).is_array())
        self.assertFalse(ClassifiedPtg(0x40).is_array())
        self.assertFalse(ClassifiedPtg(0x50).is_array())
        self.assertFalse(ClassifiedPtg(0x5F).is_array())
        self.assertTrue(ClassifiedPtg(0x60).is_array())
        self.assertTrue(ClassifiedPtg(0x70).is_array())
        self.assertTrue(ClassifiedPtg(0x7F).is_array())
