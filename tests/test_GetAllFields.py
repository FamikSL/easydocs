import unittest

from GetAllFields import readAllFields

class TestAssertions(unittest.TestCase):

    def test_zero_path(self):
        with self.assertRaises(FileNotFoundError) as context:
            readAllFields()
        self.assertTrue("You didn\'t select docx file. Please select docx file!" in str(context.exception))
    
    def test_bad_path(self):
        with self.assertRaises(FileNotFoundError) as context:
            readAllFields('BAD_FILE_NAME.docx')
        self.assertTrue("Couldn't find docx file. Check file path!" in str(context.exception))
    def test_read_merge_fields(self):
        self.assertEqual(readAllFields(r'tests/Template(test).docx'), {'Поле_1','Поле_2'})
if __name__ == '__main__':
    unittest.main()