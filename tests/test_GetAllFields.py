import unittest

from Word_Reader import readAllFields

class TestAssertions(unittest.TestCase):

    def test_empty_path(self):
        with self.assertRaises(FileNotFoundError) as context:
            readAllFields()
        self.assertTrue("You didn\'t select docx file. Please select docx file!" in str(context.exception))
    
    def test_bad_path(self):
        with self.assertRaises(FileNotFoundError) as context:
            readAllFields('BAD_FILE_NAME.docx')
        self.assertTrue("Couldn't find docx file. Check file path!" in str(context.exception))
    
    def test_read_merge_fields(self):
        self.assertEqual(readAllFields(r'sample_docs/Template(test1).docx'), {'Поле_1','Поле_2'})

    def test_zero_fields(self):
        with self.assertRaises(ValueError) as context:
            readAllFields(r'sample_docs/Template(test2).docx')
        self.assertTrue("Your docx file doesn\'t have any merge-fields!" in str(context.exception))
    


if __name__ == '__main__':
    unittest.main()