from mailmerge import MailMerge
import sys
def readAllFields(path = ''):
    if path == '':
        sys.exit('You didn\'t select docx file. Please select docx file!')
        #raise FileNotFoundError('You didn\'t select docx file. Please select docx file!')
        
    try:
        document = MailMerge(path)
    except FileNotFoundError:
        sys.exit("Couldn't find docx file. Check file path!")
        #print("Couldn't find docx file. Check file path!")
    else:
        documentFields = document.get_merge_fields()
        if len(documentFields) == 0:
            sys.exit("Your docx file doesn\'t have any merge-fields!")
        else:
            return documentFields