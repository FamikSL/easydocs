from mailmerge import MailMerge
def readAllFields(path = ''):
    if path == '':
        #sys.exit('You didn\'t select docx file. Please select docx file!')
        raise FileNotFoundError('You didn\'t select docx file. Please select docx file!')
        
    try:
        document = MailMerge(path)
    except FileNotFoundError:
        #sys.exit("Couldn't find docx file. Check file path!")
        raise FileNotFoundError("Couldn't find docx file. Check file path!")
    else:
        documentFields = document.get_merge_fields()
        if len(documentFields) == 0:
            raise ValueError("Your docx file doesn\'t have any merge-fields!")
        else:
            return documentFields