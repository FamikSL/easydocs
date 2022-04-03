from mailmerge import MailMerge
import sys
import PySimpleGUI as sg
def readAllFields(path = ''):
    error = ''
    if path == '':
        
        error = 'You didn\'t select docx file. Please select docx file!'
        return (),error
    try:
        document = MailMerge(path)
    except FileNotFoundError:
        error = "Couldn't find docx file. Check file path!"
        return (),error
    else:
        documentFields = document.get_merge_fields()
        if len(documentFields) == 0:
            error="Your docx file doesn\'t have any merge-fields!"
            return (), error
        else:
            return documentFields, error