from mailmerge import MailMerge
import PySimpleGUI as sg
def readAllFields(path = ''):
    error = ''
    if path == '':
        error = 'Файл Word не выбран!'
        return (),error
    try:
        document = MailMerge(path)
    except FileNotFoundError:
        error = "Не могу найти docx файл! Проверьте путь до файла."
        return (),error
    else:
        documentFields = document.get_merge_fields()
        if len(documentFields) == 0:
            error="В docx файле нет полей для замены(merge-fields)!"
            return (), error
        else:
            return documentFields, error