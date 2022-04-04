import PySimpleGUI as sg
from excel_from_word import print_fields
import mergre_docs_window

def window_word_to_excel():
    sg.theme('Dark Blue 3')

    layout = [
          [sg.Text('Укажите расположение до Word-шаблона:', size=(35, 1), font='Roboto'), sg.InputText(size=(90, 1), font='Roboto 10'), sg.FileBrowse(font='Roboto',file_types=(("ALL", "docx doc"),))],
          [sg.Text('Укажите путь для сохранения Паспортов:', size=(35, 1), font='Roboto'), sg.InputText(size=(90, 1), font='Roboto 10'), sg.FolderBrowse(font='Roboto')],
          [sg.Button('Старт', font='Roboto'), sg.Button('Выйти', font='Roboto')]]

    main_window = sg.Window('easydocs', layout, icon='brand/icon.ico')

    while True:  
        event, values = main_window.read()
        Word_path, Save_path = values[0], values[1],
        if event == sg.WIN_CLOSED or event == 'Выйти':
              break
        if event == 'Старт':
            print_fields(Word_path, Save_path)
        
            

    main_window.close()