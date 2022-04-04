import PySimpleGUI as sg
from excel_reader import read_excell
from word_reader import readAllFields
from Merge_Docs import merge_docs

def merge_window():
      sg.theme('Dark Blue 3')

      layout = [
            [sg.Text('Укажите расположение до Word-шаблона:', size=(35, 1), font='Roboto'), sg.InputText(size=(90, 1), font='Roboto 10'), sg.FileBrowse(font='Roboto',file_types=(("Docs", "docx doc"), ("All", "*")))],
            [sg.Text('Укажите расположение файла Excel:', size=(35, 1), font='Roboto'), sg.InputText(size=(90, 1),font='Roboto 10'), sg.FileBrowse(font='Roboto',file_types=(("Sheets", "xlsx"), ("All", "*")))],
            [sg.Text('Укажите путь для сохранения Паспортов:', size=(35, 1), font='Roboto'), sg.InputText(size=(90, 1), font='Roboto 10'), sg.FolderBrowse(font='Roboto')],
            [sg.Button('Старт', font='Roboto'), sg.Button('Выйти', font='Roboto')]]

      window = sg.Window('easydocs', layout, icon='brand/icon.ico')

      while True:  
            event, values = window.read()
            WORD_path, Excel_path, Save_path = values[0], values[1], values[2]

                  
            if event == sg.WIN_CLOSED or event == 'Выйти':
                  break
            if event == 'Старт':
                  documentFields, error_1 = readAllFields(WORD_path)
                  if error_1 != '':
                        sg.popup(error_1, title='Предупреждение',font='Roboto')
                  table_data, error_2 = read_excell(Excel_path)
                  if error_2 != '':
                        sg.popup(error_2, title='Предупреждение',font='Roboto')
                  if error_1 == '' and error_2 == '':
                        merge_docs(table_data, WORD_path, Save_path)
                  
                  
                     
      window.close()