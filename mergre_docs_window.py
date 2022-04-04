import PySimpleGUI as sg
from Excel_Reader import read_excell
from Word_Reader import readAllFields
from Merge_Docs import merge_docs

def merge_window():
      sg.theme('Dark Blue 3')

      layout = [
            [sg.Text('Укажите расположение до Word-шаблона:', size=(35, 1), font='Roboto'), sg.InputText(size=(80, 1), font='Roboto'), sg.FileBrowse(font='Roboto',file_types=(("ALL", "docx doc"),))],
            [sg.Text('Укажите расположение файла Excel:', size=(35, 1), font='Roboto'), sg.InputText(size=(80, 1),font='Roboto'), sg.FileBrowse(font='Roboto')],
            [sg.Text('Укажите путь для сохранения Паспортов:', size=(35, 1), font='Roboto'), sg.InputText(size=(80, 1), font='Roboto'), sg.FolderBrowse(font='Roboto')],
            [sg.Button('Start', font='Roboto'), sg.Button('Exit', font='Roboto')]]

      window = sg.Window('easydocs', layout, icon='brand/icon.ico')

      while True:  
            event, values = window.read()
            WORD_path, Excel_path, Save_path = values[0], values[1], values[2]

            if event == sg.WIN_CLOSED or event == 'Exit':
                  break
            if event == 'Start':
                  documentFields, error_1 = readAllFields(WORD_path)
                  if error_1 != '':
                        sg.popup(error_1, title='Предупреждение',font='Roboto')
                  table_data, error_2 = read_excell(Excel_path)
                  if error_2 != '':
                        sg.popup(error_2, title='Предупреждение',font='Roboto')
                  if error_1 == '' and error_2 == '':
                        merge_docs(table_data, WORD_path, Save_path)
                  
                  
                     
      window.close()