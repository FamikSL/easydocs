import PySimpleGUI as sg
from word_reader import readAllFields

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
            print(event, values)
            if event == sg.WIN_CLOSED or event == 'Exit':
                  break
            if event == 'Start':
                  documentFields, error = readAllFields(WORD_path)
                  if error != '':
                        sg.popup(error, title='Предупреждение',font='Roboto')   
      window.close()