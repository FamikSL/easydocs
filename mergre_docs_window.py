import PySimpleGUI as sg
from excel_reader import read_excell
from word_reader import readAllFields
from Merge_Docs import merge_docs

def merge_window():
      sg.theme('Dark Blue 3')

      layout = [
            [sg.Text('Укажите расположение до Word-шаблона:', size=(35, 1), font='Roboto'), sg.InputText(size=(90, 1), font='Roboto 10'), sg.FileBrowse(font='Roboto',file_types=(("Docs", "docx doc"), ("All", "*")))],
            [sg.Text('Укажите расположение файла Excel:', size=(35, 1), font='Roboto'), sg.InputText(size=(90, 1),font='Roboto 10'), sg.FileBrowse(font='Roboto',file_types=(("Sheets", "xlsx xlsm"), ("All", "*")))],
            [sg.Text('Укажите путь для сохранения Паспортов:', size=(35, 1), font='Roboto'), sg.InputText(size=(90, 1), font='Roboto 10'), sg.FolderBrowse(font='Roboto')],
            [sg.Text('Укажите имя для файлов:', size=(35, 1), font='Roboto'), sg.InputText(size=(30, 1), font='Roboto'), sg.Radio('Номер перед названия','index', key ='index_before', default = 'True'), sg.Radio('Номер после названия','index', key='index_after')
 ],
            [sg.Button('Старт', font='Roboto'), sg.Button('Выйти', font='Roboto')]]
      window = sg.Window('easydocs', layout, icon='brand/icon.ico')

      while True:  
            event, values = window.read()
            print(values)

            for parameter in values.values():
                  # print(parameter)
                  fields_have_errors = False
                  if parameter == '':
                        fields_have_errors = True
                        
                        
            if fields_have_errors:
                  sg.popup('Ошибка при обработке введенных данных', title='Предупреждение',font='Roboto')
                  continue
            WORD_path, Excel_path, Save_path, Doc_name, index_before, index_after = values[0], values[1], values[2], values[3], values['index_before'], values['index_after']
            
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
                        merge_docs(table_data, WORD_path, Save_path, Doc_name, index_before)
                  
                  
                     
      window.close()