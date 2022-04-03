import PySimpleGUI as sg
import mergre_docs_window
sg.theme('Dark Blue 3')

layout = [      
      [sg.Button('Создать документы по шаблону', font='Roboto')],
      [sg.Button('Выгрузить поля документа в таблицу', font='Roboto')]]

main_window = sg.Window('easydocs', layout, icon='brand/icon.ico')

while True:  
      event, values = main_window.read()
      if event == sg.WIN_CLOSED:
            break
      if event == 'Создать документы по шаблону':
          mergre_docs_window.merge_window()
          main_window.close()

main_window.close()