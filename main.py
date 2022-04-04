import PySimpleGUI as sg
import excel_from_word_window
import mergre_docs_window
def main():
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
            main_window.close()
            mergre_docs_window.merge_window()
        if event == 'Выгрузить поля документа в таблицу':
            main_window.close()
            excel_from_word_window.window_word_to_excel()
        
            

    main_window.close()
main()