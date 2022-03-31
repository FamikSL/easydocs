import PySimpleGUI as sg
from word_reader import readAllFields
def main_window():
    sg.theme('DarkAmber')

    layout2 = [
          [sg.Text('Укажите расположение до Word-шаблона:', size=(35, 1)), sg.InputText(size=(80, 1)), sg.FileBrowse()],
          [sg.Text('Укажите расположение файла Excel:', size=(35, 1)), sg.InputText(size=(80, 1)), sg.FileBrowse()],
          [sg.Text('Укажите путь для сохранения Паспортов:', size=(35, 1)), sg.InputText(size=(80, 1)), sg.FolderBrowse()],
          [sg.Submit(), sg.Cancel()]]
    window2 = sg.Window('Меню', layout2)
    event, values = window2.read()
    window2.close()
    WORD_path, Excel_path, Save_path = values[0], values[1], values[2]
    #if (WORD_path == '' or Excel_path == '' or Save_path == ''):
    #    pass
    #else:
        #return WORD_path, Excel_path, Save_path
    return WORD_path, Excel_path, Save_path
