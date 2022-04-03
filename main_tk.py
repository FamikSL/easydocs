import PySimpleGUI as sg
from word_reader import readAllFields

sg.theme('Dark Blue 3')  # please make your windows colorful

layout = [
      [sg.Text('Укажите расположение до Word-шаблона:', size=(35, 1)), sg.InputText(size=(80, 1)), sg.FileBrowse(file_types=(("ALL", "docx doc"),))],
      [sg.Text('Укажите расположение файла Excel:', size=(35, 1)), sg.InputText(size=(80, 1)), sg.FileBrowse()],
      [sg.Text('Укажите путь для сохранения Паспортов:', size=(35, 1)), sg.InputText(size=(80, 1)), sg.FolderBrowse()],
      [sg.Button('Start'), sg.Button('Exit')]]

window = sg.Window('Window Title', layout)

while True:  # Event Loop
    event, values = window.read()
    WORD_path, Excel_path, Save_path = values[0], values[1], values[2]
    print(event, values)
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Start':
        documentFields, error = readAllFields(WORD_path)
        if error != '':
            sg.popup(error)

window.close()