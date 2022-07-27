from mailmerge import MailMerge
import PySimpleGUI as sg
import os

def merge_docs(table_data, Word_path, Save_path, new_name):
    count = 0
    for i in table_data.index:
        sg.one_line_progress_meter('Прогресс', int(count) +1, len(table_data.index), orientation = "h")
        document = MailMerge( Word_path)
        dt=(table_data.loc[[i]].to_dict('records')[0])
        print(dt)
        document.merge(**dt)
        document.write(f'{Save_path}/{i}_{new_name}.docx')
        count+=1
