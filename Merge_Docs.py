from mailmerge import MailMerge
import PySimpleGUI as sg
import os

def merge_docs(table_data, Word_path, Save_path):
    new_name = os.path.basename(Word_path).split('.')[0]
    count = 0
    for i in table_data.index:
        sg.one_line_progress_meter('My Meter', int(count), len(table_data.index) - 1, 'key','Optional message')
        document = MailMerge( Word_path)
        dt=(table_data.loc[[i]].to_dict('records')[0])
        document.merge(**dt)
        document.write(f'{Save_path}/{i}_{new_name}.docx')
        count+=1
