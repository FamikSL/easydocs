from mailmerge import MailMerge
import PySimpleGUI as sg
import os

def merge_docs(table_data, Word_path, Save_path, new_name, index_pos):
    count = 0
    for i in table_data.index:
        sg.one_line_progress_meter('Прогресс', int(count) +1, len(table_data.index), orientation = "h")
        document = MailMerge( Word_path)
        dt=(table_data.loc[[i]].to_dict('records')[0])
        print(dt)
        document.merge(**dt)
        if index_pos == True:
            file_name = f'{Save_path}/{i}_{new_name}.docx'
        else:
            file_name = f'{Save_path}/{new_name}_{i}.docx'
        document.write(file_name)
        count+=1
