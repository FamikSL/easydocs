from word_reader import readAllFields
import pandas as pd
import os
import PySimpleGUI as sg

def print_fields(Word_path, Save_path):
    new_name = os.path.basename(Word_path).split('.')[0]
    documentFields = readAllFields(Word_path)[0]
    df = pd.DataFrame([['']* len(list(documentFields))],columns=list(documentFields))
    df.index+=1
    df.to_excel(f'{Save_path}/{new_name}.xlsx')
    sg.popup('Excel создан', title='Сообщение',font='Roboto')