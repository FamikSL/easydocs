from Word_Reader import readAllFields
import pandas as pd
import os

def print_fields(Word_path, Save_path):
    new_name = os.path.basename(Word_path).split('.')[0]
    documentFields = readAllFields(Word_path)[0]
    df = pd.DataFrame([['']* len(list(documentFields))],columns=list(documentFields))
    df.index+=1
    df.to_excel(f'{Save_path}/{new_name}.xlsx')  