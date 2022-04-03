import pandas as pd
def read_excell(path = ''):
    if path == '':
        
        raise FileNotFoundError('You didn\'t select excel file. Please select docx file!')
    
    try:
        dataframe = pd.read_excel(path, index_col=0)
    except FileNotFoundError:
        raise FileNotFoundError("Couldn't find xlsx file. Check file path!")
    else:
        if len(dataframe.columns.to_list()) == 0:
            raise ValueError("Your xlsx file doesn\'t have any columns!")
        if len(dataframe.index) == 0:
            raise ValueError("Your xlsx file doesn\'t have any rows!")
        return
