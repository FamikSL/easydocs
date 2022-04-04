import pandas as pd
def read_excell(path = ''):
    error = ''
    if path == '':
        error = 'You didn\'t select excel file. Please select docx file!'
        return None, error 
    try:
        dataframe = pd.read_excel(path, index_col=0, dtype=str)
        dataframe.fillna('', inplace=True)
    except FileNotFoundError:
        error = "Couldn't find xlsx file. Check file path!"
        return None, error 
    else:
        if len(dataframe.columns.to_list()) == 0:
            error = "Your xlsx file doesn\'t have any columns!"
            return None, error 
        if len(dataframe.index) == 0:
            error = "Your xlsx file doesn\'t have any rows!"
            return None, error 
        return dataframe, error
