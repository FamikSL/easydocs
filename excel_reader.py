import pandas as pd
def read_excell(path = ''):
    error = ''
    if path == '':
        error = 'Excel файл не выбран!'
        return None, error 
    try:
        dataframe = pd.read_excel(path, index_col=0, dtype=str)
        dataframe.fillna('', inplace=True)
    except FileNotFoundError:
        error = "Не могу найти xlsx файл! Проверьте путь до файла."
        return None, error 
    else:
        if len(dataframe.columns.to_list()) == 0:
            error = "Xlsx файл - пуст!"
            return None, error 
        if len(dataframe.index) == 0:
            error = "В Xlsx файле нет данных для заполнения!"
            return None, error 
        return dataframe, error
