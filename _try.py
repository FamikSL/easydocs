from mergre_docs_window import main_window
from dataclasses import Field
from pydoc import doc
from word_reader import readAllFields
from mailmerge import MailMerge
import pandas as pd

Word_path, Excel_path, Save_path = main_window()
Word_path = ''
Excel_path = 'C:/Users/Vadim/Documents/GitHub/easydocs/sample_docs/input.xlsx'
df1 = pd.read_excel(Excel_path, index_col=0)

word_fields = readAllFields(Word_path)
exel_fields = set(df1.columns.to_list())
print(word_fields)
if (word_fields.symmetric_difference(exel_fields) == set() ):
    print(1)


# for i in df1.index:
#     document = MailMerge(WORD_path)
#     dt=(df1.loc[[i]].to_dict('records')[0])
#     document.merge(**dt)
#     document.write(f'{Save_path}/{i}_output.docx')

