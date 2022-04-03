from main_sg import main_window
from dataclasses import Field
from pydoc import doc
from word_reader import readAllFields
from mailmerge import MailMerge
import pandas as pd
#path = r"C:\Users\vguzaerov\Documents\GitHub\easydocs\easydocs\sample_docs\Template.docx"
#documentFields = readAllFields('sample_docs\Template.docx')
#df = pd.DataFrame([[1]* len(list(documentFields))],columns=list(documentFields))
#df.to_excel("sample_docs/output.xlsx")  


#df1 = pd.read_excel('sample_docs/output.xlsx')
#print(df1.columns)

WORD_path, Excel_path, Save_path = main_window()



df1 = pd.read_excel(Excel_path, index_col=0)

print(df1)

for i in df1.index:
    document = MailMerge(WORD_path)
    dt=(df1.loc[[i]].to_dict('records')[0])
    document.merge(**dt)
    document.write(f'{Save_path}/{i}_output.docx')

