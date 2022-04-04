from Word_Reader import readAllFields
import pandas as pd
import os


Word_path = r"C:/Users/Vadim/Documents/GitHub/easydocs/sample_docs/Template(test1).docx"
Save_path = r'C:/Users/Vadim/Documents/GitHub/easydocs/sample_docs/out'
new_name = os.path.basename(Word_path).split('.')[0]
documentFields = readAllFields(Word_path)[0]
df = pd.DataFrame([['']* len(list(documentFields))],columns=list(documentFields))
print(df.head())
#df.to_excel(f'{Save_path}/{new_name}.xlsx') 