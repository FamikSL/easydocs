from mailmerge import MailMerge
def merge_docs(table_data, Word_path, Save_path):
    for i in table_data.index:
        document = MailMerge( Word_path)
        dt=(table_data.loc[[i]].to_dict('records')[0])
        document.merge(**dt)
        document.write(f'{Save_path}/{i}_output.docx')
