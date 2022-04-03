for i in df1.index:
    document = MailMerge(WORD_path)
    dt=(df1.loc[[i]].to_dict('records')[0])
    document.merge(**dt)
    document.write(f'{Save_path}/{i}_output.docx')
