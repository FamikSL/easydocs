from tkinter import *
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
import os

def submit():
    showinfo("In dev stage", "В разработке")

def set_text(box ,text):
    box.delete(0,END)
    box.insert(0,text)
    return

def select_file(box = None):
    filetypes = (
        ('Docs', '*.docx *.xlsx'),
        ('All files', '*.*')
    )

    filename = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)
    set_text(box, os.path.basename(filename))
    
    
def save_to_dir(box = None):
    dir = fd.askdirectory()
    set_text(box, dir)



# настройки главного окна
main_window = Tk()
main_window.title("easydocs")
main_window.geometry('600x270')
main_window.resizable(False, False)
main_window.iconbitmap(r'brand/icon.ico')

# Секция выбрать шаблон
path_to_doc_box = Entry (main_window, width = 40)
path_to_doc_box.grid(column=1, row=1) 

choose_doc_lb = Label(main_window, text="Выбрать шаблон", font=("Roboto", 14))  
choose_doc_lb.grid(column=0, row=0)  

open_doc_button = Button(main_window, text="Обзор", command = lambda : select_file(path_to_doc_box))
open_doc_button.grid(column=0, row=1, sticky=W, padx=(100), pady=(10))


# Секция выбрать таблицу
path_to_excel_box = Entry (main_window, width = 40)
path_to_excel_box.grid(column=1, row=3) 

choose_excel_lb = Label(main_window, text="Выбрать таблицу", font=("Roboto", 14))  
choose_excel_lb.grid(column=0, row=2)

open_sheet_button = Button(main_window, text="Обзор", command = lambda : select_file(path_to_excel_box))
open_sheet_button.grid(column=0, row=3, sticky=W, padx=(100), pady=(10))    

# Секция создать по шаблону
path_save_dir_box = Entry (main_window, width = 40)
path_save_dir_box.grid(column=1, row=5) 

merge_lb = Label(main_window, text="Секция создать по шаблону", font=("Roboto", 14))  
merge_lb.grid(column=0, row=4)

open_dir_to_save_button = Button(main_window, text="Выбрать папку", command = lambda : save_to_dir(path_save_dir_box))
open_dir_to_save_button.grid(column=0, row=5, sticky=W, padx=(80), pady=(10))    

# Cекция потдверждения

submit_button = Button(main_window, text="Подтвердить", command = lambda : submit())
submit_button.grid(column=1, row=6, sticky=W, padx=(10), pady=(10))   

main_window.mainloop()
