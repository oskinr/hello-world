import os
from tkinter import filedialog
from xml.dom import minidom
from xml.dom.minidom import parse, parseString
from tkinter import * 
from tkinter import ttk
from tkinter import filedialog as fd
import pandas as pd
from tkinter import messagebox

def openanyfile():
    global file,items,nom,telo,directory,dirpath,it,mydoc,rows
    try:
        global selected_file
        selected_file = filedialog.askopenfilename()
        label1.configure(text=selected_file.split('/')[-1])
        datasource = open(selected_file)
        mydoc = minidom.parse(selected_file)
        
        items = mydoc.getElementsByTagName('Data')
        rows = combobox.get()
        it = items[int(rows)].firstChild.data.split(':')[1].strip()
        column_listbox.insert(0,it)

        
        print('строка 1', rows)
        print('строка 2', mydoc.nodeName)
        print('строка 3', mydoc.firstChild)
        print('строка 4', items)
    except Exception as err:
        messagebox.showerror(
            title="ошибка", message="🔒 Привет от системы: " + str(err))
        
def spisok():
    global it
    rows = combobox.get()
    it = items[int(rows)].firstChild.data.strip()
    column_listbox.insert(0,it)
    

def del_list():
    select = list(column_listbox.curselection())
    select.reverse()
    for i in select:
        column_listbox.delete(i) 


def select():
    rows = combobox.get()
    counter = int(rows)
    for i in range(counter,333,11):
        counter =i
        print(counter) 
        
        row = items[int(counter)].firstChild.data.strip()
        column_listbox.insert(0,row)
      
  
   

        
                


window = Tk()
window.title("Сравнить файлы")
window.geometry("1000x400")


column_listbox = Listbox(selectmode=EXTENDED)
column_listbox.grid(row=1, column=1, columnspan=7, sticky=NSEW, padx=5, pady=5)

label1 = Label( text="",font="system") # создаем текстовую метку
label1.grid(column=12, row=1, padx=6, pady=6)              

Button(text="Найти период", command=spisok).grid(row=0, column=10, columnspan=2, sticky=NSEW, padx=5, pady=5)
Button(text="Выбрать файл", command=openanyfile).grid(row=1, column=10, columnspan=2, sticky=NSEW, padx=5, pady=5)
Button(text="Поиск", command=select).grid(row=2, column=10, columnspan=2, sticky=NSEW, padx=5, pady=5)
Button(text="Удалить", command=del_list).grid(row=3, column=10, columnspan=2, sticky=NSEW, padx=5, pady=5)
# Button(text="Добавить", command=add_item).grid(row=4, column=10, columnspan=2, sticky=NSEW, padx=5, pady=5)
#Button(text="Переименовать", command=openanyfile).grid(row=5, column=10, columnspan=2, sticky=NSEW, padx=5, pady=5)
combo4 = ttk.Combobox(values='')
combo4.grid(row=2, column=3, padx=6, pady=6)

combobox = ttk.Combobox( values=[0,6])
combobox.set(4)
#combobox.current(4)
rows = combobox.get()
combobox.grid(row=2, column=1, padx=6, pady=6)
window.mainloop()  