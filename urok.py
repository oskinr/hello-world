from tkinter import ttk
from openpyxl import *
from tkinter import * 
from tkinter import messagebox
from tkinter import filedialog
import pandas as pd
import tkinter as tk
from tkinter.ttk import Combobox
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog
import xml.etree.ElementTree as ET
import numpy as np





# Открываем файл 1
def openanyfile():

    global selected_file
    selected_file  = filedialog.askopenfilename() 
    label1.configure(text= selected_file)
# добавление нового элемента
def add():
    new_column = combo2.get()
    column_listbox.insert(0, new_column)
    
    st1 =(new_column)
    df1 = pd.read_excel(selected_file, usecols=[st1],skiprows=int(rows))
   
    print(df1) 
          

      
def red():  
    global d,a,c,b,st1,st2,df1,df2
    rows=combo.get()
    
    df1 = pd.read_excel(selected_file, header=0, skiprows=int(rows))
    #df1 = pd.read_csv(selected_file, encoding = 'Windows-1251')
    # neworder = ['номер','площадь' ,'товар', 'Количество']
    # df = df1.reindex(columns=neworder)
    #df = df1.reindex(columns=['DATE',0,1,2,3])#ПУСТОЙ ДАТАФРЕЙМ
    # df= df1[['A','D','C','B']]
    # cols = list(df1.columns)
    # a,b,c,d= cols.index('площадь'),cols.index('товар'),cols.index('Количество'),cols.index('номер')
    # cols2 = cols[d],cols[a],cols[c],cols[b] 
    
    # a = (df1.columns.tolist()[0])
    # b = (df1.columns.tolist()[1])
    # c = (df1.columns.tolist()[2])
    # d = (df1.columns.tolist()[3])
    # print(d,a,c,b)

    
    label4.configure(text=df1.columns.tolist())
    col_name = list(df1.columns)
    combo2['values'] = col_name
    label3.configure(text=df1)        

def highlight_col(x):
    #copy df to new - original data are not changed
    df = x.copy()
    #set by condition
    mask = df['Количество'] == 23
    df.loc[mask, :] = 'background-color: yellow'
    df.loc[~mask,:] = 'background-color: white'
    return df 

def remove_text():
    #languages_listbox.config(text="")
    label3.config(text="")
    label1.config(text="")
    label4.config(text="")       

def print_list():
    print(column_listbox.get(0, END))
    #df = (column_listbox.get(0, END))
    #modified_list = str(df).replace('(', '').replace(')', '')
    #modified_list = str(df).replace('(', '').replace(')', '')
    #modified_list = ', '.join([str(element) for element in df])
    #df = ['товар', 'площадь']
   
    df= ' '.join(column_listbox.get(0, END))
    #modified_list = (df.split())
    print(df)
    
    modified_list = df.split()
    print(modified_list)
    
  
   
  
    
    
    df2 = df1[modified_list]
   
   
    #df2['compare'] = df2['col1'] < df2['col2']
    
    #df2['compare'] = df2['Сравниваем_Файл_1'] == df2['Сравниваем_Файл_2']
    #df2.loc[((df2['compare']= 'background-color: #D8E4BC'))]
    
 
    

    df = df2
    #df2.style.set_properties(**{'text-align': 'center','border': '1.3px solid black', 'color': 'black'})
    #df2.describe().T.drop("count", axis=1)\
                 #.style.highlight_max(color="darkred")
    df2.style.apply(highlight_col, axis=None).to_excel('df2.xlsx', index=False)
    print(df2)
    #df2.style.set_properties(**{'text-align': 'center','border': '1.3px solid black', 'color': 'black'}).to_excel('df2.xlsx', index=False)
    messagebox.showinfo("Title", "Создан фал df2.xlsx")

def del_list():
    select = list(column_listbox.curselection())
    select.reverse()
    for i in select:
        column_listbox.delete(i)


    
  




window = Tk()
window.title("Сравнить файлы")
window.geometry("1000x400")

#frame.pack(expand=False)

#кнопки
file = Button(window, text="Файл 1",command=openanyfile)
file.grid(column=0, row=0, padx=6, pady=6, sticky=EW)
# file1 = Button(window, text="добавить",command=rowse)
# file1.grid(row=3, column=0)
file1 = Button(window, text="читать",command=red)
file1.grid(column=1, row=0, padx=6, pady=6, sticky=EW)
#текстовой вывод пути к  фалам
label1 = Label(window, text="",font="system") # создаем текстовую метку
label1.grid(column=5, row=2, padx=6, pady=6) 
#вывод конечного файла
label3 = Label(window, text="", justify=tk.LEFT) # создаем текстовую метку
label3.grid(column=5, row=2, padx=6, pady=6)

#вывод заголовков
label4 = Label(window, text="", justify=tk.LEFT)  # создаем текстовую метку
label4.grid(column=5, row=1, padx=6, pady=6)


#комбобоксы для ввода ключа слияния и сравнения
combo = Combobox(window, values=[0,6])
combo.current(0)
rows = combo.get()
combo.grid(row=2, column=2, pady=6)

# текстовое поле и кнопка для добавления в список
combo2 = Combobox(window, values='')
combo2.grid(column=2, row=0, padx=6, pady=6, sticky=EW)
ttk.Button(text="Добавить", command=add).grid(column=3, row=0, padx=6, pady=6)

# languages = ['']
# languages_var = StringVar(value=languages)
 

show3 = ttk.Button( text="Удалить текст", command=remove_text)
show3.grid(column=4, row=0, padx=6, pady=6)
# создаем список
column_listbox = Listbox(selectmode=EXTENDED)
column_listbox.grid(row=1, column=1, columnspan=2, sticky=NSEW, padx=5, pady=5)



Button(text="Print", command=print_list).grid(row=1, column=5, columnspan=2, sticky=NSEW, padx=5, pady=5)
Button(text="Delete", command=del_list).grid(row=1, column=7, columnspan=2, sticky=NSEW, padx=5, pady=5)




window.mainloop()
