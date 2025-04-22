import os
from tkinter import filedialog
from xml.dom import minidom
from tkinter import * 
from tkinter import ttk
from tkinter import filedialog as fd

class Formper():
    def __init__(self):
        win = Tk()
        win.title("Сравнить файлы")
        win.geometry("1000x400")
        # fr = Frame(win, width=1000, height=400)

        # frame = Frame(win)
        self.column_listbox = Listbox(win, selectmode=EXTENDED)
        self.column_listbox.grid(row=1, column=1, columnspan=2, sticky=NSEW, padx=5, pady=5)
        self.label1 = Label(win, text="",font="system") # создаем текстовую метку
        self.label1.grid(column=3, row=1, padx=6, pady=6)              
        # self.label1.pack()
        Button(win, text="Найти период", command=self.spisok).grid(row=0, column=10, columnspan=2, sticky=NSEW, padx=5, pady=5)
        Button(win, text="Папка", command=self.selectDir).grid(row=2, column=10, columnspan=2, sticky=NSEW, padx=5, pady=5)
        Button(win, text="Удалить", command=self.del_list).grid(row=3, column=10, columnspan=2, sticky=NSEW, padx=5, pady=5)
        Button(win, text="Добавить", command=self.add_item).grid(row=4, column=10, columnspan=2, sticky=NSEW, padx=5, pady=5)
        Button(win, text="Переименовать", command=self.per).grid(row=5, column=10, columnspan=2, sticky=NSEW, padx=5, pady=5)
        combo4 = ttk.Combobox(win, values='')
        combo4.grid(row=2, column=3, padx=6, pady=6)

        combobox = ttk.Combobox(win, values=[0,6])
        combobox.set(4)
        rows = combobox.get()
        combobox.grid(row=2, column=1, padx=6, pady=6)
        win.mainloop()
        
    def per(self):   
        for dirpath, dirnames, filenames in os.walk(directory):
            ig = 0
            for filenames in filenames:
                    
                if filenames.split('.')[-1].lower() == 'xls':
                    file = os.path.join(dirpath,filenames)
                    mydoc = minidom.parse(file)
                    items = mydoc.getElementsByTagName('Data')
                    rows = self.combobox.get()
                    it = items[int(rows)].firstChild.data.split(':')[1].strip()
                    self.column_listbox.insert(0,it)
                    ig += 1
                    newfile = it + '_' + filenames.split('_')[0].strip() +'_' + str(ig) + '.xls'
                    os.rename(os.path.join(dirpath, filenames), os.path.join(dirpath, newfile))

    
    def del_list(self):
        select = list(self.column_listbox.curselection())
        select.reverse()
        for i in select:
            self.column_listbox.delete(i)          


    def spisok(self):
        global it
        rows = self.combobox.get()
        it = items[int(rows)].firstChild.data.split(':')[1].strip()
        self.column_listbox.insert(0,it)
        print(rows)
        
    def add_item(self):
        self.column_listbox.insert(END, self.combo4.get())
        self.combo4.delete(0, END)
        self.combo4.insert(0,it)

    def selectDir(self):
        global file,items,nom,telo,directory,dirpath,it
        directory = filedialog.askdirectory()+'/'
        self.label1.configure(text = directory)
        for dirpath, dirnames, filenames in os.walk(directory):
            for filenames in filenames:
                if filenames.split('.')[-1].lower() == 'xls':
        
                    file = os.path.join(dirpath,filenames)
                    mydoc = minidom.parse(file)
                    items = mydoc.getElementsByTagName('Data')
                    rows = self.combobox.get()
                    it = items[int(rows)].firstChild.data.split(':')[1].strip()
                    self.column_listbox.insert(0,it)

if __name__ == '__main__':
    Formper()
  