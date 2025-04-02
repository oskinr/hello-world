import os
from tkinter import filedialog
from tkinter import messagebox
import zipfile
from pathlib import PurePath


def zip_ex():
#Путь к файлу попадает в переменную directory
    directory  = filedialog.askdirectory()

#Получить список файлов в директории/каталоге
    files = os.listdir(path=directory)
    messagebox.showinfo("Посмотрим что там", files)
#Получить список файлов в директории/каталоге
    for file in os.listdir(path=directory):
    
    #Собирает директорию и файл в одну строку цикл
        filename = os.fsdecode(file) 
  
   # Получаем полный путь к файлам каталога 
   # нужно просто объединить передаваемый аргумент path с каждым значением из списка внутри цикла
   #path = directory + "\\" + filename 
        path = os.path.join(directory, filename)
        print(path)
   #print(path)
  #Если файл в директории имеет окончание .zip то  
        if filename.endswith(".zip"):
            with zipfile.ZipFile(filename) as zf:
                filik = zf.namelist()
                namefile = filik[0]
                old_file = f'{directory}\\{namefile}'
                new_file = f'{directory}\\{PurePath(filename).stem}{".xls"}'
                zf.extract(namefile, path=directory)
                if os.path.exists(new_file):
                    os.remove(new_file)
                    os.rename(old_file, new_file)
                    print(f"из {filename} извлечен файл:{os.path.basename(new_file)}")
                else:
                    os.rename(old_file, new_file)
                    print(file)
    
zip_ex()
