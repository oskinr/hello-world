import os
from pathlib import Path
import zipfile
import sys
import codecs
sys.stdout = codecs.getwriter("utf-8")(sys.stdout.detach())

# Установка целевой папки



directory = input('Введите путь к папке для извлечения файлов:')
#directory = ("osa")

for file in os.listdir(directory):
    filename = os.fsdecode(file)
    path = directory + "\\" + filename
    #print(path)
    if filename.endswith(".zip"):
        #print(filename)
        with zipfile.ZipFile(path, 'r') as zipf:
            filik = zipf.namelist()
            # print(filik.count)
            # print(filik)
            namefaile = filik[0]
            old_file = f'{directory}\\{namefaile}'
             # print()
            new_file = f'{directory}\\{Path(filename).stem}{".xls"}'
            zipf.extract(namefaile, path=directory)
            if os.path.exists(new_file):
                    os.remove(new_file)
                    os.rename(old_file, new_file)

                    print(f"из {filename} извлечен файл:{
                          os.path.basename(new_file)}")
            else:
                    os.rename(old_file, new_file)
