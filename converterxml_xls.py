import win32com.client
from tkinter import filedialog
file = filedialog.askopenfilename().replace('/', '\\')

print(file)
wbf = file + "x"
print(wbf)

def main():
    print("конвертация xml")

    office = win32com.client.Dispatch("Excel.Application")
    wb = office.Workbooks.Open(file)
    sheet = wb.ActiveSheet
    office.DisplayAlerts = False  # не спрашивает перезаписать файл
    # val = sheet.Cells(1,1)
    num = [r for r in sheet.Range("A8:K8")]
    print(*num)
    # print(val)
    wb.SaveAs(Filename=wbf, FileFormat=51)
    wb.Close(True)
    office.Quit()
if __name__ == "__main__":
    main()


# def main2():
#     print("Hello from excel-py!")
#     dataframe = openpyxl.load_workbook("new.xlsx")
    
#     dataframe1 = dataframe.active

#     # Iterate the loop to read the cell values
#     for row in range(0, dataframe1.max_row):
#         for col in dataframe1.iter_cols(1, dataframe1.max_column):
#             pass
#     print(dataframe1.max_row, dataframe1.max_column)
#         # print(col[row].value)
#     # Создание новой книги Excel
#     # wb = Workbook()

#     # sheet = wb.active
#     # sheet.title = 'Employees'

#     # Добавление заголовков
#     # sheet.append(['ID', 'Name', 'Age', 'Position'])

#     # Заполнение данными
#     # for emp in employees:
#         # sheet.append([emp['id'], emp['name'], emp['age'], emp['position']])

#     # Сохранение файла
#     # wb.save('employees.xlsx')


# def main3():
#     f = open('19.07.xls', 'wb')
#     f.write(xml2xlsx("new2.xlsx"))
#     f.close()
    
