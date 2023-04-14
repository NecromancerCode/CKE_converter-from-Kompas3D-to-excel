from openpyxl import load_workbook
import openpyxl
import os

# Функция открытия и записи в таблицу значений, где path - путь до таблицы, cols - массив с массивами строк значений
def save_file(path, info, entry):     
    file = path
    wb = load_workbook(file)
    ws = wb.active  # Получим активный лист таблицы
    # Запишем перебором массива значения в каждую из строк таблицы
    ws.cell(row=5, column=1).value = entry
    for row_txt in range(7, 7+len(info)): 
        txt=info[row_txt-7]
        a=0
        for col_txt in [3,4,5,6,8,9]:
            value_txt = txt[a]
            a=a+1
            ws.cell(row=row_txt, column=col_txt).value = value_txt
    # Сохраним файл и откроем в excel 
    wb.save(file)
    os.startfile(file)