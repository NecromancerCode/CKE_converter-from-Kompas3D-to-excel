from openpyxl import load_workbook
import openpyxl
import os

# Функция открытия и записи в таблицу значений, где path - путь до таблицы, cols - массив с массивами строк значений
def save_file(path, settings, info, entry):     
    file = path
    wb = load_workbook(file)
    ws = wb.active  # Получим активный лист таблицы
    row = 7
    columns = []
    2, 3, 4, 5, 6, 8, 9
    if settings[0]: columns.append(2)
    if settings[1]: columns.append(3)
    if settings[2]: columns.append(4)
    if settings[3]: 
        columns.append(5)
        columns.append(6)
    if settings[4]: columns.append(8)
    if settings[5]: columns.append(9)
    # Запишем перебором массива значения в каждую из строк таблицы
    ws.cell(row = 5, column = 1).value = entry
    for row_txt in range(row, row + len(info)): 
        txt=info[row_txt-row]
        a=0
        for col_txt in columns:
            value_txt = txt[a]
            a=a+1
            ws.cell(row=row_txt, column=col_txt).value = value_txt
    # Сохраним файл и откроем в excel 
    wb.save(file)
    os.startfile(file)