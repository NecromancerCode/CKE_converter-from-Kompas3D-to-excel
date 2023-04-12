from openpyxl import load_workbook
import openpyxl

def save_file(path, cols):     
    file = path
    wb = load_workbook(file)
    ws = wb.active
        
    for row_txt in range(7, 7+len(cols)):
        txt=cols[row_txt-7]
        a=0
        for col_txt in [3,4,5,6,8,9]:
            value_txt = txt[a]
            a=a+1
            ws.cell(row=row_txt, column=col_txt).value = value_txt

    wb.save(file)