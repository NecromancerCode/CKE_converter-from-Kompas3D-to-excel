# -*- coding: utf-8 -*-
from tkinter import Label, Button, Entry, Tk, font, filedialog, Menu, Checkbutton
import tkinter as tk 
from tkinter import *
from lang import name1, name2, name3, name4, name5, name6, name7, name8, name9, name10, name11, name12, name13

mass = [0,0,0,0,0,0]

def interface():

    def clicked_folder():      
        file = filedialog.askopenfilename()
        txt1.insert(tk.INSERT,file)

    def clicked_xl():  
        file = filedialog.askopenfilename()
        txt2.insert(tk.INSERT,file)
        
    def clicked_final():  
        fnkomp = txt1.get()
        fnxl = txt2.get()
        f = open('tr.txt','w', encoding='utf-8')

        f.write(fnkomp + "\n" + fnxl + "\n" + str(mass).replace('[', '').replace(']', '')) # отсев ковычек из массива
        window.destroy()

    def clicked_eng():

        global name1, name2, name3, name4, name5, name6, name7, name8, name9, name10, name11, name12, name13 

        window.geometry('670x400')
        i = "KOMPAC to Exel convertor v.21"
        window.title(i)

        name1 = "language"
        name2 = "reference"
        name3 = "The path to the KOMPAS file :"
        name4 = "Specify the KOMPAC assembly file"
        name5 = "The path to the specification .xl :"
        name6 = "Specify the path to the table with the specification"
        name7 = "Conversion to Specification"
        name8 = "Type of DSE"
        name9 = "Name"
        name10 = "Quantity"
        name11 = "Thickness"
        name12 = "Material"
        name13 = "Where it includes" 

        lbl2.configure(text = name3)  
        btn1.configure(text = name4)
        lbl4.configure(text = name5)
        btn2.configure(text = name6)
        btn3.configure(text = name7)
        chk1.configure(text = name8)
        chk2.configure(text = name9)
        chk3.configure(text = name10)
        chk4.configure(text = name11)
        chk5.configure(text = name12)
        chk6.configure(text = name13)
        
    def clicked_ru():

        global name1, name2, name3, name4, name5, name6, name7, name8, name9, name10, name11, name12, name13

        window.geometry('590x400')
        i = "Конвертор KOMPAC в Exel v.21"
        window.title(i)

        name1 = "language"
        name2 = "file"
        name3 = "Путь к файлу KOMPAC :"
        name4 = "Указать сборочный файл KOMPAC"
        name5 = "Путь к спецификаци .xl :"
        name6 = "Указать путь к таблице со спецификацией"
        name7 = "Конвертация в спецификацию"
        name8 = "Вид ДСЕ"
        name9 = "Наименование"
        name10 = "Количество"
        name11 = "Толщина"
        name12 = "Материал (вид)"
        name13 = "Куда входит"

        lbl2.configure(text = name3)  
        btn1.configure(text = name4)
        lbl4.configure(text = name5)
        btn2.configure(text = name6)
        btn3.configure(text = name7)
        chk1.configure(text = name8)
        chk2.configure(text = name9)
        chk3.configure(text = name10)
        chk4.configure(text = name11)
        chk5.configure(text = name12)
        chk6.configure(text = name13)

    def mass_point():
        global mass # массив для записи выбранных параметров

        if chk_state1.get() == 1: mass[0] = 1
        else: mass[0] = 0
        if chk_state2.get() == 1: mass[1] = 1
        else: mass[1] = 0
        if chk_state3.get() == 1: mass[2] = 1
        else: mass[2] = 0
        if chk_state4.get() == 1: mass[3] = 1
        else: mass[3] = 0
        if chk_state5.get() == 1: mass[4] = 1
        else: mass[4] = 0
        if chk_state6.get() == 1: mass[5] = 1
        else: mass[5] = 0

    window = Tk()  # открытие окна
    window.title("Конвертор KOMPAC в Exel v.21")  
    window.geometry('590x400')

    menu = Menu(window) # шапка окна
    new_item = Menu(menu, tearoff=0)  
    new_item.add_command(label='RU', command=clicked_ru)  
    new_item.add_separator()  
    new_item.add_command(label='ENG', command=clicked_eng)  
    menu.add_cascade(label=name1, menu=new_item)
    menu.add_cascade(label=name2, command=clicked_eng)  
    window.config(menu=menu) # конец шапки окна 

    lbl1 = Label(window, text="")  
    lbl1.grid(column=1, row=0)   

    font1 = font.Font(family= "Verdana", size=11, weight="normal", slant="roman") # настройка шрифта (обычный)
    font2 = font.Font(family= "Verdana", size=14, weight="bold", slant="roman") # настройка шрифта (жирный)

    lbl2 = Label(window, text=name3, font = font2)  
    lbl2.grid(column=1, row=1)

    txt1 = Entry(window,width=40) 
    txt1.grid(column=1, row=2)    
    btn1 = Button(window, text=name4, command=clicked_folder, font = font1)  
    btn1.grid(column=1, row=3) 

    lbl3 = Label(window, text="")  
    lbl3.grid(column=1, row=4)

    lbl4 = Label(window, text=name5, font = font2)  
    lbl4.grid(column=1, row=5) 

    txt2 = txt2 = Entry(window,width=40) 
    txt2.grid(column=1, row=6)  
    btn2 = Button(window, text=name6, command=clicked_xl, font = font1)  
    btn2.grid(column=1, row=8) 

    lbl5 = Label(window, text="")  
    lbl5.grid(column=0, row=9)  

    lbl6 = Label(window, text="_"*60)  
    lbl6.grid(column=1, row=10) 

    lbl8 = Label(window, text="_"*60)  
    lbl8.grid(column=1, row=15)
    
    btn3 = Button(window, text=name7, command=clicked_final, font = font2)  
    btn3.grid(column=1, row=16)

    chk_state1 = IntVar()  
    chk1 = Checkbutton(window, text=name8, font = font1, command = mass_point, variable=chk_state1)  
    chk1.grid(column=0, row=13)

    chk_state2 = IntVar()  
    chk2 = Checkbutton(window, text=name9, font = font1, command = mass_point, variable=chk_state2)  
    chk2.grid(column=1, row=13)

    chk_state3 = IntVar()  
    chk3 = Checkbutton(window, text=name10, font = font1, command = mass_point, variable=chk_state3)  
    chk3.grid(column=2, row=13)

    chk_state4 = IntVar()  
    chk4 = Checkbutton(window, text=name11, font = font1, command = mass_point, variable=chk_state4)  
    chk4.grid(column=0, row=14)

    chk_state5 = IntVar()  
    chk5 = Checkbutton(window, text=name12, font = font1, command = mass_point, variable=chk_state5)  
    chk5.grid(column=1, row=14)

    chk_state6 = IntVar()  
    chk6 = Checkbutton(window, text=name13, font = font1, command = mass_point, variable=chk_state6)  
    chk6.grid(column=2, row=14)
    
    window.mainloop() #закрытие окна


    