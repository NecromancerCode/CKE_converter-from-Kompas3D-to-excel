# -*- coding: utf-8 -*-
import os
import subprocess
from tkinter import Label, Button, Entry, Tk, font, scrolledtext, filedialog
import tkinter as tk
import tkinter.scrolledtext as st
from table import save_file 

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
        f = open('tr.txt','w')
        f.write(fnkomp + "\n" + fnxl) 

    window = Tk()  # открытие окна
    window.title("Конвертор KOMPAC v.21")  
    window.geometry('325x325') 

    lbl1 = Label(window, text="")  
    lbl1.grid(column=1, row=0)   

    font1 = font.Font(family= "Verdana", size=10, weight="normal", slant="roman") # настройка шрифта (обычный)
    font2 = font.Font(family= "Verdana", size=12, weight="bold", slant="roman") # настройка шрифта (жирный)

    lbl2 = Label(window, text="Путь к файлу KOMPAC :", font = font2)  
    lbl2.grid(column=1, row=1)

    txt1 = Entry(window,width=40) 
    txt1.grid(column=1, row=2)    
    btn1 = Button(window, text="Указать сборочный файл KOMPAC", command=clicked_folder, font = font1)  
    btn1.grid(column=1, row=3) 

    lbl3 = Label(window, text="")  
    lbl3.grid(column=1, row=4)

    lbl4 = Label(window, text="Путь к спецификаци .xl :", font = font2)  
    lbl4.grid(column=1, row=5) 

    txt2 = txt2 = Entry(window,width=40) 
    txt2.grid(column=1, row=6)  
    btn2 = Button(window, text="Указать путь к таблице со спецификацией", command=clicked_xl, font = font1)  
    btn2.grid(column=1, row=8) 

    lbl5 = Label(window, text="")  
    lbl5.grid(column=0, row=9)  

    lbl6 = Label(window, text="_"*60)  
    lbl6.grid(column=1, row=10) 

    lbl7 = Label(window, text="")  
    lbl7.grid(column=1, row=11) 
    
    btn3 = Button(window, text="Конвертация в спецификацию", command=clicked_final, font = font2)  
    btn3.grid(column=1, row=12)

    window.mainloop() #закрытие окна

interface()
    