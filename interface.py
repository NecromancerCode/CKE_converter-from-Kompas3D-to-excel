# -*- coding: utf-8 -*-
from tkinter import Label, Button, Entry, Tk, font, filedialog, Menu, Checkbutton, IntVar, messagebox
import tkinter as tk
from lang import eng_txt, ru_txt, FAQ_txt_ru, FAQ_txt_eng

mass = [1, 1, 1, 1, 1, 1, 1, 1]

def interface():
   
    def clicked_folder(): # функция выбора файла
        file = filedialog.askopenfilename()
        txt1.insert(tk.INSERT,file)

    def clicked_xl(): # функция выбора файла 
        file = filedialog.askopenfilename()
        txt2.insert(tk.INSERT,file)
        
    def clicked_final(): 
        fnkomp = txt1.get()
        fnxl = txt2.get()
        comm = txt3.get()

        f = open('settings.txt','w', encoding='utf-8')

        f.write(fnkomp + "\n" + fnxl + "\n" + str(mass).replace('[', '').replace(']', '') + "\n" + comm) # отсев ковычек из массива
        window.destroy()

    def FAQ():
        
        if lbl2.cget("text")  == ru_txt[2]:
            messagebox.showinfo('FAQ', FAQ_txt_ru[0])
        else:
            messagebox.showinfo('FAQ', FAQ_txt_eng[0])
    
    def ch_all():

        if chk_state9.get() == 1: 

            chk_state1.set(1)
            chk_state2.set(1)
            chk_state3.set(1)
            chk_state4.set(1)
            chk_state5.set(1)
            chk_state6.set(1)
            chk_state7.set(1)
            chk_state8.set(1)

        else:

            chk_state1.set(0)
            chk_state2.set(0)
            chk_state3.set(0)
            chk_state4.set(0)
            chk_state5.set(0)
            chk_state6.set(0)
            chk_state7.set(0)
            chk_state8.set(0)

    def clicked_eng(): # функция смены языка на английский

        global eng_txt

        window.geometry('400x600')
        i = "KOMPAS to Exсel convertor v.21"
        window.title(i)

        lbl2.configure(text = eng_txt[2])  
        btn1.configure(text = eng_txt[3])
        lbl4.configure(text = eng_txt[4])
        btn2.configure(text = eng_txt[5])
        btn3.configure(text = eng_txt[6])
        chk1.configure(text = eng_txt[7])
        chk2.configure(text = eng_txt[8])
        chk3.configure(text = eng_txt[9])
        chk4.configure(text = eng_txt[10])
        chk5.configure(text = eng_txt[11])
        chk6.configure(text = eng_txt[12])
        chk7.configure(text = eng_txt[13])
        chk8.configure(text = eng_txt[14])
        chk9.configure(text = eng_txt[15])
        lbl10.configure(text = eng_txt[16])
        
    def clicked_ru(): # функция смены языка на русский

        global ru_txt

        window.geometry('380x600')
        i = "Конвертор KOMПAC в Exсel v.21"
        window.title(i)

        lbl2.configure(text = ru_txt[2])  
        btn1.configure(text = ru_txt[3])
        lbl4.configure(text = ru_txt[4])
        btn2.configure(text = ru_txt[5])
        btn3.configure(text = ru_txt[6])
        chk1.configure(text = ru_txt[7])
        chk2.configure(text = ru_txt[8])
        chk3.configure(text = ru_txt[9])
        chk4.configure(text = ru_txt[10])
        chk5.configure(text = ru_txt[11])
        chk6.configure(text = ru_txt[12])
        chk7.configure(text = ru_txt[13])
        chk8.configure(text = ru_txt[14])
        chk9.configure(text = ru_txt[15])
        lbl10.configure(text = ru_txt[16])

    def mass_point():

        global mass # массив для записи выбранных параметров

        if chk_state1.get() == 0: mass[0] = 0
        else: mass[0] = 1
        if chk_state2.get() == 0: mass[1] = 0
        else: mass[1] = 1
        if chk_state3.get() == 0: mass[2] = 0
        else: mass[2] = 1
        if chk_state4.get() == 0: mass[3] = 0
        else: mass[3] = 1
        if chk_state5.get() == 0: mass[4] = 0
        else: mass[4] = 1
        if chk_state6.get() == 0: mass[5] = 0
        else: mass[5] = 1
        if chk_state7.get() == 0: mass[6] = 0
        else: mass[6] = 1
        if chk_state8.get() == 0: mass[7] = 0
        else: mass[7] = 1

    window = Tk()  # открытие окна
    window.title("Конвертор KOMPAC в Exсel v.21")  
    window.geometry('380x600')

    menu = Menu(window) # шапка окна 
    new_item = Menu(menu, tearoff=0)  
    new_item.add_command(label='RU', command=clicked_ru)  
    new_item.add_separator()  
    new_item.add_command(label='ENG', command=clicked_eng)
    menu.add_cascade(label=ru_txt[0], menu=new_item)
    menu.add_cascade(label=ru_txt[1], command=FAQ)  
    window.config(menu=menu) # конец шапки окна  

    font1 = font.Font(family= "Verdana", size=11, weight="normal", slant="roman") # настройка шрифта (обычный)
    font2 = font.Font(family= "Verdana", size=14, weight="bold", slant="roman") # настройка шрифта (жирный)
    font3 = font.Font(family= "Verdana", size=14, weight="normal", slant="roman") # настройка для текста окна
    font4 = font.Font(family= "Verdana", size=11, weight="bold", slant="roman") # настройка для текста под окном
    
    lbl2 = Label(window, text=ru_txt[2], font = font2)  
    lbl2.grid(column=1, row=1)

    txt1 = Entry(window,width=40) 
    txt1.grid(column=1, row=2)    
    btn1 = Button(window, text=ru_txt[3], command=clicked_folder, font = font4, bg="Silver")  
    btn1.grid(column=1, row=3) 

    lbl3 = Label(window, text="")  
    lbl3.grid(column=1, row=4)

    lbl4 = Label(window, text=ru_txt[4], font = font2)  
    lbl4.grid(column=1, row=5) 

    txt2 = txt2 = Entry(window,width=40) 
    txt2.grid(column=1, row=6)  
    btn2 = Button(window, text=ru_txt[5], command=clicked_xl, font = font4, bg="Silver")  
    btn2.grid(column=1, row=8) 

    lbl6 = Label(window, text="_"*60)  
    lbl6.grid(column=1, row=10) 

    lbl8 = Label(window, text="_"*60)  
    lbl8.grid(column=1, row=26)

    lbl9 = Label(window, text=" ")  
    lbl9.grid(column=0, row=30)
    
    btn3 = Button(window, text=ru_txt[6], command=clicked_final, font = font2, bg="LightSlateGray")  
    btn3.grid(column=1, row=27,  ipadx=0, ipady=0, padx=0, pady=0)

    chk_state1 = IntVar()  
    chk1 = Checkbutton(window, text=ru_txt[7], font = font1, command = mass_point, variable=chk_state1)  
    chk1.grid(column=1, row=21, ipadx=0, ipady=0, padx=0, pady=0)

    chk_state2 = IntVar()  
    chk2 = Checkbutton(window, text=ru_txt[8], font = font1, command = mass_point, variable=chk_state2)  
    chk2.grid(column=1, row=14, ipadx=0, ipady=0, padx=0, pady=0)

    chk_state3 = IntVar()  
    chk3 = Checkbutton(window, text=ru_txt[9], font = font1, command = mass_point, variable=chk_state3)  
    chk3.grid(column=1, row=15, ipadx=0, ipady=0, padx=0, pady=0)

    chk_state4 = IntVar()  
    chk4 = Checkbutton(window, text=ru_txt[10], font = font1, command = mass_point, variable=chk_state4)  
    chk4.grid(column=1, row=16, ipadx=0, ipady=0, padx=0, pady=0)

    chk_state5 = IntVar()  
    chk5 = Checkbutton(window, text=ru_txt[11], font = font1, command = mass_point, variable=chk_state5)  
    chk5.grid(column=1, row=17, ipadx=0, ipady=0, padx=0, pady=0)

    chk_state6 = IntVar()  
    chk6 = Checkbutton(window, text=ru_txt[12], font = font1, command = mass_point, variable=chk_state6)  
    chk6.grid(column=1, row=18, ipadx=0, ipady=0, padx=0, pady=0)

    chk_state7 = IntVar()  
    chk7 = Checkbutton(window, text=ru_txt[13], font = font1, command = mass_point, variable=chk_state7)  
    chk7.grid(column=1, row=19, ipadx=0, ipady=0, padx=0, pady=0)

    chk_state8 = IntVar()  
    chk8 = Checkbutton(window, text=ru_txt[14], font = font1, command = mass_point, variable=chk_state8)  
    chk8.grid(column=1, row=20, ipadx=0, ipady=0, padx=0, pady=0)

    chk_state9 = IntVar()  
    chk9 = Checkbutton(window, text=ru_txt[15], font = font1, command = ch_all, variable=chk_state9)  
    chk9.grid(column=1, row=13, ipadx=0, ipady=0, padx=0, pady=0)

    txt3 = Entry(window,width=22, font=font3) 
    txt3.grid(column=1, row=24)

    lbl10 = Label(window, text=ru_txt[16], font = font4)  
    lbl10.grid(column=1, row=25)

    chk_state1.set(1)
    chk_state2.set(1)
    chk_state3.set(1)
    chk_state4.set(1)
    chk_state5.set(1)
    chk_state6.set(1)
    chk_state7.set(1)
    chk_state8.set(1)
    
    window.mainloop() #закрытие окна