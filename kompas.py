# -*- coding: utf-8 -*-

import pythoncom
import subprocess
from win32com.client import Dispatch, gencache

def get_kompas_api7():  # Получаем АПИ компас 7 версии
    module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    api = module.IKompasAPIObject(
        Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module.IKompasAPIObject.CLSID,
                                                                 pythoncom.IID_IDispatch))
    const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
    return module, api, const   # Возвращаем переменные АПИ

# Функция проверки, запущена-ли программа КОМПАС 3D
def is_running():
    proc_list = \
    subprocess.Popen('tasklist /NH /FI "IMAGENAME eq KOMPAS*"', shell=False, stdout=subprocess.PIPE).communicate()[0]
    return True if proc_list else False

# Функция получения активного документа - ссылки на сборку
def parse_detail_info(doc_path, settings):
    is_run = is_running()  # Проверка запуска компаса. True, если программа Компас уже запущена

    module7, api7, const7 = get_kompas_api7()  # Подключаемся к программе
    app7 = api7.Application  # Получаем основной интерфейс программы
    app7.Visible = True  # Показываем окно пользователю (если скрыто)
    app7.HideMessage = const7.ksHideMessageNo  # Отвечаем НЕТ на любые вопросы программы

    # Для пути - doc_path получаем документ
    doc7 = app7.Documents.Open(PathName=doc_path,
                                   Visible=True,
                                   ReadOnly=True)  # Откроем файл в видимом режиме без права его изменять
    doc3D = module7.IKompasDocument3D(doc7)
    iPart7 = doc3D.TopPart
    iPartCollection7 = iPart7.Parts  # Получим указатель на список всех элементов, входящих в сборку
   
    info, entry = parts_info(iPart7, iPartCollection7, module7, settings)    # Запустим парсинг информации о каждой детали

    if not is_run: app7.Quit()  # Закрываем программу при необходимости
    return info, entry  # Вернем массив информации о каждой детали для записи

# Функция перебора каждого элемента сборки
def parts_info(Part7, PartCollection, module7, settings):
    parts_array = []
    info = []
    number = 0
    for num in range(len(PartCollection)):  
        workPart = PartCollection.Part(num)    # Цикл перебора деталей
        if workPart.Name not in parts_array:    # Если деталь до этого не встречалась, берем в работу
            parts_array.append(workPart.Name)  
            row, number = detail_info(Part7, workPart, module7, settings, number)
            info.append(row)   # Вызовем функцию получения информации о детали
    entry = Part7.Marking
    return(info, entry)  # Вернем массив с массивами информации о каждой детали

# Функция получения информации о деталях
def detail_info(Part7, Part, module7, settings, number):
    # Подключим интерфейс для листовых тел, если они есть
    try:
        MetalContainer = Part._oleobj_.QueryInterface(module7.ISheetMetalContainer.CLSID, pythoncom.IID_IDispatch)
        MetalContainer = module7.ISheetMetalContainer(MetalContainer)
        Bodies = MetalContainer.SheetMetalBodies
        Body = Bodies.SheetMetalBody(0)
    except:
        Body = None
    # Опишем предварительно переменные, используемые в массиве с данными
    dse, marking, value, measure, entry, material = '', '', 0, '', '', ''
    number += 1
    # Определим вид ДСЕ - деталь, сборка, стандартный элемент, а также материал и толщину для листовых тел
    if (Part.Detail == 0):
        dse = 'Cборочные единицы'
        material = ''
    elif (Part.Standard == 1):
        dse = 'Cтандартные изделия'
        material = Part.Material
    else:
        dse = 'Детали'
        if Body != None:
            material = Part.Material + ', ' + str(Body.Thickness) + ' мм'
        else:
            material = Part.Material
    # Определим обозначение, количество, единицу измерения, вхождение в сборку
    if Part.Marking != '':
        marking = Part.Marking + ' ' + Part.Name
    else:
        marking = Part.Name
    value = Part7.InstanceCount(Part)
    measure = 'шт.'
    entry = Part7.Marking
    # Запишем полученные данные в массив
    row_text = []
    if settings[0]: row_text.append(number)
    if settings[1]: row_text.append(dse)
    if settings[2]: row_text.append(marking)
    if settings[3]: 
        row_text.append(value)
        row_text.append(measure)
    if settings[4]: row_text.append(entry)
    if settings[5]: row_text.append(material)
    return row_text, number  # Вернем массив информации о конкретной детали
###