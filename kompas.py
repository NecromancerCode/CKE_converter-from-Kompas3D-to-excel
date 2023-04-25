# Copyright (c) Converter Kompas-Excel Olesya Droysh and Martynov Ruslan
# All rights reserved.

# This source code is licensed under the license found in the
# LICENSE file in the root directory of this source tree.

# -*- coding: utf-8 -*-
import pythoncom
import subprocess
from win32com.client import Dispatch, gencache, VARIANT

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
def parse_info(doc_path, settings):
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
    iEmbodimentsManager = module7.IEmbodimentsManager(doc3D)  # Вызовем менеджер исполнений
    
    info = parse_embodiments(app7, module7, iEmbodimentsManager, doc3D, settings)  # Запустим парсинг информации о деталях для каждого исполнения
    if not is_run: app7.Quit()  # Закрываем программу при необходимости
    
    return info  # Вернем массив информации о каждой детали для записи

# Функция перебора всех исполнений для получения коллекции деталей
def parse_embodiments(app7, module7, iEmbodimentsManager, doc3D, settings):
    info = []

    for i in range (iEmbodimentsManager.EmbodimentCount):
        iEmbodimentsManager.SetCurrentEmbodiment(i)
        iPart7 = doc3D.TopPart
        iPartCollection7 = iPart7.Parts  # Получим указатель на список всех элементов, входящих в сборку
        info.append(parse_assemble(app7, module7, iPart7, iPartCollection7, settings))     # Вызовем для каждого исполнения парсинг информации о деталях
    
    return info  # Вернем массив с массивами массивов с информацией о каждой детали в исполнении

# Функция перебора всех деталей в коллекции деталей
def parse_assemble(app7, module7, iPart7, iPartCollection7, settings):
    parts_array = []
    info = []

    for num in range(len(iPartCollection7)):  
        workPart = iPartCollection7.Part(num)    # Цикл перебора деталей
        if workPart.Name not in parts_array:    # Если деталь до этого не встречалась, берем в работу
            parts_array.append(workPart.Name)  
            info.append(parse_detail(app7, module7, iPart7, workPart, settings))   # Вызовем функцию получения информации о детали
    
    return info  # Вернем массив с массивами информации о каждой детали

# Функция получения информации о каждой конкретной детали
def parse_detail(app7, module7, iPart7, workPart, settings):
    # Подключение к менеджеру свойств деталей
    try:
        iPropertyMng = module7.IPropertyMng(app7)
        iPropertyKeeper = module7.IPropertyKeeper(workPart)
    except:
        iPropertyMng = None
        iPropertyKeeper = None
    # Подключение к менеджеру листовых деталей
    try:
        MetalContainer = module7.ISheetMetalContainer(MetalContainer)
        Bodies = MetalContainer.SheetMetalBodies
        Body = Bodies.SheetMetalBody(0)
    except:
        Body = None

    # Опишем предварительно переменные, используемые в массиве с данными
    pos, dse, marking, value, measure, entry, material, cover, comment = 0, '', '', 0, '', '', '', '', ''
    # Если получен менеджер свойств, то распарсим информацию о деталях. Наименование параметров берется из стандартных свойств компаса
    if iPropertyMng != None:
        pos = iPropertyKeeper.GetPropertyValue(iPropertyMng.GetProperty(VARIANT(pythoncom.VT_EMPTY, None), "Позиция"), "", True, True)[1]
        if pos == '':
            pos = 0
        pos = int(pos)

        dse = iPropertyKeeper.GetPropertyValue(iPropertyMng.GetProperty(VARIANT(pythoncom.VT_EMPTY, None), "Раздел спецификации"), "", True, True)[1]
        if dse == '':
            dse = iPropertyKeeper.GetPropertyValue(iPropertyMng.GetProperty(VARIANT(pythoncom.VT_EMPTY, None), "Тип объекта"), "", True, True)[1]
        if dse == 'Комплекс': dse = 'Комплексы'
        if dse == 'Деталь': dse = 'Детали'
        if dse == 'Стандартное изделие': dse = 'Стандартные изделия'
        if dse == 'Прочее изделие': dse = 'Прочие изделия'
        if dse == 'Материал': dse = 'Материалы'
        if dse == 'Комплект': dse = 'Комплекты'
        if dse == 'Сборочная единица': dse = 'Сборочные единицы'



        marking = iPropertyKeeper.GetPropertyValue(iPropertyMng.GetProperty(VARIANT(pythoncom.VT_EMPTY, None), "Обозначение"), "", True, True)[1]
        if marking == '':
            marking += iPropertyKeeper.GetPropertyValue(iPropertyMng.GetProperty(VARIANT(pythoncom.VT_EMPTY, None), "Наименование"), "", True, True)[1]
        else:
            marking += ' ' + iPropertyKeeper.GetPropertyValue(iPropertyMng.GetProperty(VARIANT(pythoncom.VT_EMPTY, None), "Наименование"), "", True, True)[1]
        
        value = iPropertyKeeper.GetPropertyValue(iPropertyMng.GetProperty(VARIANT(pythoncom.VT_EMPTY, None), "Количество"), "", True, True)[1]
            # Если количество деталей в свойствах не сходится с количеством в сборке, примем кол-во в сборке за правильное
        if int(value) < iPart7.InstanceCount(workPart):
            value = iPart7.InstanceCount(workPart)
        
        material = iPropertyKeeper.GetPropertyValue(iPropertyMng.GetProperty(VARIANT(pythoncom.VT_EMPTY, None), "Материал"), "", True, True)[1]
        # Если деталь листовая, то добавим толщину к обозначению материала
        if (dse == 'Детали'):
            if Body != None:
                material += ', ' + str(Body.Thickness) + ' мм'

        cover = iPropertyKeeper.GetPropertyValue(iPropertyMng.GetProperty(VARIANT(pythoncom.VT_EMPTY, None), "Зона"), "", True, True)[1]
        comment = iPropertyKeeper.GetPropertyValue(iPropertyMng.GetProperty(VARIANT(pythoncom.VT_EMPTY, None), "Примечание"), "", True, True)[1]

    if len(comment) <= 3 and ('мм' in comment or 'см' in comment or 'м' in comment):
        measure = 'comment'
    else:
        measure = 'шт.'
    entry = iPart7.Marking

    # Запишем полученные данные в массив
    row_text = []
    row_text.append(int(pos))
    if settings[1]: row_text.append(dse)
    if settings[2]: row_text.append(marking)
    if settings[3]: 
        row_text.append(int(value))
        row_text.append(measure)
    if settings[4]: row_text.append(entry)
    if settings[5]: row_text.append(material)
    if settings[6]: row_text.append(cover)
    if settings[7]: row_text.append(comment)
    return row_text # Вернем массив информации о конкретной детали