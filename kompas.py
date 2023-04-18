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
   
    info = parts_info(iPart7, iPartCollection7, module7, settings, app7)    # Запустим парсинг информации о каждой детали

    if not is_run: app7.Quit()  # Закрываем программу при необходимости
    return info  # Вернем массив информации о каждой детали для записи

# Функция перебора каждого элемента сборки
def parts_info(Part7, PartCollection, module7, settings, app7):
    parts_array = []
    info = []
    for num in range(len(PartCollection)):  
        workPart = PartCollection.Part(num)    # Цикл перебора деталей
        if workPart.Name not in parts_array:    # Если деталь до этого не встречалась, берем в работу
            parts_array.append(workPart.Name)  
            row = detail_info(Part7, workPart, module7, settings,  app7)
            info.append(row)   # Вызовем функцию получения информации о детали
    return(info)  # Вернем массив с массивами информации о каждой детали

# Функция получения информации о деталях
def detail_info(Part7, Part, module7, settings, app7):
    
    try:
        iPropertyMng = module7.IPropertyMng(app7)
        iPropertyKeeper = module7.IPropertyKeeper(Part)
    except:
        iPropertyMng = None
        iPropertyKeeper = None
    try:
        MetalContainer = module7.ISheetMetalContainer(MetalContainer)
        Bodies = MetalContainer.SheetMetalBodies
        Body = Bodies.SheetMetalBody(0)
    except:
        Body = None

    # Опишем предварительно переменные, используемые в массиве с данными
    pos, dse, marking, value, measure, entry, material, cover, comment = 0, '', '', 0, '', '', '', '', ''

    if iPropertyMng != None:
        pos = iPropertyKeeper.GetPropertyValue(iPropertyMng.GetProperty(VARIANT(pythoncom.VT_EMPTY, None), "Позиция"), "", True, True)[1]
        if pos == '':
            pos = 0
        pos = int(pos)
        dse = iPropertyKeeper.GetPropertyValue(iPropertyMng.GetProperty(VARIANT(pythoncom.VT_EMPTY, None), "Раздел спецификации"), "", True, True)[1]
        marking = iPropertyKeeper.GetPropertyValue(iPropertyMng.GetProperty(VARIANT(pythoncom.VT_EMPTY, None), "Обозначение"), "", True, True)[1]
        if marking == '':
            marking += iPropertyKeeper.GetPropertyValue(iPropertyMng.GetProperty(VARIANT(pythoncom.VT_EMPTY, None), "Наименование"), "", True, True)[1]
        else:
            marking += ' ' + iPropertyKeeper.GetPropertyValue(iPropertyMng.GetProperty(VARIANT(pythoncom.VT_EMPTY, None), "Наименование"), "", True, True)[1]
        value = iPropertyKeeper.GetPropertyValue(iPropertyMng.GetProperty(VARIANT(pythoncom.VT_EMPTY, None), "Количество"), "", True, True)[1]
        material = iPropertyKeeper.GetPropertyValue(iPropertyMng.GetProperty(VARIANT(pythoncom.VT_EMPTY, None), "Материал"), "", True, True)[1]
        cover = iPropertyKeeper.GetPropertyValue(iPropertyMng.GetProperty(VARIANT(pythoncom.VT_EMPTY, None), "Зона"), "", True, True)[1]
        comment = iPropertyKeeper.GetPropertyValue(iPropertyMng.GetProperty(VARIANT(pythoncom.VT_EMPTY, None), "Примечание"), "", True, True)[1]

    # Определим вид ДСЕ - деталь, сборка, стандартный элемент, а также материал и толщину для листовых тел
    if (dse == 'Детали'):
        if Body != None:
            material += ', ' + str(Body.Thickness) + ' мм'
    if value != Part7.InstanceCount(Part):
        value = Part7.InstanceCount(Part)
    measure = 'шт.'
    entry = Part7.Marking
    # Запишем полученные данные в массив
    row_text = []
    if settings[0]: row_text.append(pos)
    if settings[1]: row_text.append(dse)
    if settings[2]: row_text.append(marking)
    if settings[3]: 
        row_text.append(value)
        row_text.append(measure)
    if settings[4]: row_text.append(entry)
    if settings[5]: row_text.append(material)
    if settings[6]: row_text.append(cover)
    if settings[7]: row_text.append(comment)
    return row_text # Вернем массив информации о конкретной детали