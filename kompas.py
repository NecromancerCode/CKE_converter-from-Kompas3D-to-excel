# -*- coding: utf-8 -*-
#|PartsInfo

import pythoncom
import subprocess
from win32com.client import Dispatch, gencache

def get_kompas_api7():
    module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    api = module.IKompasAPIObject(
        Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module.IKompasAPIObject.CLSID,
                                                                 pythoncom.IID_IDispatch))
    const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
    return module, api, const

# Функция проверки, запущена-ли программа КОМПАС 3D
def is_running():
    proc_list = \
    subprocess.Popen('tasklist /NH /FI "IMAGENAME eq KOMPAS*"', shell=False, stdout=subprocess.PIPE).communicate()[0]
    return True if proc_list else False

def parse_detail_info(doc_path):
    is_run = is_running()  # True, если программа Компас уже запущена

    module7, api7, const7 = get_kompas_api7()  # Подключаемся к программе
    app7 = api7.Application  # Получаем основной интерфейс программы
    app7.Visible = True  # Показываем окно пользователю (если скрыто)
    app7.HideMessage = const7.ksHideMessageNo  # Отвечаем НЕТ на любые вопросы программы

    #for path in doc_paths:
    doc7 = app7.Documents.Open(PathName=doc_path,
                                   Visible=True,
                                   ReadOnly=True)  # Откроем файл в видимом режиме без права его изменять
    doc3D = module7.IKompasDocument3D(doc7)
    iPart7 = doc3D.TopPart
    sbPartCollection = iPart7.Parts
    #MetalContainer = sbPartCollection._oleobj_.QueryInterface(module7.ISheetMetalContainer.CLSID, pythoncom.IID_IDispatch)
    #MetalContainer = module7.ISheetMetalContainer(MetalContainer)
    #MetalBodies = MetalContainer.SheetMetalBodies
    info = parts_info(iPart7, sbPartCollection)

    if not is_run: app7.Quit()  # Закрываем программу при необходимости
    return info


def detail_info(Part7, Part):

    number, dse, marking, value, measure, entry, material, note = 0, '', '', 0, '', '', '', ''

    if (Part.Detail == 0):
        dse = 'Сб.ед.'
        material = ''
    elif (Part.Standard == 1):
        dse = 'Ст.изд.'
        material = Part.Material
    else:
        dse = 'Дет.'
        material = Part.Material
        #if MetalBodies.SheetMetalBody(0) != None:
            #material = material + ' ' + str(MetalBodies.SheetMetalBody(0).Thickness) + ' мм'
    marking = Part.Marking
    number = number+1
    value = Part7.InstanceCount(Part)
    measure = 'шт.'
    entry = Part7.Marking

    row_text = [dse, marking, value, measure, entry, material]
    return row_text

def parts_info(Part7, PartCollection):
    parts_array = []
    info = []
    for num in range(len(PartCollection)):
        workPart = PartCollection.Part(num)
        if workPart.Name not in parts_array:
            parts_array.append(workPart.Name)
            info.append(detail_info(Part7, workPart))
    return(info)