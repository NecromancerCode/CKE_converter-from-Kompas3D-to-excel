def sort_func(r):
    x = r.split(' ')
    if '.' in x[0] and len(x[0].split('.')) > 2:
        return x[0]
    else:
        return 'ЯЯЯЯЯЯЯ'

def sort_dse(name, info):
    buf_info = []
    for i in range(len(info)):
        if name in info[i]:
            buf_info.append(info[i])
    buf_info.sort(key = lambda x: sort_func(x[2]))
    buf_info.sort(key = lambda x: x[0])
    return buf_info

def info_sort(info):
    info_sorted = []
    for i in range(len(info)):
        new_info = []
        buffer = sort_dse('Документация', info[i])
        for k in range(len(buffer)):
            new_info.append(buffer[k])
        buffer = sort_dse('Комплексы', info[i])
        for k in range(len(buffer)):
            new_info.append(buffer[k])
        buffer = sort_dse('Сборочные единицы', info[i])
        for k in range(len(buffer)):
            new_info.append(buffer[k])
        buffer = sort_dse('Детали', info[i])
        for k in range(len(buffer)):
            new_info.append(buffer[k])
        buffer = sort_dse('Стандартные изделия', info[i])
        for k in range(len(buffer)):
            new_info.append(buffer[k])
        buffer = sort_dse('Прочие изделия', info[i])
        for k in range(len(buffer)):
            new_info.append(buffer[k])
        buffer = sort_dse('Материалы', info[i])  
        for k in range(len(buffer)):
            new_info.append(buffer[k])
        buffer = sort_dse('Комплекты', info[i])
        for k in range(len(buffer)):
            new_info.append(buffer[k])
        info_sorted.append(new_info)
    return info_sorted