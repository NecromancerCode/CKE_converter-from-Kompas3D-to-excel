def sort_func(r):
    x = r.split(' ')
    if '.' in x[0] and len(x[0].split('.')) > 2:
        return x[0]
    else:
        return 'ЯЯЯЯЯЯЯ'

def info_sort(info):
    for i in range(len(info)):
        info[i].sort(key = lambda x: x[0])
        if 0 not in info[i][0]: 
            info[i].sort(key = lambda x: x[0])
        else:
            for k in range(len(info[i])): 
                info[i].sort(key = lambda x: x[2])
                info[i].sort(key = lambda x: sort_func(x[2]))
                info[i].sort(key = lambda x: x[1] == 'Документация')
                info[i].sort(key = lambda x: x[1] == 'Комплексы')
                info[i].sort(key = lambda x: x[1] == 'Сборочные еденицы')
                info[i].sort(key = lambda x: x[1] == 'Детали')
                info[i].sort(key = lambda x: x[1] == 'Стандартные изделия')
                info[i].sort(key = lambda x: x[1] == 'Прочие изделия')
                info[i].sort(key = lambda x: x[1] == 'Материалы')
                info[i].sort(key = lambda x: x[1] == 'Комплекты')
    return info