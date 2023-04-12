import kompas, table

file_path = open("tr.txt", encoding='utf-8')
paths = file_path.readlines()
kompas_path = paths[0]
xl_path = paths[1]

table.save_file(xl_path, kompas.parse_detail_info('C:/Users/odrin/Desktop/макросф/Приложение/Сборка платформа упрощенная.a3d'))