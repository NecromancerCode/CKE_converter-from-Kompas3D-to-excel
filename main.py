import kompas, table
from interface import interface

interface()

file_path = open("tr.txt", encoding='utf-8')
paths = file_path.readlines()
kompas_path = paths[0].replace('\n', '')
xl_path = paths[1].replace('\n', '')
table.save_file(xl_path, kompas.parse_detail_info(kompas_path))