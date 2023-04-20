import kompas, table
from interface import interface

interface()

file_path = open("settings.txt", encoding='utf-8')
paths = file_path.readlines()
kompas_path = paths[0].replace('\n', '')
xl_path = paths[1].replace('\n', '')
variables = paths[2].replace('\n', '')
settings = [int(numeric_string) for numeric_string in variables.split(", ")]

info = kompas.parse_info(kompas_path, settings)
info.sort(key = lambda x: x[0])
table.save_file(xl_path, settings, info)