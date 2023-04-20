# Copyright (c) Converter Kompas-Excel Olesya Droysh and Martynov Ruslan
# All rights reserved.

# This source code is licensed under the license found in the
# LICENSE file in the root directory of this source tree.

import kompas, table
from interface import interface
from os import remove

interface()

file_path = open("settings.txt", encoding='utf-8')
paths = file_path.readlines()

try:
    kompas_path = paths[0].replace('\n', '')
except: 
    kompas_path = ''

try:
    xl_path = paths[1].replace('\n', '')
except: 
    xl_path = ''    

variables = paths[2].replace('\n', '')

try:
    entry = paths[3].replace('\n', '')
except: 
    entry = ''

settings = [int(numeric_string) for numeric_string in variables.split(", ")]

file_path.close()
remove("settings.txt",)

if kompas_path != '':
    info = kompas.parse_info(kompas_path, settings)
    for i in range(len(info)):
        info[i].sort(key = lambda x: x[0])
        
    table.save_file(xl_path, settings, info, entry)