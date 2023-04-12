import kompas, interface, table

file_path = open("tr.txt")
paths = file_path.readlines()
kompas_path = paths[0]
xl_path = paths[1]

table.save_file(xl_path, kompas.parse_detail_info(kompas_path))
