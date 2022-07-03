from docx import Document
import pandas as pd
import openpyxl as xl
import re
import docxtpl
import numpy as np

path=input("Полный путь к документу: ")
doc = Document(path)
tables = doc.tables # Получить таблицу, установленную в файле
table = tables [0] # Получить первую таблицу в файле
list = []
test = []
final = []
sp = []
df = [[''for i in range(len(tables[0].columns))]for j in range(len(tables[0].rows))]
for i, row in enumerate(table.rows):
    for j, cell in enumerate(row.cells):
        if cell.text:
            df[i][j] = cell.text
            if '${' in cell.text:
                test.append(df[i][j])
for item in test:
    if item not in final:
        final.append(item)
str_a = ''.join(final)
str = str_a.split('$')
str.pop(0)
for i in range(len(str)):
    str[i] = str[i].split('}')[0]
    str[i] = str[i] + "}"
test = []
test2 = []
del1 = []
table2 = []
for i in range(len(str)):
    print(str[i])
    name = input("Название: ")
    test.append(name)
    print("Если составной тип, то напишите comp")
    type1 = input("Тип: ")
    test2.append(type1)
    if type1 == "comp":
        comp = input("Есть таблица внутри?")
        if comp == "да" or comp == "Да" or comp == "Yes" or comp == "yes":
            table2.append("Yes")
    else:
        table2.append("No")
for i in range(len(str)):
    del1.append("no")
name2 = input("Имя файла xlsx(Писать с расширением!): ")
name3 = input("Имя файла xml(Писать с расширением!): ")
list = pd.DataFrame({'name': test, 'key': str, 'type': test2, 'table_d': table2, 'del': del1})
list.to_excel(name2)
list.to_xml(name3)