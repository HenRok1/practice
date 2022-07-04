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
    str[i] = "$" + str[i] + "}"
test = []
test2 = []
del1 = []
table2 = []
for i in range(len(str)):
    new = []
    new2 = []
    new3 = []
    new4 = []
    new5 = []
    print(str[i])
    name = input("Название: ")
    test.append(name)
    print("Если составной тип, то напишите comp")
    type1 = input("Тип: ")
    test2.append(type1)
    if type1 == "comp" or type1 == "Comp":
        comp = input("Есть таблица внутри?")
        if comp == "да" or comp == "Да" or comp == "Yes" or comp == "yes":
            table2.append("Yes")
    elif type1 == "table" or type1 == "Table":
        table2.append("Yes")
    else:
        table2.append("No")
for i in range(len(str)):
    del1.append("no")
name2 = input("Имя файла xlsx(Писать с расширением!): ")
name3 = input("Имя файла xml(Писать с расширением!): ")
list = pd.DataFrame({'name': test, 'key': str, 'type': test2, 'table_d': table2, 'del': del1})
writer = pd.ExcelWriter(name2)
list.to_excel(writer, sheet_name='main')
frames = []
frames.append(list)
for i in range(len(test2)):
    if test2[i] == "comp" or test2[i] == "Comp":
        print(str[i])
        lend = int(input("Из скольких частей состоит? "))
        for j in range(lend):
            new.append(str[i])
            type12 = input("Тип части: ")
            new2.append(type12)
            new3.append(test[i])
            if type12 == "Table" or type12 == "table":
                new4.append("Yes")
            else:
                new4.append("No")
            new5.append("No")
        new_table = pd.DataFrame({'name': new3, 'key': new, 'type': new2, 'table_d': new4, 'del': new5})
        new_table.to_excel(writer, sheet_name=str[i])
        frames.append(new_table)
writer.save()
list = pd.concat(frames)
list.to_xml(name3)
