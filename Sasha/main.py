# name = str(input("Р’РІРµРґРёС‚Рµ РЅР°Р·РІР°РЅРёРµ С„Р°Р№Р»Р°: "))
# base = int(input("Р’РІРµРґРёС‚Рµ РѕСЃРЅРѕРІР°РЅРёРµ РґРѕРєСѓРјРµРЅС‚Р°: "))

# print(name)


# with open("/Users/aleksandrmarkov/Downloads/items.txt","r") as f:
#   list = f.read().split("\n")
# f = open('/Users/aleksandrmarkov/Downloads/items.txt', 'w')
# l = [str(i) + str(i - 1) for i in range(20)]
# for index in l:
#     f.write(index + '\n')
# print(l)
# f.close()
# for i in range():
#     command = int(input("Р’РІРµРґРёС‚Рµ РєРѕРјР°РЅРґСѓ: "))
#     if command == 1:
#         print(l[str(i)])
#     else:
#         l.save()
# import pandas as pd
#
#
# file = '/Users/aleksandrmarkov/Documents/РЈС‡РµР±Р°/РЎС‚СЂСѓРєС‚СѓСЂР°_Р±Рґ.xls'
#
#
# xl = pd.ExcelFile(file)
#
#
# print(xl.sheet_names)
#
#
# df1 = xl.parse('Sheet1')

import os
import shutil

KEY_FOR_SEARCH = input('Что ищем?\n')
PATH_FOR_COPY = input('Куда копировать файлы\n')


def search():
    for adress, dirs, files in os.walk(input('Введите путь старта\n')):
        if adress == PATH_FOR_COPY:
            continue
        for file in files:
            if file.endswith('.docx'):
                yield os.path.join(adress, file)


def read_from_pathdocx(path):
    with open(path) as r:
        for i in r:
            if KEY_FOR_SEARCH in i:
                return copy(path)


def copy(path):
    file_name = path.split('\\')[-1]
    count = 1
    while True:
        if os.path.isfile(os.path.join(PATH_FOR_COPY, file_name)):
            if f'({count - 1})' in file_name:
                file_name = file_name.replace(f'({count - 1})', '')
            file_name = f'({count}).'.join(file_name.split('.'))
            count += 1
        else:
            break

    shutil.copyfile(path, os.path.join(PATH_FOR_COPY, file_name))
    print('Р¤Р°Р№Р» СЃРєРѕРїРёСЂРѕРІР°РЅ', file_name)


for i in search():
    try:
        read_from_pathdocx(i)
    except Exception as e:
        with open(os.path.join(PATH_FOR_COPY, 'errors.docx'), 'a') as r:
            r.write(str(e) + '\n' + i + '\n')

import csv
import json

# name_1 = "id"
# name_2 = "Р¤.Р.Рћ."
# name_3 = "РћСЃРЅРѕРІР°РЅРёРµ"
# name_4 = "Р¤Р°Р№Р» в„–1"
# name_5 = "Р¤Р°Р№Р» в„–2"
# name_6 = "Р¤Р°Р№Р» в„–3"
# name_7 = "Р¤Р°Р№Р» в„–4"

# with open("practice.csv", "w") as file:
#     practitioner = csv.practitioner(file, delimiter =",")
#     practitioner.writerow(
#         #[name_1, name_2, name_3, name_4, name_5, name_6, name_7]
#         ("id", "Р¤.Р.Рћ.", "РћСЃРЅРѕРІР°РЅРёРµ", "Р¤Р°Р№Р» в„–1", "Р¤Р°Р№Р» в„–2", "Р¤Р°Р№Р» в„–3", "Р¤Р°Р№Р» в„–4")
#     )

practitioner_data = [
    ("id", "ФИО", "Основание", "Файл №1", "Файл №2", "Файл №3", "Файл №4")

]

# for practitioner in practitioner_data:
#     with open("practice.csv", "a") as file:
#         practitioner = csv.practitioner(file)
#         practitioner.writerow(
#
#         )

with open("practitioner.csv", "w") as file:
    practitioner = csv.practitioner(file)
    practitioner.writerows(
        practitioner_data
    )

# with open("/Users/aleksandrmarkov/Downloads/Reglamenty_expluatatsii_-_formy.docx") as file:
#     src = json.load(file)
#
# asks = src[name_1, name_2, name_3, name_4, name_5, name_6, name_7]
# for a in asks:
