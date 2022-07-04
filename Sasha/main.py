import os
import shutil
import csv

from docx import Document

# KEY_FOR_SEARCH = input('Что ищем?\n')
# PATH_FOR_COPY = input('Куда копировать файлы?\n')

path_1 = input("Введите путь к файлу\n")
path_2 = input("Выберите папку для сохранения")

wordDoc = (path_1)

wordDoc.save(path_2)


#
#
#
#
#
# def search():
#     for adress, dirs, files in os.walk(input('Введите путь старта\n')):
#         if adress == PATH_FOR_COPY:
#             continue
#         for file in files:
#             if file.endswith('.docx') or file.endswith('.doc'):
#                 yield os.path.join(adress, file)
#
#
# def read_from_pathdocx(path):
#     with open(path) as r:
#         for i in r:
#             if KEY_FOR_SEARCH in i:
#                 return copy(path)
#
#
# def copy(path):
#     file_name = path.split('\\')[-1]
#     count = 1
#     while True:
#         if os.path.isfile(os.path.join(PATH_FOR_COPY, file_name)):
#             if f'({count - 1})' in file_name:
#                 file_name = file_name.replace(f'({count - 1})', '')
#             file_name = f'({count}).'.join(file_name.split('.'))
#             count += 1
#         else:
#             break
#
#     shutil.copyfile(path, os.path.join(PATH_FOR_COPY, file_name))
#     print('Файл скопирован', file_name)
#
#
# for i in search():
#     try:
#         read_from_pathdocx(i)
#     except Exception as e:
#         with open(os.path.join(PATH_FOR_COPY, 'errors.txt'), 'a') as r:
#             r.write(str(e) + '\n' + i + '\n')
#
#
# practitioner_data = [
#     ("id", "Ф.И.О.", "Основание", "Файл №1", "Файл №2", "Файл №3", "Файл №4")
#
# ]
#
# with open("practitioner.csv", "w", newline='', encoding='cp1251', errors="ignore") as csvfile:
#     practitioner = csv.writer(csvfile, delimiter=";")
#     practitioner.writerows(
#         practitioner_data
#     )
