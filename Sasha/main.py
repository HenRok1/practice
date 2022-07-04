
import os
import shutil
import json
import csv
from docx import Document


file_name1 = input("Введите имя файла №1:\n")
file_name2 = input("Введите имя файла №2:\n")
file_name3 = input("Введите имя файла №3:\n")
file_name4 = input("Введите имя файла №4:\n")
path_save = input("Выберите папку для сохранения:\n")

path_file1 = os.path.abspath(file_name1)
path_file2 = os.path.abspath(file_name2)
path_file3 = os.path.abspath(file_name3)
path_file4 = os.path.abspath(file_name4)


Doc1 = shutil.copy(file_name1, path_save)
wordDoc2 = Document(path_file2)
Doc3 = shutil.copy(file_name3, path_save)
wordDoc4 = Document(path_file4)
# wordDoc1.save(path_save)
wordDoc2.save(path_save)
# wordDoc3.save(path_save)
wordDoc4.save(path_save)



practitioner_data = [
    ("id", "Документ", "Основание", "Файл №1", "Файл №2", "Файл №3", "Файл №4")

]


with open("practitioner.csv", "w", newline='', encoding='cp1251', errors="ignore") as csvfile:
    practitioner = csv.writer(csvfile, delimiter=";")
    practitioner.writerows(
        practitioner_data
    )

path_file = input("Введите название файла для импорта в таблицу:\n")
with open(path_file) as csvfile:
    result = json.load(csvfile)
#print(result)

id = input("Введите id\n")
doc_name = input("Введите название документа:\n")
base = input("Введите основание:\n")


asks = result[id][doc_name][base][path_file1][path_file2][path_file3][path_file4]
#print(asks)
for a in asks:
    with open("practitioner.csv", "a", newline='',  encoding='cp1251', errors="ignore") as csvfile:
        practitioner = csv.writer(csvfile, delimiter=";")
        practitioner.writerows(
            asks
        )

