
import os
import shutil
import json
import csv
from docx import Document


file_name1 = input("Введите имя файла №1 (.pdf или .jpg):\n")
file_name2 = input("Введите имя файла №2 (.docx):\n")
file_name3 = input("Введите имя файла №3 (.pdf или .jpg):\n")
file_name4 = input("Введите имя файла №4 (.docx):\n")
folder_save = input("Выберите папку для сохранения:\n")

path_file1 = os.path.abspath(file_name1)
path_file2 = os.path.abspath(file_name2)
path_file3 = os.path.abspath(file_name3)
path_file4 = os.path.abspath(file_name4)
path_save = os.path.abspath(folder_save)


Doc1 = shutil.copy(file_name1, path_save)
wordDoc2 = Document(path_file2)
Doc3 = shutil.copy(file_name3, path_save)
wordDoc4 = Document(path_file4)
# wordDoc1.save(path_save)
wordDoc2.save(path_save)
# wordDoc3.save(path_save)
wordDoc4.save(path_save)



id = input("Введите id:\n")
doc_name = input("Введите название документа:\n")
base = input("Введите основание:\n")
practitioner_data = [
    ("id", "Документ", "Основание", "Файл №1", "Файл №2", "Файл №3", "Файл №4")
]
practitioner_data2 = [
    (id, doc_name, base, path_file1, path_file2, path_file3, path_file4)
]


with open("practitioner.csv", "w", newline='', encoding='cp1251', errors="ignore") as csvfile:
    practitioner = csv.writer(csvfile, delimiter=";")
    practitioner.writerows(
        practitioner_data
    )

with open("practitioner.csv", "a", newline='', encoding='cp1251', errors="ignore") as csvfile:
    practitioner2 = csv.writer(csvfile, delimiter=";")
    practitioner2.writerows(
        practitioner_data2
    )
