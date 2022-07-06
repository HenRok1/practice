
import os
import shutil
import csv

practitioner_data = [
        ("id", "Документ", "Основание", "Файл №1", "Файл №2", "Файл №3", "Файл №4")
    ]
with open("practitioner.csv", "w", newline='', encoding='cp1251', errors="ignore") as csvfile:
    practitioner = csv.writer(csvfile, delimiter=";")
    practitioner.writerows(
        practitioner_data
    )
num_of_docs = int(input("Введите количество документов:\n"))

for num_of_docs in range(0, num_of_docs):

    file_name1 = input("Введите имя файла №1 (с расширением .pdf или .jpg):\n")
    file_name2 = input("Введите имя файла №2 (с расширением .doc или .docx):\n")
    file_name3 = input("Введите имя файла №3 (с расширением  .pdf или .jpg):\n")
    file_name4 = input("Введите имя файла №4 (с расширением .doc или .docx):\n")
    folder_save = input("Введите название папки для сохранения:\n")
    os.mkdir(folder_save)

    path_file1 = os.path.abspath(file_name1)
    path_file2 = os.path.abspath(file_name2)
    path_file3 = os.path.abspath(file_name3)
    path_file4 = os.path.abspath(file_name4)
    path_save = os.path.abspath(folder_save)

    Doc1 = shutil.copy(file_name1, folder_save)
    Doc2 = shutil.copy(file_name2, folder_save)
    Doc3 = shutil.copy(file_name3, folder_save)
    Doc4 = shutil.copy(file_name4, folder_save)

    id = input("Введите id\n")
    doc_name = input("Введите название документа:\n")
    base = input("Введите основание:\n")
    practitioner_data2 = [
        (id, doc_name, base, path_file1, path_file2, path_file3, path_file4)
    ]

    with open("practitioner.csv", "a", newline='', encoding='cp1251', errors="ignore") as csvfile:
        practitioner2 = csv.writer(csvfile, delimiter=";")
        practitioner2.writerows(
            practitioner_data2
        )
