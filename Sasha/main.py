
import json
import csv
from docx import Document

path_name = input("Введите путь к файлу:\n")
path_save = input("Выберите папку для сохранения:\n")

wordDoc = Document(path_name)
wordDoc.save(path_save)





practitioner_data = [
    ("id", "Имя файла", "Основание", "Файл №1", "Файл №2", "Файл №3", "Файл №4")

]


with open("practitioner.csv", "w", newline='', encoding='cp1251', errors="ignore") as csvfile:
    practitioner = csv.writer(csvfile, delimiter=";")
    practitioner.writerows(
        practitioner_data
    )

path_file = input("Введите название файла для импорта в таблицу:\n")
with open(path_file) as csvfile:
    result = json.load(csvfile)
print(result)

asks = result["id"]["Имя файла"]["Основание"]["Файл №1"]["Файл №2"]["Файл №3"]["Файл №4"]
print(asks)
for a in asks:
    with open("practitioner.csv", "a", newline='',  encoding='cp1251', errors="ignore") as csvfile:
        practitioner = csv.writer(csvfile, delimiter=";")
        practitioner.writerows(
            asks
        )

