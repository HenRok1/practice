import openpyxl
from docx import Document


wordDoc = Document('main.docx')
book = openpyxl.open("../Liza/main.xlsx", read_only=True)
sheet = book.active

list_key = []

for row in range(2, sheet.max_row + 1):
    key = sheet[row][2].value

    list_key.append(key)


tables = wordDoc.tables # Получить таблицу, установленную в файле
table = tables[0] # Получить первую таблицу в файле

df = [[''for i in range(len(tables[0].columns))]for j in range(len(tables[0].rows))]

for table in wordDoc.tables:
    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            df[i][j] = cell.text
            for type_name in range(2, sheet.max_row + 1):
                key = sheet[type_name][2].value
                if key in cell.text:
                        name = sheet[type_name][1].value
                        quest = input("Введите " + name + " ")
                        cell.text = cell.text.replace(key, quest)

wordDoc.save("example.docx")