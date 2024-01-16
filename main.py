from docx import Document
import pandas as pd
from docxtpl import DocxTemplate

# Чтение данных из Excel файла
excel_file_path = 'JV1901.xlsx'
student_data = pd.read_excel(excel_file_path, usecols="A,B,D")
doc = DocxTemplate("sample.docx")

# Количество строк данных и количество необходимых страниц
num_rows = len(student_data)
cards_per_page = 6
num_pages = -(-num_rows // cards_per_page)  # Округление вверх


for page in range(num_pages):
    doc = DocxTemplate("sample.docx")
    context = {}

    for card in range(1, cards_per_page + 1):
        row_index = page * cards_per_page + card - 1

        # Проверка, что row_index не выходит за границы DataFrame
        if row_index >= num_rows:
            break

        student_row = student_data.iloc[row_index]
        context[f'name_{card}'] = student_row[0]
        context[f'group_{card}'] = student_row[1]
        context[f'work_{card}'] = student_row[2]
        context[f'teacher'] = "Ilia Bogatyrev"  #Ilia Bogatyrev

    doc.render(context)
    doc.save(f'generated_doc{page}.docx')

from docx import Document

# Создаем новый документ, в который будем добавлять содержимое
merged_document = Document()

# Список имен файлов для объединения
files_to_merge = ['generated_doc0.docx', 'generated_doc1.docx', 'generated_doc2.docx',
                  'generated_doc3.docx', 'generated_doc4.docx']

for file_name in files_to_merge:
    # Открытие каждого документа
    sub_doc = Document(file_name)

    # Добавление каждого элемента из поддокумента в основной
    for element in sub_doc.element.body:
        merged_document.element.body.append(element)

# Сохраняем объединенный документ
merged_document.save('merged_document1901.docx')
print('Merged document')



