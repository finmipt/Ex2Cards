import pandas as pd
from docxtpl import DocxTemplate
from docx import Document
import glob
import os

# Чтение данных и подготовка к созданию документов
excel_file_path = 'JV1901.xlsx'
student_data = pd.read_excel(excel_file_path, usecols="A,B,D")
num_rows = len(student_data)
cards_per_page = 6
num_pages = -(-num_rows // cards_per_page)  # Округление вверх

# Генерация отдельных документов
for page in range(num_pages):
    doc = DocxTemplate("sample.docx")
    context = {}
    for card in range(1, cards_per_page + 1):
        row_index = page * cards_per_page + card - 1
        if row_index >= num_rows:
            break
        student_row = student_data.iloc[row_index]
        context[f'name_{card}'] = student_row[0]
        context[f'group_{card}'] = student_row[1]
        context[f'work_{card}'] = student_row[2]
        context[f'teacher'] = 'Oksana Pavelkovitš' #"Ilia Bogatyrev"
    doc.render(context)
    doc.save(f'generated_doc{page}.docx')

# Сбор имен файлов для объединения
files_to_merge = glob.glob('generated_doc*.docx')

# Создание объединенного документа
merged_document = Document()
for file_name in files_to_merge:
    sub_doc = Document(file_name)
    for element in sub_doc.element.body:
        merged_document.element.body.append(element)

# Сохранение объединенного документа
merged_document.save('merged_document0102OP.docx')

# Удаление исходных файлов
for file_name in files_to_merge:
    os.remove(file_name)

print('Merged document created and original documents deleted.')