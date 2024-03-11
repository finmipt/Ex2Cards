import pandas as pd
from docxtpl import DocxTemplate
from docx import Document
import glob
import os

# Чтение данных и подготовка к созданию документов
excel_file_path = 'JV2902.xlsx'
student_data = pd.read_excel(excel_file_path, usecols="A,B,D")
num_rows = len(student_data)
cards_per_page = 1
num_pages = -(-num_rows // cards_per_page)  # Округление вверх

for row in range(num_rows):
    doc = DocxTemplate("sample_one.docx")
    context = {}
    student_row = student_data.iloc[row]
    context[f'name'] = student_row[0]
    context[f'group'] = student_row[1]
    context[f'work'] = student_row[2]
    context[f'teacher'] = "Ilia Bogatyrev"   # 'Oksana Pavelkovitš'
    context[f'materials'] = 'Vihikud, kalkulaator'
    doc.render(context)
    doc.save(f'generated_page_{row}.docx')


# Сбор имен файлов для объединения
files_to_merge = glob.glob('generated_page*.docx')

# Создание объединенного документа
merged_document = Document()
for file_name in files_to_merge:
    sub_doc = Document(file_name)
    for element in sub_doc.element.body:
        merged_document.element.body.append(element)

# Сохранение объединенного документа
merged_document.save('merged_document_all_2902.docx')

# Удаление исходных файлов
for file_name in files_to_merge:
    os.remove(file_name)

print('Merged document created and original documents deleted.')