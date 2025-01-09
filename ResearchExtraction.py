import datetime
import os
import re
import tkinter as tk
from tkinter import filedialog
import openpyxl
import tabled
from docx import Document
import textract
import pandas as pd

from tabled.extract import extract_tables
from tabled.fileinput import load_pdfs_images
from tabled.inference.models import load_detection_models, load_recognition_models, load_layout_models


def extract_info_from_table(table):
    data = []
    for row in table.rows:
        if len(row.cells) == 2:
            cell1 = row.cells[0].text
            cell2 = row.cells[1].text
            data.append([cell1, cell2])
    return data


def dance(i, file_name):
    if i % 4 == 0:
        print(f'{file_name} в работе ヽ(￣▽￣)ﾉ')
    elif i % 4 == 1:
        print(f'{file_name} в работе \(▔∀▔)/')
    elif i % 4 == 2:
        print(f'{file_name} в работе ╰(￣▽￣)ノ')
    else:
        print(f'{file_name} в работе \(▔∀▔)/')
    i += 1
    return i


def get_from_doc(file_path):
    text = textract.process(file_path)
    data = text.decode('utf-8')
    data_df = pd.DataFrame(data.split('\n'), columns=['data'])
    return data_df


def get_from_docx(file_path):
    doc = Document(file_path)
    data = []
    for table in doc.tables:
        data.extend(extract_info_from_table(table))

    data_df = pd.DataFrame(data, columns=['data'])
    return data_df


def get_from_excel(file_path):
    if file_path.endswith('.xls'):
        # Чтение файла .xls с использованием xlrd
        excel = pd.ExcelFile(file_path, engine='xlrd')
    elif file_path.endswith('.xlsx'):
        # Чтение файла .xlsx с использованием openpyxl
        excel = pd.ExcelFile(file_path, engine='openpyxl')

    for sheet_name in excel.sheet_names:
        data_df = pd.read_excel(excel, sheet_name=sheet_name)
        data_df = data_df.dropna(axis=1, how='all')
    return data_df


def get_from_pdf(file_path):
    det_models, rec_models, layout_models = load_detection_models(), load_recognition_models(), load_layout_models()
    images, highres_images, names, text_lines = load_pdfs_images(file_path)

    page_results = extract_tables(images, highres_images, text_lines, det_models, layout_models, rec_models)
    data_df = pd.DataFrame()
    return data_df


def extract_researches(file_path, date, well_number):
    area_number = get_area_number(file_path)
    data = pd.DataFrame([[date, area_number, well_number]], columns=['Дата', 'Номер участка', 'Номер скважины'])
    if file_path.endswith('.doc'):
        data = get_from_doc(file_path)
    elif file_path.endswith('.docx'):
        data = get_from_docx(file_path)
    elif file_path.endswith('.xlsx') or file_path.endswith('.xls'):
        data = get_from_excel(file_path)
    elif file_path.endswith('.pdf'):
        data = get_from_pdf(file_path)
    return data


def write_to_excel(table_data, well_number, date_of_kvd, turn_to_red=False):
    workbook_name = 'Результат по участку ' + well_number[0] + 'А.xlsx'
    date_of_kvd = date_of_kvd.strftime('%d.%m.%Y') if date_of_kvd else None  # Пока лучше не придумал для обработки None
    try:
        workbook = openpyxl.load_workbook(workbook_name)
    except FileNotFoundError:
        workbook = openpyxl.Workbook()

    sheet = workbook.active

    row = sheet.max_row + 1  # Находим первую пустую строку для добавления данных
    col = 1  # Номер колонки для номера скважины
    if turn_to_red:
        sheet.cell(row=row, column=col, value=well_number).fill = openpyxl.styles.PatternFill(start_color='FF0000',
                                                                                              fill_type='solid')
        workbook.save(workbook_name)
        return None
    sheet.cell(row=row, column=col, value=well_number)
    col = 2  # Номер колонки для даты
    sheet.cell(row=row, column=col, value=date_of_kvd)
    for table_row in table_data:
        col = 2
        for info in table_row:
            col += 1
            sheet.cell(row=row, column=col, value=info)
        row += 1

    workbook.save(workbook_name)


def find_well_folder(root_path, current_path):
    current_dir = current_path
    while current_dir != root_path:
        folder_name = os.path.basename(current_dir)
        match = re.search(r'\d+[АAА]\d+', folder_name, re.IGNORECASE)
        if match:
            return match.group()
        current_dir = os.path.dirname(current_dir)
    return None


def get_area_number(file_path):
    # Ищем совпадение с номером участка в формате /{Численный номер участка}{Буква "А" или "A"}
    match = re.search(r'/(\d+)[АA]/', file_path)
    if match:
        area_number = match.group(1) + 'A'
        return area_number
    else:
        return None


def process_files_in_directory(folder_path):
    all_researches = []
    for root, dirs, files in os.walk(folder_path):
        well_folder_name = find_well_folder(folder_path, root)  # ищем номер скважины
        i = 0
        if well_folder_name:
            i = dance(i, well_folder_name)
            if "ГКИ" in root or 'ГРП' in root:  # убрать ГРП. Для тестирования пдф
                date = os.path.basename(root)  # Дата из названия папки
                try:
                    date = datetime.datetime.strptime(date, "%Y-%m-%d")  # Не помню, как конкретно оно там выглядит
                except ValueError:
                    date = None
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    if '~$' in file_path or '!Нету_' in file_path:
                        continue
                    research = extract_researches(file_path, date, well_folder_name)
                    all_researches.append(research)
