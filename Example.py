import os
import re
from datetime import datetime

import numpy as np
import pandas as pd

import ExcelFinder
import ExcelWriter


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


def get_area_number(file_path):
    # Ищем совпадение с номером участка в формате /{Численный номер участка}{Буква "А" или "A"}
    match = re.search(r'/(\d+)[АA]/', file_path)
    if match:
        area_number = match.group(1) + 'A'
        return area_number
    else:
        return None


def find_well_folder(root_path, current_path):
    current_dir = current_path
    while current_dir != root_path:
        folder_name = os.path.basename(current_dir)
        match = re.search(r'\d+[АAА]\d+', folder_name, re.IGNORECASE)
        if match:
            return match.group()
        current_dir = os.path.dirname(current_dir)
    return None


def extract_reports(folder_path, output_folder_path=os.getcwd()):
    for root, _, files in os.walk(folder_path):
        well_folder_name = find_well_folder(folder_path, root)  # ищем номер скважины
        i = 0
        for file_name in files:
            file_path = os.path.join(root, file_name)
            if '~' in file_path or '!Нету_' in file_path or all(map(lambda x: x not in file_name.lower(), mask_list)):
                continue

            date = datetime.strptime(file_name[4:11], '%m.%Y').date()
            if start_date and end_date:
                if date < start_date or date > end_date:
                    continue

            i = dance(i, file_name)

            if file_path.endswith('.xls'):
                # Чтение файла .xls с использованием xlrd
                excel = pd.ExcelFile(file_path, engine='xlrd')
            elif file_path.endswith('.xlsx'):
                # Чтение файла .xlsx с использованием openpyxl
                excel = pd.ExcelFile(file_path, engine='openpyxl')

            for sheet_name in excel.sheet_names:
                if 'Форма 1' in sheet_name:
                    area_number = get_area_number(file_path)
                    report_df = pd.read_excel(excel, sheet_name=sheet_name)
                    report_df = report_df.dropna(axis=1, how='all')
                    report_df = get_table(report_df)
                    report_df.columns = range(len(report_df.columns))
                    report_df.index = range(len(report_df.index))
                    report_df = ExcelFinder.fill_merged_cells(report_df)
                    report_df.insert(loc=0, column='Дата', value=date)
                    all_reports_df.append(report_df)

    ExcelWriter.write_df(all_reports_df, os.path.join(output_folder_path, f'{area_number}.xlsx'), 'Все МЭРы',
                         area_number)


def get_headers(sheet):
    start_index = ExcelFinder.find_data_start_index(sheet)
    end_index = ExcelFinder.find_well_start_index(sheet)
    table = sheet.iloc[start_index:end_index]
    table = ExcelFinder.slice_df(table)
    return table


def get_table(sheet):
    headers = get_headers(sheet)
    headers['Характер скважины'] = ['Характер скважины'] + [np.nan] * (len(headers) - 1)
    wells = ExcelFinder.get_well_purpose(sheet.iloc[ExcelFinder.find_data_start_index(sheet):])
    table = pd.concat([headers, wells])

    # Убираем скважины, номера которых равны нулю либо отсутствуют
    table = ExcelFinder.remove_invalid_well_number_rows(table)

    return table
