import os
import requests
from openpyxl import Workbook
from openpyxl import load_workbook
from lxml import etree

def download_file(url, save_path):
    response = requests.get(url)
    try:
        with open(save_path, 'wb') as file:
            file.write(response.content)
    except FileNotFoundError:
        print("Plik nie został znaleziony.")
    except IOError:
        print("Wystąpił błąd podczas operacji na pliku.")
    else:
        print(f"Plik pobrany pomyślnie.")

def load_data(xlsx_file):
    sheets_list = []

    workbook = load_workbook(filename=xlsx_file)
    sheets = workbook.sheetnames

    for sheet in sheets:
        worksheet = workbook[sheet]
        sheet_dict = {}
        rows = []

        for row in worksheet.iter_rows(values_only=True):
            rows.append(row)

        sheet_dict[sheet]=rows
        sheets_list.append(sheet_dict)

    return sheets_list


def save_data(data, output_file):
    workbook = Workbook()

    for sheet_dict in data:
        for sheet_name, rows in sheet_dict.items():
            worksheet = workbook.create_sheet(title=sheet_name)
            for row in rows:
                worksheet.append(row)

    workbook.save(filename=output_file)

def sort_column_alphabetically(rows, column_num):
    sorted_column_values = []
    for row in rows:
        column_values = row[column_num]

        if len(column_values) > 1:
            column_text = str(column_values)
            words = column_text.split('\n')
            words.sort()
            sorted_string = '\n'.join(words)
            sorted_column_values.append(sorted_string)
        else:
            return sorted_column_values.append(column_values)
        
    return sorted_column_values

def capitalize_column_values(rows, column_num):
    capitalized_column_values = []

    for i, row in enumerate(rows):
        if i == 0:
            capitalized_column_values.append(row[column_num])
        else:
            capitalized_column_values.append(str(row[column_num]).upper())
    return capitalized_column_values