import os
import requests
from openpyxl import Workbook
from openpyxl import load_workbook


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


def sort_column_alphabetically(rows, column_nums):
    sorted_rows = []
    for row in rows:
        sorted_row = list(row) #copy of row to modify
        for column_num in column_nums:
            column_values = row[column_num]

            if len(column_values) > 1:
                column_text = str(column_values)
                words = column_text.split('\n')
                words.sort()
                sorted_string = '\n'.join(words)
                sorted_row[column_num] = sorted_string #modified row
        sorted_rows.append(sorted_row)     
    return sorted_rows


def capitalize_column_values(rows, column_nums):
    capitalized_rows = []

    for i, row in enumerate(rows):
        capitalized_row = list(row) #copy of row to mdidy
        if i == 0:
            capitalized_rows.append(row)
        else:
            for column_num in column_nums:
                capitalized_row[column_num] = str(row[column_num]).upper() #modified row
            capitalized_rows.append(capitalized_row)
    return capitalized_rows

def clear_white_space(rows):
    clean_rows = []
    for row in rows:
        clean_row = list(row) #copy of row to mdidy
        for cell in row:
            cleaned = str(cell).strip()
            clean_row.append(cleaned)
        clean_rows.append(clean_row)
    return clean_rows


def save_data(data, output_file):
    if not os.path.exists("processed"):
        os.makedirs("processed")
    workbook = Workbook()

    for sheet_name, rows in data.items():
        worksheet = workbook.create_sheet(title=sheet_name)
        for row in rows:
            worksheet.append(row)

    workbook.save(filename=output_file)