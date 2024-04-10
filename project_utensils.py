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
    workbook = load_workbook(filename=xlsx_file)
    sheets_list = []

    for sheet_name in workbook.sheetnames:
        worksheet = workbook[sheet_name]
        rows = list(worksheet.iter_rows(values_only=True))
        sheets_list.append({sheet_name: rows})

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
        if rows:
            print(f"Saving {len(rows)} rows for sheet: {sheet_name}")
            worksheet = workbook.create_sheet(title=sheet_name)
            for row in rows:
                worksheet.append(row)
        else:
            print(f"No data found for sheet: {sheet_name}")

    workbook.save(filename=output_file)

def show_changes(file_1, file_2, sheet_name='Elektronarzedzia'):
    rows_file_1 = file_1[sheet_name][1:]  # Exclude header
    rows_file_2 = file_2[sheet_name][1:]

    new_in_file_2 = []
    missing_in_file_2 = []

    ids_file_1 = []
    for row in rows_file_1:
        id = row[0]
        ids_file_1.append(id)

    ids_file_2 = []
    for row in rows_file_2:
        id = row[0]
        ids_file_2.append(id)

    print(ids_file_1)
    print(ids_file_2)

    #find missing id in file 2
    for id in ids_file_1:
        if id not in ids_file_2:
            missing_in_file_2.append(id)

    #find new id in file 2
    for id in ids_file_2:
        if id not in ids_file_1:
            new_in_file_2.append(id)

    changes = {
        'new_records': new_in_file_2,
        'missing_records': missing_in_file_2
    }

    return changes
    