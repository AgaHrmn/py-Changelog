import os
import requests
from columns_enum import *
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
        capitalized_row = list(row) #copy of row to modify
        if i == 0:
            capitalized_rows.append(row)
        else:
            for column_num in column_nums:
                capitalized_row[column_num] = str(row[column_num]).title() #modified row
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

def extract_year(file, sheet_name='Intro'):
    year = int(extract_rows(file,sheet_name)[0][-1])
    return year

def extract_rows(file, sheet_name):
    rows= file[sheet_name]
    return rows
    
def show_changes(file_1, file_2, column_num, sheet_name):
    rows_file_1 = extract_rows(file_1,sheet_name) 
    rows_file_2 = extract_rows(file_2,sheet_name) 

    new_in_file_2 = []
    missing_in_file_2 = []
    changed_in_file_2 = []

    values_file_1 = []
    for row in rows_file_1:
        value = row[column_num]
        values_file_1.append(value)

    values_file_2 = []
    for row in rows_file_2:
        value = row[column_num]
        values_file_2.append(value)

    ids_file_2 = []
    for row in rows_file_2:
        id = row[0]
        ids_file_2.append(id)

    #find missing id in file 2
    for value in values_file_1:
        if value not in values_file_2:
            missing_in_file_2.append(value)

    #find new id in file 2
    for value in values_file_2:
        if value not in values_file_1:
            new_in_file_2.append(value)

    #find changed records in file 2 and return their ids
    for i, value in enumerate(values_file_2):
        if value not in values_file_1:
            changed_in_file_2.append(ids_file_2[i])

    changes = {
        'new_records': new_in_file_2,
        'missing_records': missing_in_file_2,
        'changed_records_ids' : changed_in_file_2
    }

    return changes

def get_changes_dict(file_1, file_2, enum, sheet_name): 
    changes_dict = {}
    for column in enum:
        changes_dict[column.name] = show_changes(file_1, file_2, column.value, sheet_name)
    return changes_dict

def generate_report(file_1, file_2, enum, sheet_name):
    list_of_changes=[]
    temp = ''
    if sheet_name == 'Elektronarzedzia':
        temp = 'elektronarzędzi'
    if sheet_name == 'Ostrza':
        temp = 'ostrzy'

    changes_dict = get_changes_dict(file_1, file_2, enum, sheet_name)
    discontinued = changes_dict['ID']['missing_records']
    new_products = changes_dict['ID']['new_records']

    modified = []
    for column_name, changes in changes_dict.items():
        if column_name == 'ID':
            continue
        if changes['changed_records_ids']:
            modified.append(f"Kolumna {column_name} zmieniła się dla {temp}: {changes['changed_records_ids']}")
    
    if discontinued:
        discontinued_str = "Narzędzia wycofane z oferty: " + ", ".join(discontinued) 
        list_of_changes.append(discontinued_str)
    if new_products:
        new_products_str = "Nowe produkty: " + ", ".join(new_products) 
        list_of_changes.append(new_products_str)
    if modified:
        for v in modified:
            list_of_changes.append(v)
    return list_of_changes

def save_data(data, output_file):
    if not os.path.exists("processed"):
        os.makedirs("processed")
    workbook = Workbook()

    for sheet_name, rows in data.items():
        if rows:
            print(f"Saving {len(rows)} rows for sheet: {sheet_name}")
            worksheet = workbook.create_sheet(title=sheet_name)
            for row in rows:
                if isinstance(row, list):  #check if row is a list
                    worksheet.append(row)
                else:
                    print(f"Skipping row in {sheet_name}: {row}. It's not in the expected format.")
        else:
            print(f"No data found for sheet: {sheet_name}")

    workbook.save(filename=output_file)
    

def merge_sheet_info(file_1, file_2, sheet_name, enum, enum_rep):
    rows_1 = extract_rows(file_1, sheet_name)[1:]
    rows_2 = extract_rows(file_2, sheet_name)[1:]
    changes_dict = get_changes_dict(file_1, file_2, enum, sheet_name)
    new_ids= []
    missing_ids= []
    changed_ids= []
    merged_rows = []

    temp_row = []
    for column in enum_rep:
        temp_row.append(column.name)
    merged_rows.append(temp_row)

    temp_row = []
    for missing_id in changes_dict['ID']['missing_records']:
        temp_row = ['wycofane', missing_id] + []* len(enum_rep)
        missing_ids.append(missing_id)
        merged_rows.append(temp_row)

    temp_row = []
    for new_id in changes_dict['ID']['new_records']:
        temp_row = ['nowe', new_id] + []* len(enum_rep)
        new_ids.append(new_id)
        merged_rows.append(temp_row)
    
    temp_row = []
    for column_name, changes in changes_dict.items():
        for changed_id in changes_dict[column_name]['changed_records_ids']:
            changed_ids.append(changed_id)
    changed_ids_set = set(changed_ids)
    for id in changed_ids_set:
        temp_row = ['zmodyfikowane', id]
        merged_rows.append(temp_row)
    
    processed_ids = new_ids+missing_ids+changed_ids
    temp_row = []
    for row in rows_2:
        id = row[0]
        if id not in processed_ids:
            temp_row = ['-', id]
        else :
            continue
        merged_rows.append(temp_row)    

    for row in merged_rows[1:]:
        for cells_23 in rows_1:
            if row[1] == cells_23[0]:
                insert_values(row, cells_23, enum_rep, enum,"_23")
        
        for cells_24 in rows_2:
            if row[1] == cells_24[0]:
                insert_values(row, cells_24, enum_rep, enum, "_24")  
    return merged_rows

def insert_values(row, cells, enum_rep, enum, suffix):
    for column in enum:
        if column.name == "ID":
            continue
        else:
            column_rep_name = f"{column.name}{suffix}"
            full_name = getattr(enum_rep, column_rep_name)

            if row[0] == "wycofane" or row[0] == "nowe": 
                row.insert(full_name.value, None) 
                
            row.insert(full_name.value, cells[column.value])

            # if column.name in cells:
            #     row.insert(full_name.value, cells[column.value])
            # else:
            #     row.insert(full_name.value, None)

            