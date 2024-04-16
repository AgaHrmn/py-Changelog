from project_utensils import *
from columns_enum import *

if __name__ == "__main__":
    # urls = [
    #     "https://drive.google.com/uc?export=download&id=1wKpSpTx89dbU3SrKt-3-DqQ62kHOFHSX",
    #      "https://drive.google.com/uc?export=download&id=1oYKyW7flL53smo56W9srpFdoUuz6_42x"
    # ]

    # if not os.path.exists("download"):
    #     os.makedirs("download")

    # for i, url in enumerate(urls, start=1):
    #     file_path = os.path.join("download", f"file{i}.xlsx")
    #     download_file(url, file_path)
        
# 2023
    year_2023_path = r"download\file1.xlsx"
    sheets_2023 = load_data(year_2023_path)
    intro_2023 = sheets_2023[0]['Intro']
    elektronarzedzia_2023 = sheets_2023[1]['Elektronarzedzia']
    ostrza_2023 = sheets_2023[2]['Ostrza']

#ostrza_2023
    sorted_column_ostrza_zastosowanie_23 = sort_column_alphabetically(ostrza_2023, [9])
    processed_ostrza_23 = capitalize_column_values(sorted_column_ostrza_zastosowanie_23, [1,2,7,9])

#elektronarzedzia_2023
    processed_elektronarzedzia_23 = capitalize_column_values(elektronarzedzia_2023, [1,3,5,6])

    processed_2023 = {
        'Intro' : intro_2023,
        'Elektronarzedzia' : processed_elektronarzedzia_23,
        'Ostrza' : processed_ostrza_23
    }

# 2024
    year_2024_path = r"download\file2.xlsx"
    sheets_2024 = load_data(year_2024_path)
    intro_2024 = sheets_2024[0]['Intro']
    elektronarzedzia_2024 = sheets_2024[1]['Elektronarzedzia']
    ostrza_2024 = sheets_2024[2]['Ostrza']

#ostrza_2024
    sorted_column_ostrza_zastosowanie_24 = sort_column_alphabetically(ostrza_2024, [9])
    processed_ostrza_24 = capitalize_column_values(sorted_column_ostrza_zastosowanie_24, [1,2,7,9])

#elektronarzedzia_2024
    processed_elektronarzedzia_24 = capitalize_column_values(elektronarzedzia_2024, [1,3,5,6])

    processed_2024 = {
        'Intro' : intro_2024,
        'Elektronarzedzia' : processed_elektronarzedzia_24,
        'Ostrza' : processed_ostrza_24
    }

    changes_list_elektronarzedzia= generate_report(processed_2023, processed_2024, Elektronarzedzia, 'Elektronarzedzia')
    changes_list_ostrza= generate_report(processed_2023, processed_2024, Ostrza, 'Ostrza')
    changes_list = changes_list_elektronarzedzia + changes_list_ostrza

#create report dictionary
    intro = [str(f"Raport zmian w latach {extract_year(processed_2023)} i {extract_year(processed_2024)}")]
    report = {
        'Intro' : [[i] for i in intro],
        'Lista zmian' : [[change] for change in changes_list],
        'Elektronarzedzia' : merge_sheet_info(processed_2023, processed_2024, 'Elektronarzedzia', Elektronarzedzia, ElektronarzedziaReport),
        # 'Ostrza' merge_sheet_info(processed_2023, processed_2024, 'Ostrza', Ostrza): 
    }
    save_data(report, r"processed\processed.xlsx")


