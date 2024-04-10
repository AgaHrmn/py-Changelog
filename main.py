from project_utensils import *

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
    print("ostrza_2023_processed")

#elektronarzedzia_2023
    processed_elektronarzedzia_23 = capitalize_column_values(elektronarzedzia_2023, [1,3,5,6])
    print("elektronarzedzia_2023_processed")

#save to new file
    processed_2023 = {
        'Intro' : intro_2023,
        'Elektronarzedzia' : processed_elektronarzedzia_23,
        'Ostrza' : processed_ostrza_23
    }

    # save_data(processed_2023, r"processed\processed_2023.xlsx")

#################################################################################################################

# 2024
    year_2024_path = r"download\file2.xlsx"
    sheets_2024 = load_data(year_2024_path)
    intro_2024 = sheets_2024[0]['Intro']
    elektronarzedzia_2024 = sheets_2024[1]['Elektronarzedzia']
    ostrza_2024 = sheets_2024[2]['Ostrza']

#ostrza_2024
    sorted_column_ostrza_zastosowanie_24 = sort_column_alphabetically(ostrza_2024, [9])
    processed_ostrza_24 = capitalize_column_values(sorted_column_ostrza_zastosowanie_24, [1,2,7,9])
    print("ostrza_2024_processed")

#elektronarzedzia_2024
    processed_elektronarzedzia_24 = capitalize_column_values(elektronarzedzia_2024, [1,3,5,6])
    print("elektronarzedzia_2024_processed")

#save to new file
    processed_2024 = {
        'Intro' : intro_2024,
        'Elektronarzedzia' : processed_elektronarzedzia_24,
        'Ostrza' : processed_ostrza_24
    }
    
#find missing or added records in file 2 based on ID column
    changes = show_changes(processed_2023, processed_2024)
    print(changes)

