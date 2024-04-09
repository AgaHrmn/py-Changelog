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
        
    year_2023_path = r"download\file1.xlsx"
    # year_2024_path = r"download\file2.xlsx"
    
    sheets_2023 = load_data(year_2023_path)
    elektronarzedzia_2023 = sheets_2023[1]['Elektronarzedzia']
    ostrza_2023 = sheets_2023[2]['Ostrza']

#ostrza_2023
    sorted_column_ostrza_zastosowanie = sort_column_alphabetically(ostrza_2023, 9)
    capitalized_nazwa = capitalize_column_values(ostrza_2023, 1)
    capitalized_typ = capitalize_column_values(ostrza_2023, 2)
    capitalized_material = capitalize_column_values(ostrza_2023, 7)
    capitalized_zastosowanie = capitalize_column_values(ostrza_2023, 9)
    print("ostrza_processed")
    print(sorted_column_ostrza_zastosowanie)
    print(capitalized_nazwa)
    print(capitalized_typ)
    print(capitalized_material)
    print(capitalized_zastosowanie)

#elektronarzedzia_2023
    capitalized_nazwa = capitalize_column_values(elektronarzedzia_2023, 1)
    capitalized_typ_ostrza = capitalize_column_values(elektronarzedzia_2023, 3)
    capitalized_typ_silnika = capitalize_column_values(elektronarzedzia_2023, 5)
    capitalized_typ_zasilania = capitalize_column_values(elektronarzedzia_2023, 6)
    print("elektronarzedzia_processed")
    print(capitalized_nazwa)
    print(capitalized_typ_ostrza)
    print(capitalized_typ_silnika)
    print(capitalized_typ_zasilania)

