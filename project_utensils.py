import os
import requests


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