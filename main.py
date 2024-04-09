from project_utensils import *

if __name__ == "__main__":
    urls = [
        "https://drive.google.com/uc?export=download&id=1wKpSpTx89dbU3SrKt-3-DqQ62kHOFHSX",
        "https://drive.google.com/uc?export=download&id=1oYKyW7flL53smo56W9srpFdoUuz6_42x"
    ]

    if not os.path.exists("download"):
        os.makedirs("download")

    for i, url in enumerate(urls, start=1):
        file_path = os.path.join("download", f"file{i}.xlsx")
        download_file(url, file_path)
        