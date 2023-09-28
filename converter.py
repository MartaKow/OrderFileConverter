import pandas as pd
import os

folder_location = input("Podaj ścieżkę do folderu zawierającego pliki .xls z zamówieniami do gazetki: ")
txt_location = input("Podaj ścieżkę do folderu, w którym mają zostać zapisane pliki .txt: ")


xls_files = [file for file in os.listdir(folder_location) if file.endswith(".xls")]

for xls_file in xls_files:
    data = pd.read_html(os.path.join(folder_location, xls_file))
    tabela1 = data[0]
    tabela2 = data[1]

    nazwa_pliku = f"{str(tabela1.loc[4][3])} {str(tabela1.loc[7][3])} {str(tabela1.loc[5][3])}"

    txt_file_path = os.path.join(txt_location, f"{nazwa_pliku}.txt")

    with open(txt_file_path, "w") as file:
        for index, row in tabela2.iterrows():
            if row[2] == 0:
                print(f"Uzupełnij kod EAN w pliku {nazwa_pliku}!")
                file.write(f"BRAKUJĄCY KOD;{ilosc}\n")
            else:
                kod_ean = int(row[2])
                ilosc = str(row[6])[:-2]
                file.write(f"{kod_ean};{ilosc}\n")

print(f"PLIKI .TXT ZOSTAŁY UTWORZONE :)")
input("Naciśnij Enter, aby zakończyć program.")
