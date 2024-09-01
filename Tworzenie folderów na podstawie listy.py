import pandas as pd
import os

file_path = r'C:\Users\pkawk\OneDrive\Pulpit\foldery.xlsx'
df = pd.read_excel(file_path, header=None)

# Zakładając, że nazwy folderów znajdują się w pierwszej kolumnie
folder_names = df.iloc[:, 0].tolist()
#folder_names = df['NazwaKolumny'].tolist()
# W powyższym przykładzie NazwaKolumny jest nazwą kolumny, której nagłówek znajduje się w pierwszym wierszu. 
#To jest bardziej przejrzyste, gdy nazwy kolumn są znane i w ten sposób możesz uniknąć problemu z pomyłkowym 
#uwzględnieniem nagłówka jako jednej z wartości.
# Wyznacz kolumnę na podstawie pozycji (indeksu)
# Wybieramy kolumnę o indeksie 2 (trzecia kolumna)
#column_data = df.iloc[:, 2].tolist()

# Ścieżka, w której chcesz stworzyć foldery (możesz podać dowolną ścieżkę)
base_directory = r'C:\Users\pkawk\OneDrive\Pulpit\Test' # np. 'C:/Users/TwojeImie/Documents'

# Przechodzimy do wybranej ścieżki
os.chdir(base_directory)

# Tworzenie folderów
for folder_name in folder_names:
    try:
        os.makedirs(folder_name)
        print(f"Folder '{folder_name}' został utworzony.")
    except FileExistsError:
        print(f"Folder '{folder_name}' już istnieje.")