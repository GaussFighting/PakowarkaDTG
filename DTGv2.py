import xlrd
import os
import glob
import re

def sanitize_folder_name(name):
    # Usuwamy znaki niedozwolone w nazwach folderów
    return re.sub(r'[\\/*?:"<>|]', "", name)[:10]

# Pobranie bieżącego katalogu roboczego
current_dir = os.getcwd()

# Wyszukiwanie pliku z rozszerzeniem zaczynającym się na ".x"
excel_files = glob.glob(os.path.join(current_dir, '*.x*'))

if not excel_files:
    print("Nie znaleziono pliku Excel w bieżącym katalogu.")
    exit()

# Zakładamy, że interesuje nas pierwszy znaleziony plik
excel_file = excel_files[0]
print(f"Znaleziono plik Excel: {excel_file}")

# Otwieranie pliku za pomocą xlrd (dla .xls)
wb = xlrd.open_workbook(excel_file)
sheet = wb.sheet_by_index(0)  # Pierwszy arkusz

print(f"Liczba wierszy w arkuszu: {sheet.nrows}")

# Pobranie danych z kolumny 'A' (kolumna 0 w xlrd)
column_index = 0
folder_names = []
start_collecting = False

for row_idx in range(sheet.nrows):
    cell_value = sheet.cell_value(row_idx, column_index)
    
    # Konwertujemy na string i sprawdzamy, czy zaczyna się od "Order"
    if str(cell_value).strip().startswith("Order"):
        print("Znaleziono 'Order'. Zaczynamy zbierać dane...")
        start_collecting = True
        continue

    if start_collecting:
        if cell_value:  # Jeśli komórka nie jest pusta
            # Konwertujemy na int i z powrotem na string, aby usunąć ".0"
            folder_name = sanitize_folder_name(str(int(cell_value)))  # Ograniczamy długość nazwy do 10 znaków
            folder_names.append(folder_name)
            print(f"Dodano nazwę folderu: {folder_name}")
        else:  # Jeśli napotykamy pustą komórkę, przerywamy zbieranie
            print("Pusta komórka. Kończę zbieranie danych.")
            break

# Sprawdź, czy znaleziono jakieś nazwy folderów
if not folder_names:
    print("Nie znaleziono żadnych nazw folderów.")
else:
    print(f"Znaleziono następujące nazwy folderów: {folder_names}")

# Tworzenie folderów na podstawie wczytanych nazw
for name in folder_names:
    main_folder = f"{name}_0"
    print(f"Tworzę folder: {main_folder}")
    os.makedirs(main_folder, exist_ok=True)
    
    # Tworzenie dodatkowych podfolderów w nowo utworzonym folderze
    os.makedirs(os.path.join(main_folder, "PP Pics"))
    os.makedirs(os.path.join(main_folder, "DTG"))
    os.makedirs(os.path.join(main_folder, "Symulacja"))
    os.makedirs(os.path.join(main_folder, "VAS"))
    
    # Tworzenie podfolderów w folderze DTG
    dtg_folder = os.path.join(main_folder, "DTG")
    os.makedirs(os.path.join(dtg_folder, "Pliki"))
    os.makedirs(os.path.join(dtg_folder, "Ripy"))
    os.makedirs(os.path.join(dtg_folder, "Archiwum"))
    
    print(f"Utworzono strukturę folderów dla: {name}")

print("Tworzenie folderów zakończone.")

# Usuwanie pliku Excel po zakończeniu
try:
    os.remove(excel_file)
    print(f"Usunięto plik Excel: {excel_file}")
except Exception as e:
    print(f"Wystąpił błąd podczas usuwania pliku: {e}")
