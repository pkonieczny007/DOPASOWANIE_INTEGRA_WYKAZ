import pandas as pd
from openpyxl import load_workbook

# Wczytaj dane z plików
wykaz_df = pd.read_excel("wykaz1.xlsx")
elementy_df = pd.read_excel("elementy1.xlsx")

# Dodaj nową kolumnę na wynik dopasowania
wykaz_df['propozycja'] = ""

# Funkcja do wyciągania klucza z nazwy lub referencji do dopasowania
def extract_key(name):
    if isinstance(name, str) and '_' in name:  # Upewnij się, że jest tekstem i zawiera "_"
        parts = name.split('_')
        if len(parts) >= 3:
            return '_'.join(parts[2:4])  # Wyciąga "SL60372020_p1"
    return None

# Tworzymy klucze do szybkiego porównania
elementy_df['Key'] = elementy_df['Referencja'].apply(extract_key)
element_keys = set(elementy_df['Key'].dropna())  # Usuń wartości None

# Przetwarzanie każdego wiersza z wykazu
for idx, row in wykaz_df.iterrows():
    name_key = extract_key(row['Nazwa'])
    if name_key and name_key in element_keys:
        match = elementy_df[elementy_df['Key'] == name_key]
        wykaz_df.at[idx, 'propozycja'] = match['Referencja'].values[0]

# Zapisz wynik do pliku Excel
wykaz_df.to_excel("propozycje1.xlsx", index=False)

# Otwórz zapisany plik Excel, aby dopasować szerokość kolumn
workbook = load_workbook("propozycje1.xlsx")
worksheet = workbook.active

# Dopasowanie szerokości kolumn do zawartości
for column in worksheet.columns:
    max_length = 0
    column_letter = column[0].column_letter  # Pobierz literę kolumny
    for cell in column:
        try:
            max_length = max(max_length, len(str(cell.value)))
        except:
            pass
    adjusted_width = (max_length + 2)
    worksheet.column_dimensions[column_letter].width = adjusted_width

# Zapisz plik z dopasowanymi szerokościami kolumn
workbook.save("propozycje1.xlsx")
