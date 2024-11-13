import pandas as pd

# Wczytaj dane z plików
wykaz_df = pd.read_excel("wykaz.xlsx")
elementy_df = pd.read_excel("elementy.xlsx")

# Dodaj nową kolumnę na wynik dopasowania
wykaz_df['propozycja'] = ""

# Funkcja do wyciągania klucza z nazwy lub referencji do dopasowania
def extract_key(name):
    if isinstance(name, str) and '_' in name:  # Upewnij się, że jest tekstem i zawiera "_"
        parts = name.split('_')
        # Połącz pierwsze cztery części jako unikalny klucz
        return '_'.join(parts[:4])
    return None  # Zwróć None dla wartości bez "_"

# Tworzymy klucze do szybkiego porównania
elementy_df['Key'] = elementy_df['Referencja'].apply(extract_key)
element_keys = set(elementy_df['Key'].dropna())  # Usuń wartości None

# Przetwarzanie każdego wiersza z wykazu
for idx, row in wykaz_df.iterrows():
    # Wyciągnij klucz dla bieżącego wiersza
    name_key = extract_key(row['Nazwa'])
    
    # Sprawdź, czy klucz istnieje na liście elementów
    if name_key and name_key in element_keys:
        # Znajdź i przypisz pełną nazwę z elementy_df
        match = elementy_df[elementy_df['Key'] == name_key]
        wykaz_df.at[idx, 'propozycja'] = match['Referencja'].values[0]

# Zapisz wynik do propozycje.xlsx
wykaz_df.to_excel("propozycje.xlsx", index=False)
