import pandas as pd

# Wczytaj dane z plików, wymuszając odczyt jako tekst
wykaz_df = pd.read_excel("wykaz.xlsx", dtype=str)
elementy_df = pd.read_excel("elementy1.xlsx", dtype=str)

# Dodaj nową kolumnę na wynik dopasowania
wykaz_df['propozycja'] = ""

# Upewnij się, że w kolumnie Referencja nie ma wartości NaN (zamień NaN na pusty ciąg)
elementy_df['Referencja'] = elementy_df['Referencja'].fillna('')

# Iterujemy przez każdy wiersz z wykazu
for idx, row in wykaz_df.iterrows():
    nazwa = row.get('Nazwa')
    # Sprawdzenie, czy wartość nie jest pusta lub NaN
    if not isinstance(nazwa, str) or nazwa.strip() == "":
        continue
    nazwa = nazwa.strip()

    # Wektorowe sprawdzanie: tworzymy maskę, gdzie w Referencja występuje szukany tekst
    mask = elementy_df['Referencja'].str.contains(nazwa, na=False)
    if mask.any():
        # Pobieramy pierwszy dopasowany wiersz
        match = elementy_df.loc[mask].iloc[0]
        wykaz_df.at[idx, 'propozycja'] = match.get('Referencja1', "")

# Zapisz wynik do pliku propozycje.xlsx
wykaz_df.to_excel("propozycje.xlsx", index=False)
print("Dopasowanie zakończone. Wyniki zapisano w 'propozycje.xlsx'.")
