import pandas as pd

df = pd.read_json("database.json", orient="split")

# Prüfe auf doppelte WN_SAP-Artikel-NR
dups_sap = df[df.duplicated(subset=["WN_SAP-Artikel-NR"], keep=False)]
print("Doppelte WN_SAP-Artikel-NR:")
print(dups_sap[["WN_SAP-Artikel-NR"]].drop_duplicates())

# Prüfe auf doppelte WN_HerstellerBestellnummer_1
dups_hersteller = df[df.duplicated(subset=["WN_HerstellerBestellnummer_1"], keep=False)]
print("\nDoppelte WN_HerstellerBestellnummer_1:")
print(dups_hersteller[["WN_HerstellerBestellnummer_1"]].drop_duplicates())