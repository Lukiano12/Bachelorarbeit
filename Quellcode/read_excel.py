import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, simpledialog, messagebox

def search_and_show(df, search, search_cols):
    search = str(search).strip()
    
    # Kopie zum Bearbeiten
    search_df = df[search_cols].copy()
    
    # SAP-Nummer-Spalte (float) als int und als String zum Vergleich
    if 'WN_SAP-Artikel-NR' in search_cols:
        # Nur für Zeilen, wo kein NaN
        search_df['WN_SAP-Artikel-NR'] = search_df['WN_SAP-Artikel-NR'].apply(
            lambda x: str(int(x)) if pd.notnull(x) else ""
        )
    
    # HerstellerNummer als String und getrimmt
    if 'WN_HerstellerBestellnummer_1' in search_cols:
        search_df['WN_HerstellerBestellnummer_1'] = search_df['WN_HerstellerBestellnummer_1'].astype(str).map(lambda x: x.strip() if x != 'nan' else "")
    
    # Jetzt beide Spalten (als String) durchsuchen: exakter Vergleich und Teilstring
    mask = search_df.apply(lambda row: any(search == str(cell) for cell in row), axis=1)
    indices = mask[mask].index

    if len(indices) == 0:
        return None
    start = indices[0]
    end = min(start + 4, len(df))
    return df.iloc[start:end]

def show_table(df):
    # NaN und None zu leeren Strings
    df = df.replace({pd.NA: '', None: ''}).fillna('')
    # Nur Spalten anzeigen, in denen mindestens ein Wert steht
    df = df.loc[:, (df != '').any(axis=0)]

    root = tk.Tk()
    root.title("Suchergebnis")

    frame = tk.Frame(root)
    frame.pack(fill="both", expand=True)

    tree = ttk.Treeview(frame, columns=list(df.columns), show="headings")
    for col in df.columns:
        tree.heading(col, text=str(col))
        tree.column(col, width=120, anchor="center")
    for _, row in df.iterrows():
        tree.insert("", tk.END, values=list(row))
    tree.pack(fill="both", expand=True)

    root.mainloop()



def main():
    root = tk.Tk()
    root.withdraw()
    file = filedialog.askopenfilename(title="Excel-Datei wählen", filetypes=[("Excel-Dateien", "*.xls*")])
    if not file:
        messagebox.showinfo("Abgebrochen", "Keine Datei gewählt.")
        return
    
    sheet = "DB_4erDS"
    df = pd.read_excel(file, sheet_name=sheet, header=6)

    cols_to_drop = [
    "Unnamed: 0",
    "Unnamed: 17",
    "Unnamed: 18",
    "Unnamed: 19",
    "Unnamed: 20",
    "Unnamed: 21",
    "Unnamed: 22",
]

    # Diese Spalten entfernen, falls vorhanden:
    df = df.drop(columns=[col for col in cols_to_drop if col in df.columns])

    


    # Spaltennamen genau wie im Screenshot!
    search_cols = ["WN_SAP-Artikel-NR", "WN_HerstellerBestellnummer_1"]
    missing_cols = [col for col in search_cols if col not in df.columns]
    if missing_cols:
        messagebox.showerror("Fehler", f"Folgende Spalten fehlen in der Tabelle: {missing_cols}")
        return

    while True:
        # Suchdialog
        search = simpledialog.askstring("Suche", "Artikelnummer oder SAP-Nummer eingeben:", parent=root)
        if not search:
            break
        res = search_and_show(df, search, search_cols)
        if res is not None:
            show_table(res)
        else:
            messagebox.showinfo("Kein Treffer", f"Keine Zeile mit '{search}' in den Spalten {search_cols} gefunden.")

    root.destroy()

if __name__ == "__main__":
    main()
