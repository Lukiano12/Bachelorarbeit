import requests
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

def get_price_automotive_connectors(article_number):
    search_url = f"https://www.automotive-connectors.com/en/search?search={article_number}"
    headers = {"User-Agent": "Mozilla/5.0"}

    try:
        response = requests.get(search_url, headers=headers, timeout=10)
        if response.status_code != 200:
            return [(article_number, "", "", f"Fehler: {response.status_code}")]

        soup = BeautifulSoup(response.text, 'html.parser')
        product_boxes = soup.find_all("div", class_="product-box")
        if not product_boxes:
            return [(article_number, "", "", "Keine Produkte gefunden")]

        for box in product_boxes:
            title_tag = box.find("a", class_="product-name")
            title = title_tag.text.strip() if title_tag else "Kein Titel"
            if article_number.lower() in title.lower():
                price_table = box.find("table", class_="product-block-prices-grid")
                if price_table:
                    rows = price_table.find_all("tr", class_="product-block-prices-row")
                    if rows:
                        last_row = rows[-1]
                        qty_tag = last_row.find("span", class_="product-block-prices-quantity")
                        price_td = last_row.find_all("td", class_="product-block-prices-cell")
                        if qty_tag and price_td:
                            qty = qty_tag.text.strip()
                            price_div = price_td[0].find("div")
                            price = price_div.text.strip().replace('\n', ' ') if price_div else price_td[0].text.strip()
                            return [(title, article_number, qty, price)]
                return [(title, article_number, "-", "Kein Preis gefunden")]
        return [(article_number, "", "", "Kein passendes Produkt")]
    except Exception as e:
        return [(article_number, "", "", f"Fehler: {e}")]

def lade_excel_datei():
    dateipfad = filedialog.askopenfilename(filetypes=[("Excel-Dateien", "*.xls;*.xlsx")])
    if not dateipfad:
        return

    try:
        df = pd.read_excel(dateipfad, header=6)
        df.columns = df.columns.str.strip()

        zielspalte = "Manufacturer-\nOrderno. 1"
        if zielspalte not in df.columns:
            messagebox.showerror("Fehler", f"Spalte '{zielspalte}' nicht gefunden!\nVerfügbare Spalten: {list(df.columns)}")
            return

        artikelnummern = df[zielspalte].dropna().astype(str).unique()
        daten = []

        for index, nr in enumerate(artikelnummern, 1):
            tree.insert('', tk.END, values=("Suche...", nr, "", ""))
            root.update()
            result = get_price_automotive_connectors(nr)
            for i, (name, nr, menge, preis) in enumerate(result):
                if i == 0:
                    tree.delete(tree.get_children()[-1])  # ersetze Platzhalter
                tree.insert('', tk.END, values=(name, nr, menge, preis))
                daten.append((name, nr, menge, preis))

        global tabelle_daten
        tabelle_daten = daten

    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Lesen der Excel-Datei:\n{e}")

def suche_preis_einzel():
    artikelnummer = entry.get().strip()
    if not artikelnummer:
        messagebox.showwarning("Eingabe fehlt", "Bitte eine Artikelnummer eingeben.")
        return

    tree.delete(*tree.get_children())
    results = get_price_automotive_connectors(artikelnummer)
    for name, nr, menge, preis in results:
        tree.insert('', tk.END, values=(name, nr, menge, preis))

    global tabelle_daten
    tabelle_daten = results

def exportiere_excel():
    if not tabelle_daten:
        messagebox.showwarning("Keine Daten", "Es sind keine Daten zum Exportieren vorhanden.")
        return

    pfad = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel-Datei", "*.xlsx")])
    if not pfad:
        return

    df = pd.DataFrame(tabelle_daten, columns=["Artikelname", "Artikelnummer", "Menge", "Preis"])
    try:
        df.to_excel(pfad, index=False)
        messagebox.showinfo("Export erfolgreich", f"Datei gespeichert unter:\n{pfad}")
    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Exportieren:\n{e}")

# --- GUI Aufbau ---
root = tk.Tk()
root.title("Automotive Connectors Preisabfrage")
root.geometry("1000x600")

frame_top = tk.Frame(root)
frame_top.pack(padx=10, pady=10)

tk.Label(frame_top, text="Artikelnummer:").grid(row=0, column=0, padx=5)
entry = tk.Entry(frame_top, width=40)
entry.grid(row=0, column=1, padx=5)
tk.Button(frame_top, text="Einzelsuche", command=suche_preis_einzel).grid(row=0, column=2, padx=5)
tk.Button(frame_top, text="Excel laden", command=lade_excel_datei).grid(row=0, column=3, padx=5)
tk.Button(frame_top, text="Export nach Excel", command=exportiere_excel).grid(row=0, column=4, padx=5)

# Tabelle (Treeview)
columns = ("Artikelname", "Artikelnummer", "Menge", "Preis")
tree = ttk.Treeview(root, columns=columns, show="headings", height=25)
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=200, anchor='center')
tree.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

tabelle_daten = []  # globale Liste zum Exportieren

root.mainloop()
