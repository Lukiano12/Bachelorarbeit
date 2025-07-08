import requests
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import scrolledtext, filedialog, messagebox, ttk
import pandas as pd

ergebnisse_liste = []  # Für den späteren Excel-Export

def get_price_automotive_connectors(article_number):
    search_url = f"https://www.automotive-connectors.com/en/search?search={article_number}"
    headers = {"User-Agent": "Mozilla/5.0"}

    try:
        response = requests.get(search_url, headers=headers, timeout=10)
        if response.status_code != 200:
            return None, f"Fehler beim Laden der Suchseite: {response.status_code}"

        soup = BeautifulSoup(response.text, 'html.parser')
        product_boxes = soup.find_all("div", class_="product-box")
        if not product_boxes:
            return None, "Keine Produkte gefunden."

        for box in product_boxes:
            title_tag = box.find("a", class_="product-name")
            title = title_tag.text.strip() if title_tag else "Kein Titel"
            if article_number.lower() in title.lower():
                price_table = box.find("table", class_="product-block-prices-grid")
                prices = []
                if price_table:
                    rows = price_table.find_all("tr", class_="product-block-prices-row")
                    for row in rows:
                        qty_tag = row.find("span", class_="product-block-prices-quantity")
                        price_td = row.find_all("td", class_="product-block-prices-cell")
                        if qty_tag and price_td:
                            qty = qty_tag.text.strip()
                            price_div = price_td[0].find("div")
                            price = price_div.text.strip().replace('\n', ' ') if price_div else price_td[0].text.strip()
                            prices.append((title, article_number, qty, price))
                else:
                    prices.append((title, article_number, "-", "Kein Preis gefunden"))
                return prices, None

        return None, "Kein passendes Produkt mit dieser Artikelnummer gefunden."
    except Exception as e:
        return None, f"Fehler: {e}"

def lade_excel_datei():
    global ergebnisse_liste
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
        output.delete(1.0, tk.END)
        output.insert(tk.END, f"Datei geladen: {dateipfad}\nGefundene Artikelnummern: {len(artikelnummern)}\n\n")

        tree.delete(*tree.get_children())  # Vorherige Ergebnisse löschen
        ergebnisse_liste.clear()

        for index, nr in enumerate(artikelnummern, 1):
            output.insert(tk.END, f"{index}. Artikel: {nr}\n")
            output.update()
            preise, fehler = get_price_automotive_connectors(nr)
            if fehler:
                output.insert(tk.END, fehler + "\n\n")
            else:
                for title, article, qty, price in preise:
                    tree.insert("", tk.END, values=(title, article, qty, price))
                    ergebnisse_liste.append({"Artikelname": title, "Artikelnummer": article, "Menge": qty, "Preis": price})
                output.insert(tk.END, "✓ Daten hinzugefügt\n\n")
            output.update()
    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Lesen der Excel-Datei:\n{e}")

def suche_preis_einzel():
    artikelnummer = entry.get().strip()
    if not artikelnummer:
        messagebox.showwarning("Hinweis", "Bitte eine Artikelnummer eingeben.")
        return
    output.delete(1.0, tk.END)
    tree.delete(*tree.get_children())
    ergebnisse_liste.clear()

    preise, fehler = get_price_automotive_connectors(artikelnummer)
    if fehler:
        output.insert(tk.END, fehler)
    else:
        for title, article, qty, price in preise:
            tree.insert("", tk.END, values=(title, article, qty, price))
            ergebnisse_liste.append({"Artikelname": title, "Artikelnummer": article, "Menge": qty, "Preis": price})

def exportiere_excel():
    if not ergebnisse_liste:
        messagebox.showinfo("Keine Daten", "Keine Daten zum Exportieren vorhanden.")
        return
    pfad = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel-Datei", "*.xlsx")])
    if not pfad:
        return
    df_export = pd.DataFrame(ergebnisse_liste)
    df_export.to_excel(pfad, index=False)
    messagebox.showinfo("Erfolg", f"Daten erfolgreich exportiert nach:\n{pfad}")

# --- GUI ---
root = tk.Tk()
root.title("Preisabfrage Automotive Connectors")

frame_top = tk.Frame(root)
frame_top.pack(padx=10, pady=5, fill="x")

tk.Label(frame_top, text="Artikelnummer:").grid(row=0, column=0, padx=5)
entry = tk.Entry(frame_top, width=40)
entry.grid(row=0, column=1, padx=5)
tk.Button(frame_top, text="Suche starten", command=suche_preis_einzel).grid(row=0, column=2, padx=5)
tk.Button(frame_top, text="Excel laden", command=lade_excel_datei).grid(row=0, column=3, padx=5)
tk.Button(frame_top, text="Exportieren", command=exportiere_excel).grid(row=0, column=4, padx=5)

output = scrolledtext.ScrolledText(root, width=100, height=10)
output.pack(padx=10, pady=5)

# --- Tabelle für Ergebnisse ---
frame_table = tk.Frame(root)
frame_table.pack(padx=10, pady=10, fill="both", expand=True)

tree = ttk.Treeview(frame_table, columns=("Artikelname", "Artikelnummer", "Menge", "Preis"), show="headings")
tree.heading("Artikelname", text="Artikelname")
tree.heading("Artikelnummer", text="Artikelnummer")
tree.heading("Menge", text="Menge")
tree.heading("Preis", text="Preis")

tree.column("Artikelname", width=250)
tree.column("Artikelnummer", width=150)
tree.column("Menge", width=100)
tree.column("Preis", width=100)
tree.pack(fill="both", expand=True)

root.mainloop()
