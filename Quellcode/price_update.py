import requests
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import scrolledtext, filedialog, messagebox
import pandas as pd

def get_price_automotive_connectors(article_number):
    search_url = f"https://www.automotive-connectors.com/en/search?search={article_number}"
    headers = {"User-Agent": "Mozilla/5.0"}

    try:
        response = requests.get(search_url, headers=headers, timeout=10)
        if response.status_code != 200:
            return f"Fehler beim Laden der Suchseite: {response.status_code}"

        soup = BeautifulSoup(response.text, 'html.parser')
        product_boxes = soup.find_all("div", class_="product-box")
        if not product_boxes:
            return "Keine Produkte gefunden."

        results = []
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
                            prices.append(f"{qty}: {price}")
                else:
                    prices.append("Kein Preis gefunden")

                results.append(f"{title} → {' | '.join(prices)}")

        return "\n".join(results) if results else "Kein passendes Produkt mit dieser Artikelnummer gefunden."
    except Exception as e:
        return f"Fehler: {e}"

def lade_excel_datei():
    dateipfad = filedialog.askopenfilename(filetypes=[("Excel-Dateien", "*.xls;*.xlsx")])
    if not dateipfad:
        return

    try:
        # Lade Datei, aber setze Header auf Zeile 6 (Index 5)
        df = pd.read_excel(dateipfad, header=6)
        df.columns = df.columns.str.strip()  # Whitespace entfernen

        artikel_spalte = next((col for col in df.columns if "Orderno" in col or "article" in col.lower()), None)

        if not artikel_spalte:
            messagebox.showerror("Fehler", f"Spalte mit Artikelnummer nicht gefunden!\nVerfügbare Spalten: {list(df.columns)}")
            return

        artikelnummern = df[artikel_spalte].dropna().astype(str)

        output.delete(1.0, tk.END)
        output.insert(tk.END, f"Datei geladen: {dateipfad}\nGefundene Artikelnummern: {len(artikelnummern)}\n\n")

        for nr in artikelnummern:
            output.insert(tk.END, f"→ {nr}\n")
            ergebnis = get_price_automotive_connectors(nr)
            output.insert(tk.END, ergebnis + "\n\n")
            output.update()

    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Lesen der Excel-Datei:\n{e}")

def suche_preis_einzel():
    artikelnummer = entry.get().strip()
    if not artikelnummer:
        output.delete(1.0, tk.END)
        output.insert(tk.END, "Bitte eine Artikelnummer eingeben.")
        return
    output.delete(1.0, tk.END)
    output.insert(tk.END, f"Suche nach: {artikelnummer}\n")
    ergebnis = get_price_automotive_connectors(artikelnummer)
    output.insert(tk.END, ergebnis)

# GUI
root = tk.Tk()
root.title("Preisabfrage Automotive Connectors")

frame_top = tk.Frame(root)
frame_top.pack(padx=10, pady=5)

tk.Label(frame_top, text="Artikelnummer:").grid(row=0, column=0, padx=5)
entry = tk.Entry(frame_top, width=40)
entry.grid(row=0, column=1, padx=5)
tk.Button(frame_top, text="Suche starten", command=suche_preis_einzel).grid(row=0, column=2, padx=5)
tk.Button(frame_top, text="Excel laden", command=lade_excel_datei).grid(row=0, column=3, padx=5)

output = scrolledtext.ScrolledText(root, width=80, height=25)
output.pack(padx=10, pady=10)

root.mainloop()
