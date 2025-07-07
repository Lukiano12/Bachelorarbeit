import requests
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import scrolledtext

def get_price_automotive_connectors(article_number):
    search_url = f"https://www.automotive-connectors.com/en/search?search={article_number}"
    headers = {
        "User-Agent": "Mozilla/5.0"
    }

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

        if results:
            return "\n".join(results)
        else:
            return "Kein passendes Produkt mit dieser Artikelnummer gefunden."
    except Exception as e:
        return f"Fehler: {e}"

# --- GUI Teil ---
def suche_preise():
    artikelnummer = entry.get().strip()
    if not artikelnummer:
        output.delete(1.0, tk.END)
        output.insert(tk.END, "Bitte eine Artikelnummer eingeben.")
        return
    output.delete(1.0, tk.END)
    output.insert(tk.END, "Suche läuft...\n")
    ergebnis = get_price_automotive_connectors(artikelnummer)
    output.delete(1.0, tk.END)
    output.insert(tk.END, ergebnis)

root = tk.Tk()
root.title("Preisabfrage Automotive Connectors")

tk.Label(root, text="Artikelnummer:").pack(padx=10, pady=5)
entry = tk.Entry(root, width=40)
entry.pack(padx=10, pady=5)
tk.Button(root, text="Preise suchen", command=suche_preise).pack(padx=10, pady=5)

output = scrolledtext.ScrolledText(root, width=60, height=15)
output.pack(padx=10, pady=10)

root.mainloop()