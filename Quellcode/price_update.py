import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, simpledialog, messagebox
from datetime import date
import requests, re
from bs4 import BeautifulSoup

def ac_price(article):
    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        url     = f"https://www.automotive-connectors.com/en/search?search={article}"
        soup    = BeautifulSoup(requests.get(url, headers=headers, timeout=10).text, "html.parser")
        box = soup.find("div", class_="product-box")
        if not box:
            return None
        rows = box.select("tr.product-block-prices-row")
        if not rows:
            link = box.find("a", class_="product-name")["href"]
            if not link.startswith("http"):
                link = "https://www.automotive-connectors.com" + link
            detail_html = requests.get(link, headers=headers, timeout=10).text
            rows = BeautifulSoup(detail_html, "html.parser").select("tr.product-block-prices-row")
            if not rows:
                return None
        last = rows[-1]
        qty_txt = last.select_one(".product-block-prices-quantity").get_text(strip=True)
        qty     = re.sub(r"[^\d]", "", qty_txt) or ""
        price_cell = last.find_all("td", class_="product-block-prices-cell")[-1]
        price_txt  = price_cell.get_text(" ", strip=True)
        return {
            "Datum": date.today().strftime("%d.%m.%Y"),
            "Preis": price_txt,
            "Losgröße": qty,
            "Quelle": "Automotive-Connectors"
        }
    except Exception:
        return None

def search_and_show(df, search, search_cols):
    search = str(search).strip()
    search_df = df[search_cols].copy()
    if 'WN_SAP-Artikel-NR' in search_cols:
        search_df['WN_SAP-Artikel-NR'] = search_df['WN_SAP-Artikel-NR'].apply(
            lambda x: str(int(x)) if pd.notnull(x) and str(x) != 'nan' else ""
        )
    if 'WN_HerstellerBestellnummer_1' in search_cols:
        search_df['WN_HerstellerBestellnummer_1'] = search_df['WN_HerstellerBestellnummer_1'].astype(str).map(
            lambda x: x.strip() if x != 'nan' else ""
        )
    mask = search_df.apply(lambda row: any(search == str(cell) for cell in row), axis=1)
    indices = mask[mask].index
    if len(indices) == 0:
        return None
    start = indices[0]
    end = min(start + 4, len(df))
    return df.iloc[start:end]

def merge_results(db_rows, online_results):
    # Fall 1: Datenbanktreffer -> hänge Online-Spalten rechts an
    if db_rows is not None and not db_rows.empty:
        db_rows = db_rows.copy()
        db_rows["Online_Datum"] = ""
        db_rows["Online_Preis"] = ""
        db_rows["Online_Losgroesse"] = ""
        db_rows["Online_Quelle"] = ""
        # Trage die Online-Ergebnisse nur in der ersten DB-Zeile ein (wie Excel bei einem Preisblock)
        if online_results:
            db_rows.loc[db_rows.index[0], "Online_Datum"] = online_results.get("Datum", "")
            db_rows.loc[db_rows.index[0], "Online_Preis"] = online_results.get("Preis", "")
            db_rows.loc[db_rows.index[0], "Online_Losgroesse"] = online_results.get("Losgröße", "")
            db_rows.loc[db_rows.index[0], "Online_Quelle"] = online_results.get("Quelle", "")
        return db_rows
    # Fall 2: Kein Datenbanktreffer, aber Online-Treffer
    elif online_results:
        return pd.DataFrame([{
            "Online_Datum": online_results.get("Datum", ""),
            "Online_Preis": online_results.get("Preis", ""),
            "Online_Losgroesse": online_results.get("Losgröße", ""),
            "Online_Quelle": online_results.get("Quelle", "")
        }])
    else:
        return None

def show_table(df):
    df = df.replace({pd.NA: '', None: ''}).fillna('')
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
    file = filedialog.askopenfilename(
        title="Excel-Datei wählen", filetypes=[("Excel-Dateien", "*.xls*")]
    )
    if not file:
        messagebox.showinfo("Abgebrochen", "Keine Datei gewählt.")
        return

    sheet = "DB_4erDS"
    df = pd.read_excel(file, sheet_name=sheet, header=6)
    cols_to_drop = [
        "Unnamed: 0", "Unnamed: 17", "Unnamed: 18",
        "Unnamed: 19", "Unnamed: 20", "Unnamed: 21", "Unnamed: 22",
    ]
    df = df.drop(columns=[col for col in cols_to_drop if col in df.columns])
    search_cols = ["WN_SAP-Artikel-NR", "WN_HerstellerBestellnummer_1"]
    missing_cols = [col for col in search_cols if col not in df.columns]
    if missing_cols:
        messagebox.showerror(
            "Fehler", f"Folgende Spalten fehlen in der Tabelle: {missing_cols}"
        )
        return

    while True:
        search = simpledialog.askstring(
            "Suche", "Artikelnummer oder SAP-Nummer eingeben:", parent=root
        )
        if not search:
            break
        db_rows = search_and_show(df, search, search_cols)
        artikelnummer = None
        if db_rows is not None:
            if 'WN_HerstellerBestellnummer_1' in db_rows.columns:
                artikelnummer = db_rows.iloc[0]['WN_HerstellerBestellnummer_1']
        else:
            artikelnummer = search
        online_results = None
        if artikelnummer and isinstance(artikelnummer, str) and artikelnummer and artikelnummer.lower() != 'nan':
            online_results = ac_price(artikelnummer)
        merged = merge_results(db_rows, online_results)
        if merged is not None and not merged.empty:
            show_table(merged)
        else:
            messagebox.showinfo(
                "Kein Treffer", f"Keine Zeile mit '{search}' gefunden (weder DB noch online)."
            )
    root.destroy()

if __name__ == "__main__":
    main()
