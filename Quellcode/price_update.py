import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, simpledialog, messagebox
from datetime import date
import requests, re
from bs4 import BeautifulSoup

# Mouser API Key hier eintragen, falls vorhanden!
MOUSER_API_KEY = ""

def format_value(value, field=None):
    # Datum formatieren
    if field and "Datum" in field and pd.notna(value) and str(value).strip() != "":
        try:
            date_val = pd.to_datetime(value, errors="coerce")
            if pd.notna(date_val):
                return date_val.strftime("%d.%m.%Y")
            else:
                return str(value)
        except Exception:
            return str(value)
    # Preis formatieren
    if field and "Preis" in field and pd.notna(value):
        try:
            val = re.sub(r"[^\d.,]", "", str(value)).replace(",", ".")
            price = float(val)
            return "{:.2f} €".format(price).replace(".", ",")
        except Exception:
            pass
    # Sonst Standard
    return "" if pd.isna(value) else str(value)

def sapnr_to_str(x):
    try:
        if pd.isna(x) or str(x).strip() == "" or str(x).lower() == "nan":
            return ""
        val = float(x)
        return str(int(val))
    except Exception:
        return str(x)

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

def mouser_price(article, MOUSER_API_KEY=""):
    if not MOUSER_API_KEY:
        return None
    payload = {
        "SearchByPartRequest": {
            "mouserPartNumber": article,
            "partSearchOptions": "None"
        }
    }
    try:
        r = requests.post(
            f"https://api.mouser.com/api/v1/search/partnumber?apiKey={MOUSER_API_KEY}",
            json=payload, headers={"Content-Type":"application/json"}, timeout=10
        )
        parts = r.json().get("SearchResults", {}).get("Parts", [])
        if not parts:
            return None
        part = parts[0]
        brk  = part["PriceBreaks"][-1]
        qty  = brk["Quantity"]
        raw  = brk["Price"]
        try:
            price_val = float(re.sub(r"[^\d,\.]", "", raw).replace(",", "."))
            price_txt = f"{price_val:.2f} €".replace(".", ",")
        except Exception:
            price_txt = raw
        return {
            "Datum": date.today().strftime("%d.%m.%Y"),
            "Preis": price_txt,
            "Losgröße": qty,
            "Quelle": "Mouser"
        }
    except Exception:
        return None

def search_and_show(df, search, search_cols):
    search = str(search).strip()
    search_df = df[search_cols].copy()
    if 'WN_SAP-Artikel-NR' in search_cols:
        search_df['WN_SAP-Artikel-NR'] = search_df['WN_SAP-Artikel-NR'].apply(
            lambda x: str(int(float(x))) if pd.notnull(x) and str(x).replace('.', '', 1).isdigit() else ""
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

def merge_results(db_rows, online_results_list):
    if db_rows is None or db_rows.empty:
        if online_results_list:
            data = {}
            for res in online_results_list:
                if res:
                    colname = res.get("Quelle", "Online")
                    data[colname] = [
                        format_value(res.get("Datum", ""), "Datum"),
                        format_value(res.get("Preis", ""), "Preis"),
                        res.get("Losgröße", ""),
                        res.get("Quelle", "")
                    ]
            if not data:
                return None
            return pd.DataFrame(data)
        else:
            return None

    df = db_rows.copy()

    # Formatierung nach Zeilennummer (0 = Datum, 1 = Preis)
    for col in df.columns:
        # Zeile 0: Datum
        df.iloc[0, df.columns.get_loc(col)] = format_value(df.iloc[0, df.columns.get_loc(col)], "Datum")
        # Zeile 1: Preis
        df.iloc[1, df.columns.get_loc(col)] = format_value(df.iloc[1, df.columns.get_loc(col)], "Preis")

    for res in online_results_list:
        if res:
            colname = res.get("Quelle", "Online")
            df[colname] = [
                format_value(res.get("Datum", ""), "Datum"),
                format_value(res.get("Preis", ""), "Preis"),
                res.get("Losgröße", ""),
                res.get("Quelle", ""),
            ]
    return df

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
        "WN_PinClass", "WN_PolCount_NUM", "WN_Color"
    ]

    if "WN_SAP-Artikel-NR" in df.columns:
        df["WN_SAP-Artikel-NR"] = df["WN_SAP-Artikel-NR"].apply(sapnr_to_str)

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

        # Hole Onlinequellen als Liste (AC + Mouser)
        online_results_list = []
        if artikelnummer and isinstance(artikelnummer, str) and artikelnummer and artikelnummer.lower() != 'nan':
            ac_res = ac_price(artikelnummer)
            if ac_res:
                online_results_list.append(ac_res)
            mouser_res = mouser_price(artikelnummer, MOUSER_API_KEY)
            if mouser_res:
                online_results_list.append(mouser_res)
        merged = merge_results(db_rows, online_results_list)
        if merged is not None and not merged.empty:
            show_table(merged)
        else:
            messagebox.showinfo(
                "Kein Treffer", f"Keine Zeile mit '{search}' gefunden (weder DB noch online)."
            )
    root.destroy()

if __name__ == "__main__":
    main()
