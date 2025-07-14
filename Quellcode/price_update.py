import pandas as pd, tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import date, datetime
import pathlib, requests, re
from bs4 import BeautifulSoup
from openpyxl import load_workbook

# ───────── KONFIG ─────────
SHEET_NAME  = "DB_4erDS"
HEADER_ROW  = 5
MOUSER_API_KEY = ""                # ← hier API-Key einsetzen

STATIC_COLS = [
    "Unnamed: 0", "Bauteilname", "WN_SAP-Artikel-NR", "Klasse",
    "Datenbankklasse\ndeutsch", "Datenbankbeschreibung \ndeutsch 2",
    "Beschreibung", "Kategorie", "WN_Hersteller", "WN_Hersteller\nBestellnr."
]
SAP_COL = "WN_SAP-Artikel-NR"
HBN_COL = "WN_Hersteller\nBestellnr."

PRICE_COLS   = ("Datum", "Preis", "Losgroesse", "Quelle")
VISIBLE_COLS = (*STATIC_COLS, *PRICE_COLS)

# ───────── Hilfs-Formatter ─────────
def fmt(val):
    if pd.isna(val) or val in ("", "nan"):          return ""
    if isinstance(val, (pd.Timestamp, datetime)):   return "" if val.year == 1900 else val.strftime("%d.%m.%Y")
    if isinstance(val, float):                      return str(int(val)) if val.is_integer() else f"{val:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")
    return str(val)

# ───────── Datenbank laden (alter, fester Ansatz; funktioniert) ─────────
def load_db(path: pathlib.Path) -> pd.DataFrame:
    xl     = pd.ExcelFile(path)
    sheet  = SHEET_NAME if SHEET_NAME in xl.sheet_names else xl.sheet_names[0]
    raw    = pd.read_excel(path, sheet_name=sheet, header=HEADER_ROW)
    raw    = raw.reindex(columns=STATIC_COLS + list(raw.columns[len(STATIC_COLS):]))

    static, price = raw.iloc[:, :len(STATIC_COLS)], raw.iloc[:, len(STATIC_COLS):]
    blocks = []
    for i in range(0, price.shape[1], 4):
        sub = price.iloc[:, i:i+4]
        if sub.shape[1] < 4: break
        sub.columns = PRICE_COLS
        blocks.append(sub)

    tidy = pd.concat([pd.concat([static]*len(blocks), ignore_index=True),
                      pd.concat(blocks,            ignore_index=True)], axis=1)\
            .dropna(subset=["Preis", "Datum"], how="all")

    tidy["Datum"] = pd.to_datetime(tidy["Datum"], errors="coerce")
    tidy["Preis"] = (tidy["Preis"].astype(str)
                     .str.replace(",", ".").str.extract(r"([\d\.]+)")[0]
                     .astype(float, errors="ignore"))
    return tidy.reset_index(drop=True)

# ───────── Online-Preise ─────────
def ac_price(article: str):
    """liest den *letzten* (größten) Mengen-Preis von automotive-connectors.com"""
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        # ─ 1. Suchtrefferseite ───────────────────────────────────────────
        search_url = f"https://www.automotive-connectors.com/en/search?search={article}"
        soup = BeautifulSoup(requests.get(search_url, headers=headers, timeout=10).text,
                             "html.parser")
        box = soup.find("div", class_="product-box")
        if not box:
            return None
        title = box.find("a", class_="product-name").text.strip()
        link  = box.find("a", class_="product-name")["href"]
        if not link.startswith("http"):
            link = "https://www.automotive-connectors.com" + link

        # versuche evtl. bereits vorhandene Tabelle (z. B. bei schnellen Seiten)
        rows = box.select("tr.product-block-prices-row")

        # ─ 2. Wenn nichts gefunden → Produktdetailseite aufrufen ─────────
        if not rows:
            prod_soup = BeautifulSoup(requests.get(link, headers=headers, timeout=10).text,
                                      "html.parser")
            rows = prod_soup.select("table.product-block-prices-grid tr.product-block-prices-row")
            if not rows:
                return None

        last = rows[-1]
        qty  = last.select_one(".product-block-prices-quantity")
        qty  = qty.text.strip() if qty else ""
        price_txt = last.find("td").get_text(" ", strip=True)
        price_num = re.search(r"[\d\.,]+", price_txt)
        if not price_num:
            return None
        price_val = float(price_num.group().replace(".", "").replace(",", "."))

        blank = {c: "" for c in STATIC_COLS}
        return {**blank, "Bauteilname": title, HBN_COL: article,
                "Datum": date.today(), "Losgroesse": qty,
                "Preis": price_val, "Quelle": "Automotive-Conn."}
    except Exception:
        return None

def mouser_price(article: str):
    if not MOUSER_API_KEY:
        return None
    payload = {"SearchByPartRequest":{"mouserPartNumber":article,"partSearchOptions":"None"}}
    try:
        r = requests.post(
            f"https://api.mouser.com/api/v1/search/partnumber?apiKey={MOUSER_API_KEY}",
            json=payload, headers={"Content-Type":"application/json"}, timeout=10)
        part        = r.json()["SearchResults"]["Parts"][0]
        qty, price  = part["PriceBreaks"][-1]["Quantity"], part["PriceBreaks"][-1]["Price"]
        price_val   = float(price.replace("€", "").replace(",", "."))
        blank = {c:"" for c in STATIC_COLS}
        return {**blank, "Bauteilname": part["ManufacturerPartNumber"],
                HBN_COL: article, "Datum": date.today(),
                "Losgroesse": qty, "Preis": price_val, "Quelle": "Mouser"}
    except Exception:
        return None

def online_rows(a): return [r for r in (ac_price(a), mouser_price(a)) if r]

# ───────── Matrix-Update  (unverändert) ─────────
def append_prices_to_workbook(wb_path, part, rows):
    if not rows:
        messagebox.showinfo("Info", "Keine Live-Preise geladen."); return
    wb, ws = load_workbook(wb_path, keep_vba=True), None
    if SHEET_NAME not in wb.sheetnames:
        messagebox.showerror("Fehler", f"Blatt '{SHEET_NAME}' fehlt."); return
    ws = wb[SHEET_NAME]

    last_col, new_col = ws.max_column, ws.max_column + 1
    for o in range(4):
        ws.cell(row=HEADER_ROW+1, column=new_col+o)._style = \
        ws.cell(row=HEADER_ROW+1, column=last_col-3+o)._style

    cnt = 0
    for r in rows:
        tgt = None
        for row in ws.iter_rows(min_row=HEADER_ROW+1, max_col=len(STATIC_COLS)):
            sap = str(row[STATIC_COLS.index(SAP_COL)].value).replace(".0","").strip()
            hbn = str(row[STATIC_COLS.index(HBN_COL)].value).strip().lower()
            if part == sap or part.lower() == hbn:
                tgt = row[0].row; break
        if tgt is None: continue
        ws.cell(row=HEADER_ROW+1, column=new_col).value = r["Datum"].strftime("%d.%m.%Y")
        ws.cell(row=tgt+1, column=new_col).value = fmt(r["Preis"])
        ws.cell(row=tgt+2, column=new_col).value = r["Losgroesse"]
        ws.cell(row=tgt+3, column=new_col).value = r["Quelle"]
        new_col += 4; cnt += 1
    wb.save(wb_path)
    messagebox.showinfo("Matrix", f"{cnt} Live-Preis(e) angehängt.")

# ───────── GUI ─────────
root = tk.Tk(); root.withdraw()
file = filedialog.askopenfilename(title="Excel-Datenbank wählen", filetypes=[("Excel","*.xls*")])
if not file: exit()
db_path = pathlib.Path(file)
try:
    db_df = load_db(db_path)
except Exception as e:
    messagebox.showerror("Fehler", str(e)); exit()

root.deiconify(); root.title("Preis-DB + Online-Preise"); root.geometry("1400x700")
top = tk.Frame(root); top.pack(padx=10, pady=5, fill="x")
tk.Label(top, text="Artikelnummer (SAP / Hersteller):").pack(side="left", padx=5)
entry = tk.Entry(top, width=35); entry.pack(side="left", padx=5)

tree = ttk.Treeview(root, columns=VISIBLE_COLS, show="headings", height=28)
for c in VISIBLE_COLS:
    tree.heading(c, text=c)
    tree.column(c, width=140 if c not in ("Beschreibung","Datenbankbeschreibung \ndeutsch 2") else 260, anchor="center")
tree.pack(fill="both", expand=True, padx=8, pady=(8,0))
hbar = ttk.Scrollbar(root, orient="horizontal", command=tree.xview)
tree.configure(xscrollcommand=hbar.set); hbar.pack(fill="x", padx=8, pady=(0,8))

last_part, live_rows = "", []

def do_search():
    global last_part, live_rows
    art = entry.get().strip()
    if not art: return
    tree.delete(*tree.get_children())

    sap = db_df[SAP_COL].astype(str).str.replace(r"\.0$","",regex=True).str.strip()
    hbn = db_df[HBN_COL].astype(str).str.strip().str.lower()
    db_rows = db_df[(sap == art) | (hbn == art.lower())].to_dict("records")

    part      = db_rows[0][HBN_COL] if db_rows else art
    live_rows = online_rows(part)
    template  = db_rows[0] if db_rows else {c:"" for c in STATIC_COLS}
    live_rows = [{**template, **lr} for lr in live_rows]
    last_part = part

    for r in db_rows + live_rows:
        tree.insert("", "end", values=[fmt(r.get(c,"")) for c in VISIBLE_COLS])
    if not tree.get_children():
        messagebox.showinfo("Info","Keine Daten gefunden.")

def export():
    if not tree.get_children():
        messagebox.showwarning("Keine Daten","Nichts zu exportieren."); return
    fname = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")])
    if not fname: return
    rows = [tree.item(i)["values"] for i in tree.get_children()]
    pd.DataFrame(rows, columns=VISIBLE_COLS).to_excel(fname, index=False)
    wb = load_workbook(fname); wb.active.auto_filter.ref = f"A1:{chr(64+len(VISIBLE_COLS))}1"; wb.save(fname)
    messagebox.showinfo("Export", "Datei gespeichert.")

tk.Button(top, text="Suche",                command=do_search).pack(side="left",  padx=5)
tk.Button(top, text="Export Tabelle",       command=export   ).pack(side="right", padx=5)
tk.Button(top, text="DB-Matrix aktualisieren",
          command=lambda: append_prices_to_workbook(db_path, last_part, live_rows))\
    .pack(side="right", padx=5)

root.mainloop()
