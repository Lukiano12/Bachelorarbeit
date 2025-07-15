import pandas as pd, tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import date, datetime
import pathlib, requests, re
from bs4 import BeautifulSoup
from openpyxl import load_workbook

# ───────── KONFIG ─────────
SHEET_NAME     = "DB_4erDS"
HEADER_ROW     = 5
MOUSER_API_KEY = ""  # ← Hier ggf. Mouser‑API‑Key eintragen, wenn genutzt

STATIC_COLS = [
    "Unnamed: 0", "Bauteilname", "WN_SAP-Artikel-NR", "Klasse",
    "Datenbankklasse\ndeutsch", "Datenbankbeschreibung \ndeutsch 2",
    "Beschreibung", "Kategorie", "WN_Hersteller", "WN_Hersteller\nBestellnr."
]
SAP_COL, HBN_COL = "WN_SAP-Artikel-NR", "WN_Hersteller\nBestellnr."
PRICE_COLS     = ("Datum", "Preis", "Losgroesse", "Quelle")
VISIBLE_COLS   = (*STATIC_COLS, *PRICE_COLS)

# ───────── Hilfs‑Formatter ─────────
def fmt(v):
    if pd.isna(v) or v in ("", "nan"):
        return ""
    if isinstance(v, (pd.Timestamp, datetime)):
        return "" if v.year == 1900 else v.strftime("%d.%m.%Y")
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else f"{v:,.2f} €"\
            .replace(",", "X").replace(".", ",").replace("X", ".")
    return str(v)

# ───────── Datenbank laden ─────────
def load_db(path: pathlib.Path) -> pd.DataFrame:
    xl    = pd.ExcelFile(path)
    sheet = SHEET_NAME if SHEET_NAME in xl.sheet_names else xl.sheet_names[0]
    raw   = pd.read_excel(path, sheet_name=sheet, header=HEADER_ROW)
    raw   = raw.reindex(columns=STATIC_COLS + list(raw.columns[len(STATIC_COLS):]))

    static, price = raw.iloc[:, :len(STATIC_COLS)], raw.iloc[:, len(STATIC_COLS):]
    blocks = []
    for i in range(0, price.shape[1], 4):
        sub = price.iloc[:, i:i+4]
        if sub.shape[1] < 4:
            break
        sub.columns = PRICE_COLS
        blocks.append(sub)

    tidy = pd.concat([
        pd.concat([static]*len(blocks), ignore_index=True),
        pd.concat(blocks,           ignore_index=True)
    ], axis=1).dropna(subset=["Preis","Datum"], how="all")

    tidy["Datum"] = pd.to_datetime(tidy["Datum"], errors="coerce")
    # Datenbankpreise bleiben numerisch, nur Online-Preise werden als String eingefügt
    tidy["Preis"] = (
        tidy["Preis"].astype(str)
             .str.replace(",", ".")
             .str.extract(r"([\d\.]+)")[0]
             .astype(float, errors="ignore")
    )
    return tidy.reset_index(drop=True)

# ───────── Automotive‑Connectors (Einzelstück + roher Preis‑String) ─────────
def ac_price(article: str):
    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        url     = f"https://www.automotive-connectors.com/en/search?search={article}"
        soup    = BeautifulSoup(requests.get(url, headers=headers, timeout=10).text,
                                "html.parser")

        box = soup.find("div", class_="product-box")
        if not box:
            return None

        title = box.find("a", class_="product-name").get_text(strip=True)

        # Preistabelle (Trefferliste oder Detailseite)
        rows = box.select("tr.product-block-prices-row")
        if not rows:
            link = box.find("a", class_="product-name")["href"]
            if not link.startswith("http"):
                link = "https://www.automotive-connectors.com" + link
            detail_html = requests.get(link, headers=headers, timeout=10).text
            rows = BeautifulSoup(detail_html, "html.parser")\
                      .select("tr.product-block-prices-row")
            if not rows:
                return None

        # Immer den letzten Break (größte Losgröße) nehmen
        last = rows[-1]

        # Menge extrahieren
        qty_txt = last.select_one(".product-block-prices-quantity").get_text(strip=True)
        qty     = re.sub(r"[^\d]", "", qty_txt) or ""

        # Preis‑String komplett übernehmen (inkl. "€0.07*/pcs.")
        price_cell = last.find_all("td", class_="product-block-prices-cell")[-1]
        price_txt  = price_cell.get_text(" ", strip=True)

        # Baustein-Dikt
        blank = {c: "" for c in STATIC_COLS}
        return {
            **blank,
            "Bauteilname": title,
            HBN_COL      : article,
            "Datum"      : date.today(),
            "Losgroesse" : qty,
            "Preis"      : price_txt,
            "Quelle"     : "Automotive‑Conn."
        }
    except Exception:
        return None

# ───────── Mouser – unverändert (numerisch) ─────────
def mouser_price(article: str):
    if not MOUSER_API_KEY:
        return None
    payload = {"SearchByPartRequest": {
        "mouserPartNumber": article,
        "partSearchOptions": "None"
    }}
    try:
        r     = requests.post(
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
        price_val = float(re.sub(r"[^\d,\.]", "", raw).replace(",", "."))

        blank = {c: "" for c in STATIC_COLS}
        return {
            **blank,
            "Bauteilname": part["ManufacturerPartNumber"],
            HBN_COL      : article,
            "Datum"      : date.today(),
            "Losgroesse" : qty,
            "Preis"      : price_val,
            "Quelle"     : "Mouser"
        }
    except Exception:
        return None

# ───────── GUI‑Funktionen ─────────
rows_cache, last_part = [], ""

def do_search():
    art = entry.get().strip()
    if not art:
        return
    tree.delete(*tree.get_children())
    sel = source_var.get()
    global rows_cache, last_part
    rows_cache = []

    # Datenbank
    if sel in ("Alle", "Datenbank"):
        sap = db_df[SAP_COL].astype(str).str.replace(r"\.0$","",regex=True).str.strip()
        hbn = db_df[HBN_COL].astype(str).str.strip().str.lower()
        rows_cache += db_df[(sap == art) | (hbn == art.lower())].to_dict("records")

    # Automotive‑Connectors
    if sel in ("Alle", "Automotive‑Connectors"):
        r = ac_price(art)
        if r:
            rows_cache.append(r)

    # Mouser
    if sel in ("Alle", "Mouser"):
        r = mouser_price(art)
        if r:
            rows_cache.append(r)

    last_part = art
    for r in rows_cache:
        tree.insert("", tk.END,
                    values=[fmt(r.get(c, "")) for c in VISIBLE_COLS])
    if not rows_cache:
        messagebox.showinfo("Info", "Keine Treffer gefunden.")

def export_tbl():
    if not rows_cache:
        messagebox.showwarning("Leere Tabelle", "Nichts zu exportieren.")
        return
    fname = filedialog.asksaveasfilename(
        defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")]
    )
    if not fname:
        return
    pd.DataFrame(rows_cache, columns=VISIBLE_COLS).to_excel(fname, index=False)
    wb = load_workbook(fname)
    wb.active.auto_filter.ref = f"A1:{chr(64+len(VISIBLE_COLS))}1"
    wb.save(fname)
    messagebox.showinfo("Export", "Datei gespeichert.")

# ───────── GUI Aufbau ─────────
root = tk.Tk(); root.withdraw()
db_file = filedialog.askopenfilename(
    title="Excel‑Datenbank wählen", filetypes=[("Excel","*.xls*")]
)
if not db_file:
    exit()
db_df = load_db(pathlib.Path(db_file))

root.deiconify()
root.title("Preis‑DB + Online‑Preise")
root.geometry("1450x730")

top = tk.Frame(root); top.pack(padx=10, pady=5, fill="x")
tk.Label(top, text="Artikelnummer:").pack(side="left")
entry = tk.Entry(top, width=36); entry.pack(side="left", padx=6)

source_var = tk.StringVar(value="Alle")
ttk.Combobox(
    top, textvariable=source_var,
    values=["Alle","Datenbank","Automotive‑Connectors","Mouser"],
    state="readonly", width=25
).pack(side="left", padx=6)

tk.Button(top, text="Suche",  width=10, command=do_search  ).pack(side="left",  padx=6)
tk.Button(top, text="Export", width=8,  command=export_tbl).pack(side="right", padx=6)

tree = ttk.Treeview(root, columns=VISIBLE_COLS, show="headings", height=28)
for c in VISIBLE_COLS:
    w = 150 if c not in ("Beschreibung","Datenbankbeschreibung \ndeutsch 2") else 260
    tree.column(c, width=w, anchor="center")
    tree.heading(c, text=c)
tree.pack(fill="both", expand=True, padx=10, pady=(4,0))

hbar = ttk.Scrollbar(root, orient="horizontal", command=tree.xview)
tree.configure(xscrollcommand=hbar.set)
hbar.pack(fill="x", padx=10)

root.mainloop()
