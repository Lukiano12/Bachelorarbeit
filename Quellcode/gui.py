import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import threading
from datetime import datetime, time
from excel_search import load_excel, search_and_show, merge_results
from online_sources import get_online_results
from bom_tools import read_bom, detect_both_part_columns
import win32com.client

try:
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

DB_JSON_FILE = "database.json"

def show_table(df, tree, veraltet_indices=None):
    hide_cols = ["ENTRY", "Description_deutsch_2"]  # Spalten, die ausgeblendet werden sollen
    df = df.drop(columns=[col for col in hide_cols if col in df.columns], errors="ignore")
    if veraltet_indices is None:
        veraltet_indices = []
    if "Status" in df.columns:
        df = df.drop(columns=["Status"])
    df = df.replace({None: '', 'nan': '', float('nan'): ''}).fillna('')
    df = df.loc[:, (df != '').any(axis=0)]

    tree["columns"] = ["Auswahl"] + list(df.columns)
    tree.heading("Auswahl", text="✓")
    tree.column("Auswahl", width=40, anchor="center")
    for col in df.columns:
        tree.heading(col, text=str(col))
        tree.column(col, width=120, anchor="center")
    tree.delete(*tree.get_children())
    tree.tag_configure("veraltet", background="#ffcccc")

    for idx, (_, row) in enumerate(df.iterrows()):
        values = [""]  # Auswahl immer leer setzen!
        for x in row:
            if isinstance(x, (pd.Timestamp, datetime)):
                values.append(x.strftime("%d.%m.%Y"))
            else:
                values.append(str(x) if x is not None else '')
        tags = ()
        if idx in veraltet_indices:
            tags = ("veraltet",)
        tree.insert("", "end", values=values, tags=tags)

def load_db_from_json(json_path=DB_JSON_FILE):
    if not os.path.exists(json_path):
        return None
    try:
        df = pd.read_json(json_path, orient="split")
        block_size = 4
        for col in df.columns:
            for i in range(0, len(df), block_size):
                val = df.iloc[i][col]
                if isinstance(val, (int, float)) and val > 1e12:
                    try:
                        df.iloc[i, df.columns.get_loc(col)] = pd.to_datetime(val, unit='ms')
                    except Exception:
                        pass
        return df
    except Exception as e:
        print(f"Fehler beim Laden von JSON: {e}")
        return None

def save_db_to_json(df, json_path=DB_JSON_FILE):
    try:
        df.to_json(json_path, orient="split", force_ascii=False)
        print(f"Datenbank gespeichert unter: {os.path.abspath(json_path)}")
    except Exception as e:
        print(f"Fehler beim Speichern zu JSON: {e}")

def clean_price(value):
    if isinstance(value, str):
        value = value.replace("€", "").replace(" ", "").replace(",", ".")
    try:
        return float(value)
    except Exception:
        return value

def is_online_source(colname):
    online_keywords = ["mouser", "octopart", "digi-key", "arrow", "online", "connector"]
    return any(key in colname.lower() for key in online_keywords)

def normalize(val):
    if val is None:
        return ""
    return str(val).strip().lower()

def normalize_losgroesse(val):
    try:
        return str(int(float(str(val).replace(",", ".").strip())))
    except Exception:
        return str(val).strip().lower()

def normalize_quelle(val):
    # Nimmt nur den Namen vor der ersten Klammer und wandelt in Kleinbuchstaben um
    v = str(val).strip().lower()
    v = v.split("(")[0].strip()
    return v

def normalize_nummer_1000er(val):
    try:
        return str(int(float(str(val).replace(",", ".").strip())))
    except Exception:
        return str(val).strip().lower()

def build_excel_index(ws, max_row):
    """
    Erzeugt ein Dictionary, das (artikelnummer, nummer_1000er, losgroesse, quelle) auf die Startzeile des Blocks abbildet.
    """
    index = {}
    for row in range(8, max_row + 1, 4):
        artikelnummer = str(ws.Cells(row, 2).Value).strip().lower()
        nummer_1000er = normalize_nummer_1000er(ws.Cells(row, 3).Value)
        losgroesse = normalize_losgroesse(ws.Cells(row + 2, 24).Value)
        quelle = normalize_quelle(ws.Cells(row + 3, 24).Value)
        key = (artikelnummer, nummer_1000er, losgroesse, quelle)
        index[key] = row
    return index

def update_excel_prices_win32com(excel_path, updates, progress_var=None, status_label=None, root=None):
    if not HAS_WIN32:
        if root:
            root.after(0, messagebox.showerror, "Error", "win32com.client ist nicht installiert. Kann Excel nicht automatisieren.")
        else:
            messagebox.showerror("Error", "win32com.client ist nicht installiert. Kann Excel nicht automatisieren.")
        return
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        wb = excel.Workbooks.Open(os.path.abspath(excel_path))
        ws = wb.Worksheets("DB_4erDS")
        ws.Unprotect("cpq6ve")
        max_row = ws.UsedRange.Rows.Count

        # Index aufbauen
        excel_index = build_excel_index(ws, max_row)

        total = len(updates)
        for idx, upd in enumerate(updates):
            artikelnummer = str(upd.get("artikelnummer", "")).strip().lower()
            nummer_1000er = normalize_nummer_1000er(upd.get("1000ernummer", ""))
            price_block = upd['price_block']
            losgroesse = normalize_losgroesse(price_block[2])
            quelle = normalize_quelle(price_block[3])
            key = (artikelnummer, nummer_1000er, losgroesse, quelle)
            print(f"[DEBUG] Suche Key: {key}")
            print(f"[DEBUG] Excel-Keys: {list(excel_index.keys())}")

            row = excel_index.get(key)
            def set_cell_value(cell, value):
                # Versuche, Wert als float zu setzen, falls möglich
                if isinstance(value, str):
                    v = value.strip()
                    # Versuche, als Datum zu setzen
                    try:
                        from dateutil.parser import parse
                        dt = parse(v, dayfirst=True, fuzzy=True)
                        if isinstance(dt, datetime):
                            cell.Value = datetime(dt.year, dt.month, dt.day, 12, 0, 0)
                        else:
                            cell.Value = datetime(dt.year, dt.month, dt.day, 12, 0, 0)
                            cell.NumberFormat = "DD.MM.YYYY"
                        return
                    except Exception:
                        pass
                    # Versuche, als float zu setzen
                    try:
                        cell.Value = float(v.replace(",", "."))
                        return
                    except Exception:
                        pass
                    # Sonst als String
                    cell.Value = v
                elif isinstance(value, (int, float)):
                    cell.Value = value
                elif value is None:
                    cell.Value = ""
                else:
                    cell.Value = str(value).strip()

            if row:
                for i, value in enumerate(price_block):
                    ws.Cells(row + i, 24).Value = str(value).strip()
                # Nach jedem Eintrag Makro ausführen und speichern, damit Excel den Block verschiebt!
                try:
                    excel.Application.Run(f"'{wb.Name}'!NewPricesInDB")
                except Exception as e:
                    print(f"[WARN] Makro konnte nicht ausgeführt werden: {e}")
                wb.Save()
            else:
                # Falls kein passender Block existiert, neuen Block suchen/anhängen
                for row_candidate in range(8, max_row + 1, 4):
                    zelle_artikel = str(ws.Cells(row_candidate, 2).Value).strip().lower()
                    zelle_1000er = normalize_nummer_1000er(ws.Cells(row_candidate, 3).Value)
                    if artikelnummer in (zelle_artikel, zelle_1000er):
                        empty = True
                        for i in range(4):
                            val = ws.Cells(row_candidate + i, 24).Value
                            if val is not None and str(val).strip().lower() not in ("", "none", "nan"):
                                empty = False
                                break
                        if empty:
                            for i, value in enumerate(price_block):
                                ws.Cells(row_candidate + i, 24).Value = str(value).strip()
                            try:
                                excel.Application.Run(f"'{wb.Name}'!NewPricesInDB")
                            except Exception as e:
                                print(f"[WARN] Makro konnte nicht ausgeführt werden: {e}")
                            wb.Save()
                            break

            if progress_var and status_label and root:
                def update_gui(idx=idx):
                    progress = (idx + 1) / total * 100
                    progress_var.set(progress)
                    status_label.config(text=f"Aktualisiere Preise: {idx + 1} / {total}")
                    root.update_idletasks()
                root.after(0, update_gui)

        ws.Protect("cpq6ve", DrawingObjects=True, Contents=True, Scenarios=True, AllowFiltering=True)
        wb.Save()
        wb.Close()
        excel.Quit()
        print("[DEBUG] Excel geschlossen.")
        if root and status_label:
            root.after(0, status_label.config, {"text": "Preis-Update abgeschlossen."})
    except Exception as e:
        print(f"[ERROR] Fehler beim Schreiben in Excel: {e}")
        if root:
            root.after(0, messagebox.showerror, "Excel-Fehler", f"Fehler beim Schreiben in Excel:\n{e}")
        else:
            messagebox.showerror("Excel-Fehler", f"Fehler beim Schreiben in Excel:\n{e}")

def start_app():
    root = tk.Tk()
    root.title("Preis-DB & Online-Preise")
    root.geometry("1600x900")

    # Forest-Theme laden
    style = ttk.Style()
    root.tk.call("source", os.path.join("themes", "forest-light.tcl"))
    style.theme_use("forest-light")

    df = None
    search_cols = ["WN_SAP-Artikel-NR", "WN_HerstellerBestellnummer_1"]

    frame = ttk.Frame(root, padding=8)
    frame.pack(fill="x", padx=10, pady=4)
    label = ttk.Label(frame, text="Artikelnummer oder SAP-Nummer:")
    label.pack(side="left")
    entry = ttk.Entry(frame, width=30)
    entry.pack(side="left", padx=5)
    search_btn = ttk.Button(frame, text="Suche", width=16, state="disabled")
    search_btn.pack(side="left", padx=5)
    bom_btn = ttk.Button(frame, text="BOM laden & Suchen", width=20, state="disabled")
    bom_btn.pack(side="left", padx=5)
    export_btn = ttk.Button(frame, text="Export als Excel", width=18)
    export_btn.pack(side="left", padx=5)
    update_db_btn = ttk.Button(frame, text="Datenbank aktualisieren", width=22)
    update_db_btn.pack(side="left", padx=5)
    update_excel_btn = ttk.Button(frame, text="Excel-Preise aktualisieren", width=22)
    update_excel_btn.pack(side="left", padx=5)

    use_online_var = tk.BooleanVar(value=False)
    online_check = ttk.Checkbutton(frame, text="Online Quellen nutzen", variable=use_online_var, onvalue=True, offvalue=False)
    online_check.pack(side="left", padx=10)

    result_frame = ttk.Frame(root, padding=4)
    result_frame.pack(fill="both", expand=True, padx=10, pady=4)

    tree_scroll_y = ttk.Scrollbar(result_frame, orient="vertical")
    tree_scroll_x = ttk.Scrollbar(result_frame, orient="horizontal")
    tree = ttk.Treeview(
        result_frame, columns=[], show="headings",
        yscrollcommand=tree_scroll_y.set,
        xscrollcommand=tree_scroll_x.set,
        selectmode="extended"
    )
    tree_scroll_y.config(command=tree.yview)
    tree_scroll_y.pack(side="right", fill="y")
    tree_scroll_x.config(command=tree.xview)
    tree_scroll_x.pack(side="bottom", fill="x")
    tree.pack(fill="both", expand=True)

    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(root, variable=progress_var, mode="determinate")
    progress_bar.pack(fill="x", padx=10, pady=2)
    status_label = ttk.Label(root, text="Bereit")
    status_label.pack(fill="x", padx=10, pady=(0, 5))

    def on_tree_click(event):
        item = tree.identify_row(event.y)
        col = tree.identify_column(event.x)
        if not item or col != "#1":
            return
        idx = int(tree.index(item))
        block_size = 4
        if idx % block_size == 0:
            current = tree.set(item, "Auswahl")
            tree.set(item, "Auswahl", "✓" if current != "✓" else "")

    tree.bind("<Button-1>", on_tree_click)

    def initialize_db():
        nonlocal df
        df_loaded = load_db_from_json(DB_JSON_FILE)
        if df_loaded is not None and all(col in df_loaded.columns for col in search_cols):
            df = df_loaded
            search_btn.config(state="normal")
            bom_btn.config(state="normal")
        else:
            messagebox.showinfo(
                "Keine Datenbank",
                "Es wurde keine Datenbank gefunden. Bitte laden Sie eine Excel-Datei zur Initialisierung."
            )
            update_db_from_excel()

    def update_db_from_excel():
        nonlocal df
        file = filedialog.askopenfilename(
            title="Excel-Datei wählen", filetypes=[("Excel-Dateien", "*.xls* *.xlsx *.xlsm *.xlsb")]
        )
        if not file:
            return
        try:
            df_loaded = load_excel(file)
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Laden der Datei:\n{str(e)}")
            return
        missing_cols = [col for col in search_cols if col not in df_loaded.columns]
        if missing_cols:
            messagebox.showerror(
                "Fehler", f"Folgende Spalten fehlen in der Tabelle: {missing_cols}"
            )
            return
        df = df_loaded
        save_db_to_json(df, DB_JSON_FILE)
        df = load_db_from_json(DB_JSON_FILE)
        search_btn.config(state="normal")
        bom_btn.config(state="normal")
        messagebox.showinfo("Erfolg", "Datenbank wurde aktualisiert und gespeichert.")

    def do_search(event=None):
        if df is None:
            messagebox.showwarning("Keine Datenbank", "Bitte laden Sie zuerst eine Datenbank-Datei.")
            return
        search = entry.get().strip()
        if not search:
            return
        db_rows = search_and_show(df, search, search_cols)
        artikelnummer = db_rows.iloc[0]['WN_HerstellerBestellnummer_1'] if (db_rows is not None and 'WN_HerstellerBestellnummer_1' in db_rows.columns) else search
        if use_online_var.get():
            online_results_list = get_online_results(artikelnummer)
        else:
            online_results_list = []
        merged, veraltet_indices = merge_results(db_rows, online_results_list)
        if merged is not None and not merged.empty:
            merged = merged.replace({None: '', 'nan': '', float('nan'): ''}).fillna('')
            show_table(merged, tree, veraltet_indices)
            tree.gesamt_df = merged
            tree.veraltet_indices = veraltet_indices
            tree.anzeige_df = merged
        else:
            messagebox.showinfo(
                "Kein Treffer", f"Keine Zeile mit '{search}' gefunden (weder DB noch online)."
            )

    def update_selected_prices_in_excel():
        selected_items = []
        block_size = 4
        for item in tree.get_children():
            idx = int(tree.index(item))
            if idx % block_size == 0 and tree.set(item, "Auswahl") == "✓":
                selected_items.append(item)
        if not selected_items:
            messagebox.showinfo("No selection", "Bitte mindestens einen Block per Checkbox auswählen.")
            return

        anzeige_df = getattr(tree, "anzeige_df", None)
        if anzeige_df is None:
            messagebox.showinfo("No data", "No outdated blocks to update.")
            return

        excel_path = filedialog.askopenfilename(
            title="Excel-Datei wählen",
            filetypes=[
                ("Makro-fähige Excel-Dateien", "*.xlsm"),
                ("Alle Excel-Dateien", "*.xlsx *.xls *.xlsm *.xlsb"),
                ("Alle Dateien", "*.*")
            ]
        )
        if not excel_path:
            return

        updates = []
        for item in selected_items:
            idx = int(tree.index(item))
            block_rows = anzeige_df.iloc[idx:idx+block_size]
            found_any = False
            for col in anzeige_df.columns:
                if not is_online_source(col):
                    continue
                price_block = [
                    block_rows.iloc[0][col],
                    clean_price(block_rows.iloc[1][col]),
                    block_rows.iloc[2][col],
                    block_rows.iloc[3][col],
                ]
                if all(str(x).strip() not in ("", "nan", "None") for x in price_block):
                    artikelnummer = str(block_rows.iloc[0][anzeige_df.columns[0]])
                    nummer_1000er = str(block_rows.iloc[0][anzeige_df.columns[1]]) if len(anzeige_df.columns) > 1 else ""
                    updates.append({
                        'artikelnummer': artikelnummer,
                        '1000ernummer': nummer_1000er,
                        'price_block': price_block,
                        'quelle': col
                    })
                    found_any = True
            if not found_any:
                print(f"[WARN] Kein vollständiger Online-Block für Index {idx} gefunden.")

        if not updates:
            messagebox.showinfo("No update", "Kein vollständiger und neuer Online-Block gefunden.")
            return

        print(f"[DEBUG] Anzahl Updates, die an Excel übergeben werden: {len(updates)}")

        progress_var.set(0)
        progress_bar.update()
        status_label.config(text="Starte Preis-Update ...")
        root.update_idletasks()

        def worker():
            update_excel_prices_win32com(excel_path, updates, progress_var, status_label, root)
            root.after(0, messagebox.showinfo, "Done", f"Excel-Datei wurde aktualisiert und gespeichert:\n{excel_path}")

        threading.Thread(target=worker, daemon=True).start()

    def load_bom_and_search():
        def worker():
            if df is None:
                messagebox.showwarning("Keine Datenbank", "Bitte laden Sie zuerst eine Datenbank-Datei.")
                return
            bomfile = filedialog.askopenfilename(
                title="BOM-Datei wählen",
                filetypes=[("Excel/CSV", "*.xls*;*.csv")]
            )
            if not bomfile:
                return
            try:
                bom_df = read_bom(bomfile)
                sap_col, art_col = detect_both_part_columns(bom_df)
            except Exception as e:
                messagebox.showerror("BOM-Fehler", str(e))
                return

            bauteile = bom_df[art_col].dropna().unique()
            bauteile = [teil for teil in bauteile if str(teil).strip().upper() != "SPLICE"]
            gesamt_ergebnisse = []
            gesamt_veraltet = []

            total = len(bauteile)
            def update_progress(idx):
                progress = (idx + 1) / total * 100
                progress_var.set(progress)
                status_label.config(text=f"Lade BOM: {idx + 1}/ {total} Teile werden verarbeitet ...")
                root.update_idletasks()

            progress_var.set(0)
            progress_bar.update()
            status_label.config(text=f"Lade BOM: 0/ {total} Teile werden verarbeitet ...")
            root.update_idletasks()

            for idx, suchwert in enumerate(bauteile):
                suchwert = str(suchwert).strip()
                if not suchwert:
                    continue
                db_rows = search_and_show(df, suchwert, search_cols)
                artikelnummer = db_rows.iloc[0]['WN_HerstellerBestellnummer_1'] if (db_rows is not None and 'WN_HerstellerBestellnummer_1' in db_rows.columns) else suchwert
                if use_online_var.get():
                    online_results_list = get_online_results(artikelnummer)
                else:
                    online_results_list = []
                merged, veraltet_indices = merge_results(db_rows, online_results_list)
                if merged is not None and not merged.empty:
                    merged = merged.replace({None: '', 'nan': '', float('nan'): ''}).fillna('')
                    offset = sum(len(df) for df in gesamt_ergebnisse)
                    gesamt_ergebnisse.append(merged)
                    gesamt_veraltet.extend([i + offset for i in veraltet_indices])
                root.after(0, update_progress, idx)

            root.after(0, progress_var.set, 100)
            root.after(0, status_label.config, {"text": "BOM-Laden abgeschlossen."})
            root.after(0, progress_bar.update)
            root.after(0, root.update_idletasks()

            )

            if not gesamt_ergebnisse:
                root.after(0, messagebox.showinfo, "Info", "Keine Ergebnisse für die BOM-Bauteile.")
                return

            gesamt_df = pd.concat(gesamt_ergebnisse, ignore_index=True)
            gesamt_df = gesamt_df.replace({None: '', 'nan': '', float('nan'): ''}).fillna('')
            db_spalten = [col for col in gesamt_df.columns if not ("mouser" in col.lower() or "octopart" in col.lower())]
            online_spalten = [col for col in gesamt_df.columns if ("mouser" in col.lower() or "octopart" in col.lower())]
            gesamt_df = gesamt_df[db_spalten + online_spalten]
            root.after(0, show_table, gesamt_df, tree, gesamt_veraltet)
            tree.gesamt_df = gesamt_df
            tree.veraltet_indices = gesamt_veraltet
            tree.anzeige_df = gesamt_df

        threading.Thread(target=worker, daemon=True).start()

    def export_as_excel():
        df_to_export = getattr(tree, "gesamt_df", None)
        if df_to_export is None or df_to_export.empty:
            messagebox.showwarning("Export", "Keine Daten zum Exportieren!")
            return
        df_to_export = df_to_export.replace({None: '', 'nan': '', float('nan'): ''}).fillna('')
        fname = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            title="Datei speichern als"
        )
        if not fname:
            return
        df_to_export.to_excel(fname, index=False)
        messagebox.showinfo("Export", f"Erfolgreich gespeichert:\n{fname}")

    update_db_btn.config(command=update_db_from_excel)
    search_btn.config(command=do_search)
    entry.bind("<Return>", do_search)
    bom_btn.config(command=load_bom_and_search)
    export_btn.config(command=export_as_excel)
    update_excel_btn.config(command=update_selected_prices_in_excel)

    initialize_db()
    root.mainloop()