import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import threading
from datetime import datetime
from excel_search import load_excel, search_and_show, merge_results
from online_sources import get_online_results
from bom_tools import read_bom, detect_both_part_columns

try:
    import win32com.client
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

DB_JSON_FILE = "database.json"

def show_table(df, tree, veraltet_indices=None):
    # Spalten, die NICHT angezeigt werden sollen (aber im DataFrame bleiben!)
    hide_cols = ["ENTRY", "Description_deutsch_2"]
    if veraltet_indices is None:
        veraltet_indices = []
    if "Status" in df.columns:
        df = df.drop(columns=["Status"])
    # Nur für die Anzeige ausblenden:
    display_df = df[[col for col in df.columns if col not in hide_cols]]
    display_df = display_df.replace({None: '', 'nan': '', float('nan'): ''}).fillna('')
    display_df = display_df.loc[:, (display_df != '').any(axis=0)]

    tree["columns"] = ["Auswahl"] + list(display_df.columns)
    tree.heading("Auswahl", text="✓")
    tree.column("Auswahl", width=40, anchor="center")
    for col in display_df.columns:
        tree.heading(col, text=str(col))
        tree.column(col, width=120, anchor="center")
    tree.delete(*tree.get_children())
    tree.tag_configure("veraltet", background="#ffcccc")

    for idx, (_, row) in enumerate(display_df.iterrows()):
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

def update_excel_prices_win32com(excel_path, updates):
    if not HAS_WIN32:
        messagebox.showerror("Error", "win32com.client ist nicht installiert. Kann Excel nicht automatisieren.")
        return
    try:
        import win32com.client
        print(f"[DEBUG] Starte Excel-Update für Datei: {excel_path}")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        wb = excel.Workbooks.Open(os.path.abspath(excel_path))
        ws = wb.Worksheets("DB_4erDS")
        ws.Unprotect("cpq6ve")
        max_row = ws.UsedRange.Rows.Count
        print(f"[DEBUG] Maximale Zeile in Excel: {max_row}")
        print(f"[DEBUG] Anzahl Updates: {len(updates)}")

        for upd_idx, upd in enumerate(updates):
            suchnummer = str(upd.get("artikelnummer") or upd.get("1000ernummer")).strip().lower()
            price_block = upd['price_block']
            losgroesse = str(price_block[2]).strip()
            quelle = str(price_block[3]).strip().lower()
            gefunden = False

            print(f"[DEBUG] Update {upd_idx+1}/{len(updates)}: suchnummer={suchnummer}, losgroesse={losgroesse}, quelle={quelle}, price_block={price_block}")

            # Prüfe, ob alle 4 Felder gefüllt sind
            if not all(str(x).strip() not in ("", "nan", "None") for x in price_block):
                print(f"[DEBUG] Unvollständiger Block wird übersprungen: {price_block}")
                continue

            # Suche nach bestehendem Block mit gleicher Artikelnummer, Losgröße & Quelle
            for row in range(8, max_row + 1, 4):
                zelle_artikel = str(ws.Cells(row, 2).Value).strip().lower()
                zelle_1000er = str(ws.Cells(row, 3).Value).strip().lower()
                alt_losgroesse = str(ws.Cells(row + 2, 24).Value).strip()
                alt_quelle = str(ws.Cells(row + 3, 24).Value).strip().lower()
                print(f"[DEBUG]   Prüfe Zeile {row}: artikel={zelle_artikel}, 1000er={zelle_1000er}, alt_losgroesse={alt_losgroesse}, alt_quelle={alt_quelle}")

                if suchnummer in (zelle_artikel, zelle_1000er):
                    if normalize_losgroesse(alt_losgroesse) == normalize_losgroesse(losgroesse) and normalize(alt_quelle) == normalize(quelle):
                        for i, value in enumerate(price_block):
                            print(f"[DEBUG]     Schreibe Wert '{value}' in Zeile {row + i}, Spalte 24 (Überschreiben)")
                            ws.Cells(row + i, 24).Value = value
                        gefunden = True
                        print(f"[DEBUG]   -> Überschrieben! (row={row})")
                        break

            if gefunden:
                continue

            # Kein passender Block: Füge neuen Block am Tabellenende an
            letzte_zeile = None
            for row in range(8, max_row + 1, 4):
                zelle_artikel = str(ws.Cells(row, 2).Value).strip().lower()
                zelle_1000er = str(ws.Cells(row, 3).Value).strip().lower()
                if suchnummer in (zelle_artikel, zelle_1000er):
                    letzte_zeile = row
            if letzte_zeile is not None:
                neue_zeile = letzte_zeile + 4
            else:
                neue_zeile = max_row + 1
            print(f"[DEBUG]   Neuer Block für {suchnummer} ab Zeile {neue_zeile}")
            for i, value in enumerate(price_block):
                print(f"[DEBUG]     Schreibe Wert '{value}' in Zeile {neue_zeile + i}, Spalte 24 (Neuer Block)")
                ws.Cells(neue_zeile + i, 24).Value = value

        print("[DEBUG] Speichere Arbeitsmappe...")
        wb.Save()
        try:
            print("[DEBUG] Starte Makro 'NewPricesInDB'...")
            excel.Application.Run(f"'{wb.Name}'!NewPricesInDB")
            wb.Save()
            print("[DEBUG] Makro erfolgreich ausgeführt und gespeichert.")
        except Exception as e:
            print(f"[ERROR] Makro-Fehler: {e}")
            messagebox.showerror("Makro-Fehler", f"Makro konnte nicht ausgeführt werden:\n{e}")

        ws.Protect("cpq6ve", DrawingObjects=True, Contents=True, Scenarios=True, AllowFiltering=True)
        wb.Save()
        wb.Close()
        excel.Quit()
        print("[DEBUG] Excel geschlossen.")
    except Exception as e:
        print(f"[ERROR] Fehler beim Schreiben in Excel: {e}")
        messagebox.showerror("Excel-Fehler", f"Fehler beim Schreiben in Excel:\n{e}")

def start_app():
    root = tk.Tk()
    root.title("Preis-DB & Online-Preise")
    root.geometry("1600x900")

    # Forest-Theme laden (optional, falls vorhanden)
    try:
        style = ttk.Style()
        root.tk.call("source", os.path.join("themes", "forest-light.tcl"))
        style.theme_use("forest-light")
    except Exception as e:
        print(f"[DEBUG] Konnte Forest-Theme nicht laden: {e}")

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

    use_online_var = tk.BooleanVar(value=True)
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
        veraltet_indices = getattr(tree, "veraltet_indices", [])
        if anzeige_df is None or not veraltet_indices:
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
            if idx not in veraltet_indices:
                continue
            block_rows = anzeige_df.iloc[idx:idx+block_size]
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
                    print(f"[DEBUG] PREPARE UPDATE: artikelnummer={artikelnummer}, 1000ernummer={nummer_1000er}, price_block={price_block}, quelle={col}")
                    updates.append({
                        'artikelnummer': artikelnummer,
                        '1000ernummer': nummer_1000er,
                        'price_block': price_block,
                        'quelle': col
                    })
                else:
                    print(f"[DEBUG] Unvollständiger Block wird übersprungen (GUI): {price_block}")
        if not updates:
            messagebox.showinfo("No update", "Kein vollständiger und neuer Online-Block gefunden.")
            return

        print(f"[DEBUG] Anzahl Updates, die an Excel übergeben werden: {len(updates)}")
        update_excel_prices_win32com(excel_path, updates)
        messagebox.showinfo("Done", f"Excel-Datei wurde aktualisiert und gespeichert:\n{excel_path}")

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
            root.after(0, root.update_idletasks)

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