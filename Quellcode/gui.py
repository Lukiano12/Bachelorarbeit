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

DB_JSON_FILE = "database.json"   # Your default DB filename

def show_table(df, tree, veraltet_indices=None):
    print("Veraltete Indices:", veraltet_indices)
    if veraltet_indices is None:
        veraltet_indices = []
    if "Status" in df.columns:
        df = df.drop(columns=["Status"])
    df = df.replace({None: '', 'nan': '', float('nan'): ''}).fillna('')
    df = df.loc[:, (df != '').any(axis=0)]

    # --- Duplikate pro Spalte/Block entfernen ---
    block_size = 4
    unique_blocks = []
    seen_blocks = set()
    for i in range(0, len(df), block_size):
        block = df.iloc[i:i+block_size]
        # Für jede Spalte prüfen
        for col in df.columns:
            block_tuple = tuple(block[col].values)
            key = (col, block_tuple)
            if key not in seen_blocks:
                seen_blocks.add(key)
                # Block in neues DataFrame übernehmen
                if len(unique_blocks) <= i:
                    unique_blocks.extend([{} for _ in range(i + block_size - len(unique_blocks))])
                for j in range(block_size):
                    unique_blocks[i + j][col] = block.iloc[j][col]
    # Neues DataFrame aus den eindeutigen Blöcken bauen
    if unique_blocks:
        df = pd.DataFrame(unique_blocks)

    # Treeview vorbereiten: Checkbox-Spalte + Daten
    tree["columns"] = ["Auswahl"] + list(df.columns)
    tree.heading("Auswahl", text="✓")
    tree.column("Auswahl", width=40, anchor="center")
    for col in df.columns:
        tree.heading(col, text=str(col))
        tree.column(col, width=120, anchor="center")
    tree.delete(*tree.get_children())
    tree.tag_configure("veraltet", background="#ffcccc")

    block_size = 4
    for idx, (_, row) in enumerate(df.iterrows()):
        # Nur in der ersten Zeile eines Blocks Checkbox anzeigen
        if idx % block_size == 0:
            values = [""]
        else:
            values = [""]  # Leere Checkbox-Spalte für die anderen Zeilen
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
    # Entfernt €-Zeichen, Leerzeichen und wandelt Komma in Punkt für float
    if isinstance(value, str):
        value = value.replace("€", "").replace(" ", "").replace(",", ".")
    try:
        return float(value)
    except Exception:
        return value  # Falls es kein Preis ist (z.B. Text), bleibt es wie es ist

def block_exists_in_db(df, artikelnummer, nummer_1000er, price_block):
    # Suche alle Blöcke für diesen Artikel in der Datenbank
    block_size = 4
    for i in range(0, len(df), block_size):
        row_artikel = str(df.iloc[i][df.columns[0]])
        row_1000er = str(df.iloc[i][df.columns[1]]) if len(df.columns) > 1 else ""
        if artikelnummer == row_artikel or nummer_1000er == row_1000er:
            for col in df.columns:
                alt_block = [
                    str(df.iloc[i][col]).strip(),
                    str(df.iloc[i+1][col]).strip(),
                    str(df.iloc[i+2][col]).strip(),
                    str(df.iloc[i+3][col]).strip()
                ]
                # Preis im Altblock auch bereinigen
                try:
                    alt_block[1] = str(clean_price(alt_block[1]))
                except Exception:
                    pass
                # Preis im neuen Block als String für Vergleich
                new_block = [
                    str(price_block[0]).strip(),
                    str(price_block[1]).strip(),
                    str(price_block[2]).strip(),
                    str(price_block[3]).strip()
                ]
                if alt_block == new_block:
                    return True
    return False

def is_online_source(colname):
    # Passe die Liste an deine Online-Quellen an!
    online_keywords = ["mouser", "octopart", "digi-key", "arrow", "online", "connector"]
    return any(key in colname.lower() for key in online_keywords)

def update_excel_prices_win32com(excel_path, updates):
    if not HAS_WIN32:
        messagebox.showerror("Error", "win32com.client ist nicht installiert. Kann Excel nicht automatisieren.")
        return
    try:
        import win32com.client
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        wb = excel.Workbooks.Open(os.path.abspath(excel_path))
        ws = wb.Worksheets("DB_4erDS")
        ws.Unprotect("cpq6ve")
        max_row = ws.UsedRange.Rows.Count

        max_row = ws.UsedRange.Rows.Count

        for upd in updates:
            suchnummer = str(upd.get("artikelnummer") or upd.get("1000ernummer"))
            gefunden = False
            # Suche nach bestehendem Block mit gleicher Losgröße & Quelle
            for row in range(8, max_row + 1, 4):
                zelle_artikel = str(ws.Cells(row, 2).Value)
                zelle_1000er = str(ws.Cells(row, 3).Value)
                if suchnummer in (zelle_artikel, zelle_1000er):
                    alt_losgroesse = str(ws.Cells(row + 2, 24).Value).strip()
                    alt_quelle = str(ws.Cells(row + 3, 24).Value).strip()
                    neu_losgroesse = str(upd['price_block'][2]).strip()
                    neu_quelle = str(upd['price_block'][3]).strip()
                    if alt_losgroesse == neu_losgroesse and alt_quelle == neu_quelle:
                        # Block gefunden: Überschreiben!
                        for i, value in enumerate(upd['price_block']):
                            ws.Cells(row + i, 24).Value = value
                            print(f"[DEBUG] Aktualisiere Wert '{value}' in Zeile {row + i}, Spalte 24 (X)")
                        gefunden = True
                        break
            if not gefunden:
                # Kein passender Block: Füge neuen Block am Ende an
                for row in range(8, max_row + 1, 4):
                    zelle_artikel = str(ws.Cells(row, 2).Value)
                    zelle_1000er = str(ws.Cells(row, 3).Value)
                    if suchnummer in (zelle_artikel, zelle_1000er):
                        if all(str(ws.Cells(row + i, 24).Value).strip() in ("", "None", "nan") for i in range(4)):
                            for i, value in enumerate(upd['price_block']):
                                ws.Cells(row + i, 24).Value = value
                                print(f"[DEBUG] Schreibe neuen Wert '{value}' in Zeile {row + i}, Spalte 24 (X)")
                            gefunden = True
                            break
                # Falls kein leerer Block gefunden, ggf. am Ende anfügen (optional)

            wb.Save()
            # Makro nach jedem Block ausführen (auf geöffneter Datei)
            try:
                print("[DEBUG] Starte Makro...")
                excel.Application.Run(f"'{wb.Name}'!NewPricesInDB")
                wb.Save()
            except Exception as e:
                print(f"[ERROR] Fehler beim Makro: {e}")
                messagebox.showerror("Makro-Fehler", f"Makro konnte nicht ausgeführt werden:\n{e}")

        ws.Protect("cpq6ve", DrawingObjects=True, Contents=True, Scenarios=True, AllowFiltering=True)
        wb.Save()
        wb.Close()
        excel.Quit()
        print("[DEBUG] Excel geschlossen.")
    except Exception as e:
        print(f"[ERROR] Fehler beim Schreiben in Excel: {e}")
        messagebox.showerror("Excel-Fehler", f"Fehler beim Schreiben in Excel:\n{e}")

def run_excel_macro(excel_path, macro_name):
    if not HAS_WIN32:
        messagebox.showerror("Error", "win32com.client ist nicht installiert. Kann Makros nicht ausführen.")
        return
    try:
        print(f"[DEBUG] Starte Excel für Makro: {excel_path}")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True  # Sichtbar machen für Debug!
        wb = excel.Workbooks.Open(os.path.abspath(excel_path))
        print(f"[DEBUG] Führe Makro aus: {macro_name}")
        excel.Application.Run(f"'{wb.Name}'!{macro_name}")
        wb.Save()
        print("[DEBUG] Nach Makro: Datei gespeichert.")
        wb.Close()
        excel.Quit()
        print("[DEBUG] Excel geschlossen nach Makro.")
    except Exception as e:
        print(f"[ERROR] Fehler beim Makro: {e}")
        messagebox.showerror("Makro-Fehler", f"Excel konnte die Datei nicht öffnen oder das Makro nicht ausführen.\n\n{e}")

def start_app():
    root = tk.Tk()
    root.title("Preis-DB & Online-Preise")
    root.geometry("1600x900")

    df = None
    search_cols = ["WN_SAP-Artikel-NR", "WN_HerstellerBestellnummer_1"]

    # --- GUI Elements ---
    frame = tk.Frame(root)
    frame.pack(fill="x", padx=10, pady=4)
    tk.Label(frame, text="Artikelnummer oder SAP-Nummer:").pack(side="left")
    entry = tk.Entry(frame, width=35)
    entry.pack(side="left", padx=5)
    search_btn = tk.Button(frame, text="Suche", width=12, state="disabled")
    search_btn.pack(side="left", padx=5)
    bom_btn = tk.Button(frame, text="BOM laden & Suchen", width=18, state="disabled")
    bom_btn.pack(side="left", padx=5)
    export_btn = tk.Button(frame, text="Export als Excel", width=15)
    export_btn.pack(side="left", padx=5)
    update_db_btn = tk.Button(frame, text="Datenbank aktualisieren", width=18)
    update_db_btn.pack(side="left", padx=5)
    update_excel_btn = tk.Button(frame, text="Update Excel Prices", width=18)
    update_excel_btn.pack(side="left", padx=5)

    use_online_var = tk.BooleanVar(value=True)
    online_check = tk.Checkbutton(frame, text="Online Quellen nutzen",variable=use_online_var )
    online_check.pack(side="left", padx=10)

    result_frame = tk.Frame(root)
    result_frame.pack(fill="both", expand=True, padx=10, pady=4)

    tree_scroll_y = tk.Scrollbar(result_frame, orient="vertical")
    tree_scroll_x = tk.Scrollbar(result_frame, orient="horizontal")
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
    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
    progress_bar.pack(fill="x", padx=10, pady=2)
    status_label = tk.Label(root, text="Bereit")
    status_label.pack(fill="x", padx=10, pady=(0, 5))

    # --- Checkbox-Umschaltung im Treeview ---
    def on_tree_click(event):
        item = tree.identify_row(event.y)
        col = tree.identify_column(event.x)
        if not item or col != "#1":
            return
        idx = int(tree.index(item))
        block_size = 4
        # Nur Checkbox in der ersten Zeile eines Blocks toggeln
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
            show_table(merged, tree, veraltet_indices)
            tree.gesamt_df = merged  # Für Export
            tree.veraltet_indices = veraltet_indices
            tree.anzeige_df = merged
        else:
            messagebox.showinfo(
                "Kein Treffer", f"Keine Zeile mit '{search}' gefunden (weder DB noch online)."
            )

    def update_selected_prices_in_excel():
        # Nur Checkbox in der ersten Zeile jedes Blocks zählt!
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

        # Excel-Datei auswählen (wird direkt bearbeitet!)
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
            if idx * block_size not in veraltet_indices:
                continue
            entry_row = idx * block_size
            block_rows = anzeige_df.iloc[entry_row:entry_row+block_size]
            for col in anzeige_df.columns:
                if not is_online_source(col):
                    continue  # Nur Online-Quellen prüfen/überschreiben!
                price_block = [
                    block_rows.iloc[0][col],
                    clean_price(block_rows.iloc[1][col]),
                    block_rows.iloc[2][col],
                    block_rows.iloc[3][col],
                ]
                if all(str(x).strip() not in ("", "nan", "None") for x in price_block):
                    artikelnummer = str(block_rows.iloc[0][anzeige_df.columns[0]])
                    nummer_1000er = str(block_rows.iloc[0][anzeige_df.columns[1]]) if len(anzeige_df.columns) > 1 else ""
                    # Prüfe, ob Block schon existiert:
                    if not block_exists_in_db(df, artikelnummer, nummer_1000er, price_block):
                        updates.append({
                            'artikelnummer': artikelnummer,
                            '1000ernummer': nummer_1000er,
                            'price_block': price_block,
                            'quelle': col
                        })
                        print(f"[DEBUG] Update: artikelnummer={artikelnummer}, 1000ernummer={nummer_1000er}, quelle={col}, price_block={price_block}")

        if not updates:
            messagebox.showinfo("No update", "Kein vollständiger und neuer Online-Block gefunden.")
            return

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
            db_spalten = [col for col in gesamt_df.columns if not ("mouser" in col.lower() or "octopart" in col.lower())]
            online_spalten = [col for col in gesamt_df.columns if ("mouser" in col.lower() or "octopart" in col.lower())]
            gesamt_df = gesamt_df[db_spalten + online_spalten]
            root.after(0, show_table, gesamt_df, tree, gesamt_veraltet)
            tree.gesamt_df = gesamt_df  # Für Export
            tree.veraltet_indices = gesamt_veraltet
            tree.anzeige_df = gesamt_df

        threading.Thread(target=worker, daemon=True).start()

    def export_as_excel():
        df_to_export = getattr(tree, "gesamt_df", None)
        if df_to_export is None or df_to_export.empty:
            messagebox.showwarning("Export", "Keine Daten zum Exportieren!")
            return
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