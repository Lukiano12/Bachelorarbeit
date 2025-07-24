import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import threading
from excel_search import load_excel, search_and_show, merge_results
from online_sources import get_online_results
from bom_tools import read_bom, detect_both_part_columns

DB_JSON_FILE = "database.json"   # Your default DB filename

def show_table(df, tree):
    df = df.replace({None: '', 'nan': '', float('nan'): ''}).fillna('')
    df = df.loc[:, (df != '').any(axis=0)]
    for col in tree["columns"]:
        tree.heading(col, text="")
        tree.column(col, width=0)
    tree.delete(*tree.get_children())
    tree["columns"] = list(df.columns)
    for col in df.columns:
        tree.heading(col, text=str(col))
        tree.column(col, width=120, anchor="center")
    for _, row in df.iterrows():
        values = [str(x) if x is not None else '' for x in row]
        tree.insert("", tk.END, values=values)

def load_db_from_json(json_path):
    if not os.path.exists(json_path):
        return None
    try:
        return pd.read_json(json_path, orient="split")
    except Exception as e:
        print(f"Fehler beim Laden von JSON: {e}")
        return None

def save_db_to_json(df, json_path):
    try:
        df.to_json(json_path, orient="split")
    except Exception as e:
        print(f"Fehler beim Speichern zu JSON: {e}")

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

    use_online_var = tk.BooleanVar(value=True)
    online_check = tk.Checkbutton(frame, text="Online Quellen nutzen",variable=use_online_var )
    online_check.pack(side="left", padx=10)

    result_frame = tk.Frame(root)
    result_frame.pack(fill="both", expand=True, padx=10, pady=4)

    # --- SCROLLBARS for Treeview ---
    tree_scroll_y = tk.Scrollbar(result_frame, orient="vertical")
    tree_scroll_x = tk.Scrollbar(result_frame, orient="horizontal")
    tree = ttk.Treeview(
        result_frame, columns=[], show="headings",
        yscrollcommand=tree_scroll_y.set,
        xscrollcommand=tree_scroll_x.set
    )
    tree_scroll_y.config(command=tree.yview)
    tree_scroll_y.pack(side="right", fill="y")
    tree_scroll_x.config(command=tree.xview)
    tree_scroll_x.pack(side="bottom", fill="x")
    tree.pack(fill="both", expand=True)

    # Fortschrittsbalken und Status-Label
    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
    progress_bar.pack(fill="x", padx=10, pady=2)
    status_label = tk.Label(root, text="Bereit")
    status_label.pack(fill="x", padx=10, pady=(0, 5))

    def initialize_db():
        nonlocal df
        # Try to load from JSON first
        df_loaded = load_db_from_json(DB_JSON_FILE)
        if df_loaded is not None and all(col in df_loaded.columns for col in search_cols):
            df = df_loaded
            search_btn.config(state="normal")
            bom_btn.config(state="normal")
        else:
            # Prompt user for initial Excel import if DB not found
            messagebox.showinfo(
                "Keine Datenbank",
                "Es wurde keine Datenbank gefunden. Bitte laden Sie eine Excel-Datei zur Initialisierung."
            )
            update_db_from_excel()

    def update_db_from_excel():
        nonlocal df
        file = filedialog.askopenfilename(
            title="Excel-Datei wählen", filetypes=[("Excel-Dateien", "*.xls*")]
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
        # Nur wenn Checkbox aktiviert ist, Online-Quellen abfragen
        if use_online_var.get():
            online_results_list = get_online_results(artikelnummer)
        else:
            online_results_list = []
        merged = merge_results(db_rows, online_results_list)
        if merged is not None and not merged.empty:
            show_table(merged, tree)
            tree.gesamt_df = merged  # Für Export
        else:
            messagebox.showinfo(
                "Kein Treffer", f"Keine Zeile mit '{search}' gefunden (weder DB noch online)."
            )

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
                merged = merge_results(db_rows, online_results_list)
                if merged is not None and not merged.empty:
                    gesamt_ergebnisse.append(merged)
                # Fortschritt im Hauptthread aktualisieren
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
            root.after(0, show_table, gesamt_df, tree)
            tree.gesamt_df = gesamt_df  # Für Export

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

    # Initialize database from JSON (or Excel if missing)
    initialize_db()

    root.mainloop()
