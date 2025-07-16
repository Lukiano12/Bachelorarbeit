import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from excel_search import load_excel, search_and_show, merge_results
from online_sources import get_online_results
from bom_tools import read_bom, detect_both_part_columns

def show_table(df, tree):
    df = df.replace({None: '', 'nan': '', float('nan'): ''}).fillna('')
    df = df.loc[:, (df != '').any(axis=0)]
    # Lösche alte Inhalte
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

def start_app():
    root = tk.Tk()
    root.title("Preis-DB & Online-Preise")
    root.geometry("1200x600")

    # --- GUI Elements ---
    frame = tk.Frame(root)
    frame.pack(fill="x", padx=10, pady=4)
    tk.Label(frame, text="Artikelnummer oder SAP-Nummer:").pack(side="left")
    entry = tk.Entry(frame, width=35)
    entry.pack(side="left", padx=5)
    search_btn = tk.Button(frame, text="Suche", width=12)
    search_btn.pack(side="left", padx=5)
    bom_btn = tk.Button(frame, text="BOM laden & Suchen", width=18, state="disabled")
    bom_btn.pack(side="left", padx=5)
    stop_btn = tk.Button(frame, text="Stopp", width=8, state="disabled")
    stop_btn.pack(side="left", padx=5)
    export_btn = tk.Button(frame, text="Export als Excel", width=15)
    export_btn.pack(side="left", padx=5)

    status_label = tk.Label(root, text="", fg="blue")
    status_label.pack(anchor="w", padx=10)

    result_frame = tk.Frame(root)
    result_frame.pack(fill="both", expand=True, padx=10, pady=4)
    tree = ttk.Treeview(result_frame, columns=[], show="headings")
    tree.pack(fill="both", expand=True)

    # Datenbankdatei laden
    file = filedialog.askopenfilename(
        title="Excel-Datei wählen", filetypes=[("Excel-Dateien", "*.xls*")]
    )
    if not file:
        messagebox.showinfo("Abgebrochen", "Keine Datei gewählt.")
        root.destroy()
        return

    df = load_excel(file)
    search_cols = ["WN_SAP-Artikel-NR", "WN_HerstellerBestellnummer_1"]
    missing_cols = [col for col in search_cols if col not in df.columns]
    if missing_cols:
        messagebox.showerror(
            "Fehler", f"Folgende Spalten fehlen in der Tabelle: {missing_cols}"
        )
        root.destroy()
        return

    # Erst jetzt BOM Button aktivieren
    bom_btn.config(state="normal")

    # --- Suche ---
    def do_search(event=None):
        search = entry.get().strip()
        if not search:
            return
        db_rows = search_and_show(df, search, search_cols)
        if db_rows is not None and 'WN_HerstellerBestellnummer_1' in db_rows.columns:
            artikelnummer = db_rows.iloc[0]['WN_HerstellerBestellnummer_1']
        else:
            artikelnummer = search
        online_results_list = get_online_results(artikelnummer)
        merged = merge_results(db_rows, online_results_list)
        if merged is not None and not merged.empty:
            show_table(merged, tree)
            tree.gesamt_df = merged  # Für Export
        else:
            messagebox.showinfo(
                "Kein Treffer", f"Keine Zeile mit '{search}' gefunden (weder DB noch online)."
            )

    # --- BOM laden & Suchen ---
    def load_bom_and_search():
        bomfile = filedialog.askopenfilename(
            title="BOM-Datei wählen",
            filetypes=[("Excel/CSV", "*.xls*;*.csv")]
        )
        if not bomfile:
            return

        try:
            # Falls BOM-Header z.B. ab Zeile 6 (also Header=6 → Zeile 7), dann anpassen:
            bom_df = None
            try:
                bom_df = pd.read_excel(bomfile, header=6)
            except Exception:
                bom_df = read_bom(bomfile)

            sap_col, art_col = detect_both_part_columns(bom_df)
        except Exception as e:
            messagebox.showerror("BOM-Fehler", str(e))
            return

        # Artikelnummern extrahieren (SAP & Artikel)
        suchwerte = []
        if sap_col:
            suchwerte += list(bom_df[sap_col].dropna().astype(str).unique())
        if art_col and art_col != sap_col:
            suchwerte += list(bom_df[art_col].dropna().astype(str).unique())
        # Duplikate & Leere raus:
        suchwerte = [v for v in pd.unique(suchwerte) if v and v.lower() != "nan"]

        results_list = []
        stop_flag = [False]

        def do_bom_search(i=0):
            if stop_flag[0] or i >= len(suchwerte):
                status_label.config(text="Abgebrochen." if stop_flag[0] else "Fertig.")
                stop_btn.config(state="disabled")
                return

            status_label.config(text=f"({i+1}/{len(suchwerte)}) Suche: {suchwerte[i]}")
            root.update_idletasks()

            suchwert = suchwerte[i]
            db_rows = search_and_show(df, suchwert, search_cols)
            if db_rows is not None and 'WN_HerstellerBestellnummer_1' in db_rows.columns:
                artikelnummer = db_rows.iloc[0]['WN_HerstellerBestellnummer_1']
            else:
                artikelnummer = suchwert
            online_results_list = get_online_results(artikelnummer)
            merged = merge_results(db_rows, online_results_list)
            if merged is not None and not merged.empty:
                results_list.append(merged)
                # Im Treeview immer aktuell alles zeigen (wie Excel, "Blockweise")
                show_table(pd.concat(results_list, ignore_index=True), tree)
                tree.gesamt_df = pd.concat(results_list, ignore_index=True)

            # Nächsten Suchwert abarbeiten (mit kurzer Pause)
            root.after(30, lambda: do_bom_search(i+1))

        # Start!
        stop_flag[0] = False
        stop_btn.config(state="normal")
        status_label.config(text="Starte BOM-Auswertung...")
        results_list.clear()
        show_table(pd.DataFrame(), tree)
        root.after(100, lambda: do_bom_search(0))

        def stop_search():
            stop_flag[0] = True
            stop_btn.config(state="disabled")
        stop_btn.config(command=stop_search)

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

    search_btn.config(command=do_search)
    entry.bind("<Return>", do_search)
    bom_btn.config(command=load_bom_and_search)
    export_btn.config(command=export_as_excel)

    root.mainloop()
