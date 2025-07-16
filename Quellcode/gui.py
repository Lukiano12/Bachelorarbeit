import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from excel_search import load_excel, search_and_show, merge_results
from online_sources import get_online_results

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

def start_app():
    root = tk.Tk()
    root.title("Preis-DB & Online-Preise")
    root.geometry("1200x400")

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

    # GUI elements
    frame = tk.Frame(root)
    frame.pack(fill="x", padx=10, pady=4)
    tk.Label(frame, text="Artikelnummer oder SAP-Nummer:").pack(side="left")
    entry = tk.Entry(frame, width=35)
    entry.pack(side="left", padx=5)
    search_btn = tk.Button(frame, text="Suche", width=12)
    search_btn.pack(side="left", padx=5)

    result_frame = tk.Frame(root)
    result_frame.pack(fill="both", expand=True, padx=10, pady=4)
    tree = ttk.Treeview(result_frame, columns=[], show="headings")
    tree.pack(fill="both", expand=True)

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
        else:
            messagebox.showinfo(
                "Kein Treffer", f"Keine Zeile mit '{search}' gefunden (weder DB noch online)."
            )

    search_btn.config(command=do_search)
    entry.bind("<Return>", do_search)
    root.mainloop()
