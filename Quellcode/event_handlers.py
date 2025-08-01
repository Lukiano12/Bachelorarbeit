# Diese Datei enthält die Kernlogik der Anwendung. Sie behandelt alle Benutzerinteraktionen (Events) aus der GUI.

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import threading
from excel_search import load_excel, search_and_show, merge_results
from online_sources import get_online_results
from bom_tools import read_bom, detect_both_part_columns
from data_manager import load_db_from_json, save_db_to_json
from excel_updater import update_excel_prices_win32com
from utils import is_online_source
from config import SEARCH_COLS

class EventHandlers:
    def __init__(self, ui_manager):
        self.ui = ui_manager
        self.df = None
        self.initialize_db()

    def initialize_db(self):
        self.df = load_db_from_json()
        if self.df is not None and all(col in self.df.columns for col in SEARCH_COLS):
            self.ui.search_btn.config(state="normal")
            self.ui.bom_btn.config(state="normal")
        else:
            messagebox.showinfo("Keine Datenbank", "Keine Datenbank gefunden. Bitte Excel-Datei zur Initialisierung laden.")
            self.update_db_from_excel()

    def update_db_from_excel(self):
        file = filedialog.askopenfilename(title="Excel-Datei wählen", filetypes=[("Excel-Dateien", "*.xls*")])
        if not file: return
        try:
            df_loaded = load_excel(file)
            if any(col not in df_loaded.columns for col in SEARCH_COLS):
                messagebox.showerror("Fehler", f"Benötigte Spalten fehlen: {SEARCH_COLS}")
                return
            self.df = df_loaded
            save_db_to_json(self.df)
            self.ui.search_btn.config(state="normal")
            self.ui.bom_btn.config(state="normal")
            messagebox.showinfo("Erfolg", "Datenbank wurde aktualisiert.")
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Laden der Datei:\n{e}")

    def do_search(self, event=None):
        if self.df is None: return
        search_term = self.ui.entry.get().strip()
        if not search_term: return

        def worker():
            db_rows = search_and_show(self.df, search_term, SEARCH_COLS)
            artikelnummer = db_rows.iloc[0]['WN_HerstellerBestellnummer_1'] if (db_rows is not None and not db_rows.empty) else search_term
            
            online_results = get_online_results(artikelnummer) if self.ui.use_online_var.get() else []
            
            merged, veraltet_indices = merge_results(db_rows, online_results)
            
            if merged is not None and not merged.empty:
                self.ui.root.after(0, lambda: self.ui.show_table(merged, veraltet_indices))
                self.ui.tree.anzeige_df = merged
            else:
                self.ui.root.after(0, lambda: messagebox.showinfo("Kein Treffer", f"Keine Daten für '{search_term}' gefunden."))
        
        threading.Thread(target=worker, daemon=True).start()

    def update_selected_prices_in_excel(self):
        selected_items = [item for item in self.ui.tree.get_children() if self.ui.tree.index(item) % 4 == 0 and self.ui.tree.set(item, "Auswahl") == "✓"]
        if not selected_items:
            messagebox.showinfo("Keine Auswahl", "Bitte mindestens einen Block auswählen.")
            return

        anzeige_df = getattr(self.ui.tree, "anzeige_df", None)
        if anzeige_df is None: return

        excel_path = filedialog.askopenfilename(title="Excel-Datei für Update wählen", filetypes=[("Makro-fähige Excel", "*.xlsm")])
        if not excel_path: return

        updates_per_entry = []
        for item in selected_items:
            idx = self.ui.tree.index(item)
            block_rows = anzeige_df.iloc[idx:idx+4]
            artikelnummer = str(block_rows.iloc[0][anzeige_df.columns[0]])
            nummer_1000er = str(block_rows.iloc[0][anzeige_df.columns[1]]) if len(anzeige_df.columns) > 1 else ""
            
            sources = [{'price_block': [block_rows.iloc[j][col] for j in range(4)]}
                       for col in anzeige_df.columns if is_online_source(col) and all(str(block_rows.iloc[j][col]).strip() for j in range(4))]
            
            if sources:
                updates_per_entry.append({'artikelnummer': artikelnummer, '1000ernummer': nummer_1000er, 'sources': sources})

        if not updates_per_entry:
            messagebox.showinfo("Kein Update", "Keine vollständigen Online-Blöcke zum Aktualisieren gefunden.")
            return

        def worker():
            update_excel_prices_win32com(excel_path, updates_per_entry, self.ui.progress_var, self.ui.status_label, self.ui.root)
            self.ui.root.after(0, lambda: messagebox.showinfo("Fertig", "Excel-Datei wurde aktualisiert."))
        
        threading.Thread(target=worker, daemon=True).start()

    def load_bom_and_search(self):
        if self.df is None: return
        bomfile = filedialog.askopenfilename(title="BOM-Datei wählen", filetypes=[("Excel/CSV", "*.xls*;*.csv")])
        if not bomfile: return

        def worker():
            try:
                bom_df = read_bom(bomfile)
                _, art_col = detect_both_part_columns(bom_df)
                bauteile = [str(t).strip() for t in bom_df[art_col].dropna().unique() if str(t).strip().upper() != "SPLICE"]
                total = len(bauteile)
                all_results = []
                all_veraltet = []

                for idx, part in enumerate(bauteile):
                    progress = (idx + 1) / total * 100
                    self.ui.root.after(0, lambda p=progress: self.ui.progress_var.set(p))
                    self.ui.root.after(0, lambda i=idx: self.ui.status_label.config(text=f"Lade BOM: {i+1}/{total}"))
                    
                    db_rows = search_and_show(self.df, part, SEARCH_COLS)
                    online_res = get_online_results(part) if self.ui.use_online_var.get() else []
                    merged, veraltet = merge_results(db_rows, online_res)
                    if merged is not None:
                        offset = sum(len(r) for r in all_results)
                        all_results.append(merged)
                        all_veraltet.extend([i + offset for i in veraltet])
                
                if not all_results:
                    self.ui.root.after(0, lambda: messagebox.showinfo("Info", "Keine Ergebnisse für BOM gefunden."))
                    return

                final_df = pd.concat(all_results, ignore_index=True).fillna('')
                self.ui.root.after(0, lambda: self.ui.show_table(final_df, all_veraltet))
                self.ui.tree.anzeige_df = final_df
            except Exception as e:
                self.ui.root.after(0, lambda: messagebox.showerror("BOM-Fehler", str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def export_as_excel(self):
        df_to_export = getattr(self.ui.tree, "anzeige_df", None)
        if df_to_export is None or df_to_export.empty:
            messagebox.showwarning("Export", "Keine Daten zum Exportieren.")
            return
        fname = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if fname:
            df_to_export.to_excel(fname, index=False)
            messagebox.showinfo("Export", f"Erfolgreich gespeichert:\n{fname}")

    def on_tree_click(self, event):
        item = self.ui.tree.identify_row(event.y)
        col = self.ui.tree.identify_column(event.x)
        if item and col == "#1" and self.ui.tree.index(item) % 4 == 0:
            current = self.ui.tree.set(item, "Auswahl")
            self.ui.tree.set(item, "Auswahl", "✓" if current != "✓" else "")