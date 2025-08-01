# Dieses Modul enthält die Logik zur Automatisierung von Excel über win32com, um Preise zu aktualisieren.

import os
from datetime import datetime
from tkinter import messagebox
import win32com.client
from utils import clean_price
from config import EXCEL_SHEET_NAME, EXCEL_SHEET_PASSWORD

def normalize_losgroesse(val):
    try:
        return str(int(float(str(val).replace(",", ".").strip())))
    except (ValueError, TypeError):
        return str(val).strip().lower()

def normalize_quelle(val):
    return str(val).strip().lower().split("(")[0].strip()

def normalize_nummer_1000er(val):
    try:
        return str(int(float(str(val).replace(",", ".").strip())))
    except (ValueError, TypeError):
        return str(val).strip().lower()

def build_excel_index(ws, max_row):
    """Erstellt einen Index aus der Excel-Datei für schnellen Zugriff."""
    index = {}
    for row in range(8, max_row + 1, 4):
        artikelnummer = str(ws.Cells(row, 2).Value).strip().lower()
        nummer_1000er = normalize_nummer_1000er(ws.Cells(row, 3).Value)
        losgroesse = normalize_losgroesse(ws.Cells(row + 2, 24).Value)
        quelle = normalize_quelle(ws.Cells(row + 3, 24).Value)
        key = (artikelnummer, nummer_1000er, losgroesse, quelle)
        index[key] = row
    return index

def update_excel_prices_win32com(excel_path, updates_per_entry, progress_var=None, status_label=None, root=None):
    """Aktualisiert Preise in der Excel-Datei mit win32com und ruft das Makro explizit auf."""
    excel = None
    wb = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.ScreenUpdating = False

        wb = excel.Workbooks.Open(os.path.abspath(excel_path))
        ws = wb.Worksheets(EXCEL_SHEET_NAME)
        
        if EXCEL_SHEET_PASSWORD:
            ws.Unprotect(EXCEL_SHEET_PASSWORD)
            
        max_row = ws.UsedRange.Rows.Count
        excel_index = build_excel_index(ws, max_row)
        
        total_entries = len(updates_per_entry)

        # Interne Funktion, um das Schreiben, Makro-Aufrufen und Speichern zu kapseln
        def write_and_trigger(target_row, price_block):
            try:
                # 1. Daten mit korrektem Typ schreiben
                ws.Cells(target_row, 24).Value = datetime.strptime(price_block[0], "%d.%m.%Y")
                ws.Cells(target_row + 1, 24).Value = float(clean_price(price_block[1]))
                ws.Cells(target_row + 2, 24).Value = int(price_block[2])
                ws.Cells(target_row + 3, 24).Value = str(price_block[3])
                
                # 2. KORREKTUR: Makro explizit aufrufen, wie im alten Skript
                excel.Application.Run(f"'{wb.Name}'!NewPricesInDB")
                
                # 3. KORREKTUR: Arbeitsmappe sofort speichern, um die Änderung zu übernehmen
                wb.Save()
                return True
            except Exception as e:
                print(f"Fehler beim Schreiben/Makro-Aufruf für Zeile {target_row}: {e}")
                return False

        for entry_idx, entry_upd in enumerate(updates_per_entry):
            artikelnummer = str(entry_upd.get("artikelnummer", "")).strip().lower()
            nummer_1000er = normalize_nummer_1000er(entry_upd.get("1000ernummer", ""))

            for source_upd in entry_upd['sources']:
                price_block = source_upd['price_block']
                key = (artikelnummer, nummer_1000er, normalize_losgroesse(price_block[2]), normalize_quelle(price_block[3]))
                row = excel_index.get(key)

                if row:
                    write_and_trigger(row, price_block)
                else:
                    # Neuen leeren Platz für diese Artikelnummer finden und befüllen
                    for row_candidate in range(8, max_row + 1, 4):
                        zelle_artikel = str(ws.Cells(row_candidate, 2).Value).strip().lower()
                        if artikelnummer == zelle_artikel:
                            if all(ws.Cells(row_candidate + i, 24).Value is None for i in range(4)):
                                if write_and_trigger(row_candidate, price_block):
                                    break
            
            if progress_var and status_label and root:
                progress = (entry_idx + 1) / total_entries * 100
                root.after(0, lambda p=progress: progress_var.set(p))
                root.after(0, lambda i=entry_idx: status_label.config(text=f"Aktualisiere Einträge: {i + 1} / {total_entries}"))

        if EXCEL_SHEET_PASSWORD:
            ws.Protect(EXCEL_SHEET_PASSWORD, DrawingObjects=True, Contents=True, Scenarios=True, AllowFiltering=True)
        
        wb.Close(SaveChanges=True)
        wb = None

        if root and status_label:
            root.after(0, status_label.config, {"text": "Preis-Update abgeschlossen."})

    except Exception as e:
        messagebox.showerror("Excel-Fehler", f"Ein Fehler ist beim Schreiben in Excel aufgetreten:\n{e}")
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            excel.Quit()
        ws = None
        wb = None