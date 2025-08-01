# Dieses Modul ist für den Aufbau und die Verwaltung aller grafischen Elemente (Widgets) der Benutzeroberfläche zuständig.

import tkinter as tk
from tkinter import ttk
from datetime import datetime
import pandas as pd
from config import HIDE_COLS

class UIManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Preis-DB & Online-Preise")
        self.root.geometry("1600x900")
        
        style = ttk.Style()
        try:
            self.root.tk.call("source", "themes/forest-light.tcl")
            style.theme_use("forest-light")
        except tk.TclError:
            print("Forest-Theme nicht gefunden.")

        self._create_widgets()

    def _create_widgets(self):
        # Top frame for controls
        frame = ttk.Frame(self.root, padding=8)
        frame.pack(fill="x", padx=10, pady=4)

        ttk.Label(frame, text="Artikelnummer oder SAP-Nummer:").pack(side="left")
        self.entry = ttk.Entry(frame, width=30)
        self.entry.pack(side="left", padx=5)

        self.search_btn = ttk.Button(frame, text="Suche", width=16, state="disabled")
        self.search_btn.pack(side="left", padx=5)
        self.bom_btn = ttk.Button(frame, text="BOM laden & Suchen", width=20, state="disabled")
        self.bom_btn.pack(side="left", padx=5)
        self.export_btn = ttk.Button(frame, text="Export als Excel", width=18)
        self.export_btn.pack(side="left", padx=5)
        self.update_db_btn = ttk.Button(frame, text="Datenbank aktualisieren", width=22)
        self.update_db_btn.pack(side="left", padx=5)
        self.update_excel_btn = ttk.Button(frame, text="Excel-Preise aktualisieren", width=22)
        self.update_excel_btn.pack(side="left", padx=5)

        self.use_online_var = tk.BooleanVar(value=False)
        online_check = ttk.Checkbutton(frame, text="Online Quellen nutzen", variable=self.use_online_var)
        online_check.pack(side="left", padx=10)

        # Result frame with Treeview
        result_frame = ttk.Frame(self.root, padding=4)
        result_frame.pack(fill="both", expand=True, padx=10, pady=4)

        tree_scroll_y = ttk.Scrollbar(result_frame, orient="vertical")
        tree_scroll_x = ttk.Scrollbar(result_frame, orient="horizontal")
        self.tree = ttk.Treeview(
            result_frame, columns=[], show="headings",
            yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set, selectmode="extended"
        )
        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_y.pack(side="right", fill="y")
        tree_scroll_x.config(command=self.tree.xview)
        tree_scroll_x.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)

        # Status bar
        self.progress_var = tk.DoubleVar()
        progress_bar = ttk.Progressbar(self.root, variable=self.progress_var, mode="determinate")
        progress_bar.pack(fill="x", padx=10, pady=2)
        self.status_label = ttk.Label(self.root, text="Bereit")
        self.status_label.pack(fill="x", padx=10, pady=(0, 5))

    def show_table(self, df, veraltet_indices=None):
        if veraltet_indices is None:
            veraltet_indices = []
        
        df_display = df.drop(columns=[col for col in HIDE_COLS if col in df.columns], errors="ignore").copy()
        if "Status" in df_display.columns:
            df_display = df_display.drop(columns=["Status"])
        df_display.fillna('', inplace=True)
        df_display = df_display.loc[:, (df_display != '').any(axis=0)]

        self.tree["columns"] = ["Auswahl"] + list(df_display.columns)
        self.tree.heading("Auswahl", text="✓")
        self.tree.column("Auswahl", width=40, anchor="center")
        for col in df_display.columns:
            self.tree.heading(col, text=str(col))
            self.tree.column(col, width=120, anchor="center")
        
        self.tree.delete(*self.tree.get_children())
        self.tree.tag_configure("veraltet", background="#ffcccc")

        for idx, (_, row) in enumerate(df_display.iterrows()):
            values = [""]
            for col_name, x in row.items():
                if isinstance(x, (pd.Timestamp, datetime)):
                    values.append(x.strftime("%d.%m.%Y"))
                elif "Preis" in col_name and isinstance(x, (int, float)):
                    # Formatiere den Preis für die deutsche Anzeige mit 4 Nachkommastellen
                    values.append(f"{x:,.4f} €".replace(",", "X").replace(".", ",").replace("X", "."))
                else:
                    values.append(str(x))
            tags = ("veraltet",) if idx in veraltet_indices else ()
            self.tree.insert("", "end", values=values, tags=tags)