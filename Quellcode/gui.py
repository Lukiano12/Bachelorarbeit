# Diese Datei dient als "Klebstoff", der die Benutzeroberfläche (UIManager) und die Anwendungslogik (EventHandlers) zusammenfügt und die App startet.

import tkinter as tk
from ui_manager import UIManager
from event_handlers import EventHandlers

def start_app():
    """Initialisiert und startet die Anwendung."""
    root = tk.Tk()
    
    # 1. Erstelle die Benutzeroberfläche
    ui = UIManager(root)
    
    # 2. Erstelle die Logik-Handler und verbinde sie mit der UI
    handlers = EventHandlers(ui)
    
    # 3. Verknüpfe die Buttons mit den Handler-Funktionen
    ui.search_btn.config(command=handlers.do_search)
    ui.entry.bind("<Return>", handlers.do_search)
    ui.bom_btn.config(command=handlers.load_bom_and_search)
    ui.export_btn.config(command=handlers.export_as_excel)
    ui.update_db_btn.config(command=handlers.update_db_from_excel)
    ui.update_excel_btn.config(command=handlers.update_selected_prices_in_excel)
    ui.tree.bind("<Button-1>", handlers.on_tree_click)
    
    # 4. Starte die Hauptschleife der Anwendung
    root.mainloop()