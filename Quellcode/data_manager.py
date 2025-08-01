# Dieses Modul verwaltet das Laden und Speichern des lokalen Daten-Caches (database.json).

import os
import pandas as pd
from datetime import datetime
from config import DB_JSON_FILE

def load_db_from_json():
    if not os.path.exists(DB_JSON_FILE):
        return None
    try:
        df = pd.read_json(DB_JSON_FILE, orient="split")
        # Konvertiert Timestamps, die als Zahlen gespeichert wurden, zurück in Datums-Objekte
        block_size = 4
        for col in df.columns:
            for i in range(0, len(df), block_size):
                val = df.iloc[i][col]
                if isinstance(val, (int, float)) and not isinstance(val, bool) and val > 1e12:
                    try:
                        df.iloc[i, df.columns.get_loc(col)] = pd.to_datetime(val, unit='ms')
                    except Exception:
                        pass
        return df
    except Exception as e:
        print(f"Fehler beim Laden von JSON: {e}")
        return None

def save_db_to_json(df):
    try:
        df.to_json(DB_JSON_FILE, orient="split", force_ascii=False)
        print(f"Datenbank gespeichert unter: {os.path.abspath(DB_JSON_FILE)}")
    except Exception as e:
        print(f"Fehler beim Speichern zu JSON: {e}")