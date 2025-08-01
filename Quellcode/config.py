# Diese Datei enthält zentrale Konfigurationen und Konstanten für die gesamte Anwendung.
# Zentrale Konfigurationsdatei

import os

# Dateiname für den JSON-Cache
DB_JSON_FILE = "database.json"

# Spalten, die für die Suche in der Datenbank verwendet werden
SEARCH_COLS = ["WN_SAP-Artikel-NR", "WN_HerstellerBestellnummer_1"]

# Schlüsselwörter zur Identifizierung von Online-Preis-Spalten
ONLINE_KEYWORDS = ["mouser", "octopart", "digi-key", "arrow", "online", "connector"]

# Spalten, die in der Tabellenansicht ausgeblendet werden sollen
HIDE_COLS = ["ENTRY", "Description_deutsch_2"]

# Name des Excel-Arbeitsblatts
EXCEL_SHEET_NAME = "DB_4erDS"

# Passwort für den Blattschutz in Excel
EXCEL_SHEET_PASSWORD = os.getenv("EXCEL_PASSWORD")