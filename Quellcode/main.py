# Dies ist der Haupteinstiegspunkt der Anwendung. Er importiert und startet die GUI.

import os
from dotenv import load_dotenv

# KORREKTUR: Lade die Umgebungsvariablen ganz am Anfang, bevor andere Module importiert werden.
# Dadurch wird sichergestellt, dass die Variablen für die gesamte Anwendung verfügbar sind.
load_dotenv(dotenv_path=os.path.join("venv", ".env"))

from gui import start_app

if __name__ == "__main__":
    start_app()