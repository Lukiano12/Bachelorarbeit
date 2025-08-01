# Dieses Modul ist für die Abfrage von Preisen über die Mouser-API zuständig.

from datetime import date
import requests
import re
import os

MOUSER_API_KEY = os.getenv("MOUSER_API_KEY")

def mouser_price(article, MOUSER_API_KEY=MOUSER_API_KEY):
    if not MOUSER_API_KEY:
        return None
    payload = {
        "SearchByPartRequest": {
            "mouserPartNumber": article,
            "partSearchOptions": "None"
        }
    }
    try:
        r = requests.post(
            f"https://api.mouser.com/api/v1/search/partnumber?apiKey={MOUSER_API_KEY}",
            json=payload, headers={"Content-Type":"application/json"}, timeout=10
        )
        r.raise_for_status()
        parts = r.json().get("SearchResults", {}).get("Parts", [])
        if not parts:
            return None
        
        part = parts[0]
        if not part.get("PriceBreaks"):
            return None

        brk = part["PriceBreaks"][-1]
        qty_str = brk.get("Quantity", "0")
        raw_price_str = brk.get("Price", "0")
        
        discounted_price = 0.0
        try:
            # Konvertiere den Preis-String in eine Zahl
            price_val = float(re.sub(r"[^\d,\.]", "", raw_price_str).replace(",", "."))
            # Wende den 30% Rabatt an
            discounted_price = price_val * 0.7
        except (ValueError, TypeError):
             # Falls die Konvertierung fehlschlägt, bleibt der Preis 0.0
             pass

        # Gib die berechneten Zahlen zurück, nicht die formatierten Strings.
        return {
            "Datum": date.today().strftime("%d.%m.%Y"),
            "Preis": discounted_price,
            "Losgröße": int(qty_str),
            "Quelle": "Mouser (-30%)"
        }
    except Exception:
        return None