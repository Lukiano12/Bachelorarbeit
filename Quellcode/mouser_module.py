from datetime import date
import requests, re

MOUSER_API_KEY = "8de8bdef-945e-421a-9b7d-ce77d61e1b03"  # Füge hier deinen API-Key ein!

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
        parts = r.json().get("SearchResults", {}).get("Parts", [])
        if not parts:
            return None
        part = parts[0]
        brk  = part["PriceBreaks"][-1]
        qty  = brk["Quantity"]
        raw  = brk["Price"]
        try:
            price_val = float(re.sub(r"[^\d,\.]", "", raw).replace(",", "."))
            discounted_price = price_val * 0.7
            price_txt = f"{discounted_price:.2f} €".replace(".", ",")
        except Exception:
             price_txt = raw

        return {
            "Datum": date.today().strftime("%d.%m.%Y"),
            "Preis": price_txt,
            "Losgröße": qty,
            "Quelle": "Mouser (-30%)"
        }
    except Exception:
        return None
