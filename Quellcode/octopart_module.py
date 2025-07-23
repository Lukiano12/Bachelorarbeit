from datetime import date
import requests
import os
from dotenv import load_dotenv

load_dotenv(dotenv_path=os.path.join("venv", ".env"))

OCTOPART_API_KEY = os.getenv("OCTOPART_API_KEY")

def get_usd_to_eur():
    url = "https://api.exchangerate.host/latest?base=USD&symbols=EUR"
    try:
        response = requests.get(url, timeout=5)
        data = response.json()
        return data["rates"]["EUR"]
    except Exception as e:
        print("Konnte Wechselkurs nicht abrufen:", e)
        return 0.92

def octopart_price_nexar(article, octopart_api_key=OCTOPART_API_KEY):
    url = "https://api.nexar.com/graphql"
    headers = {
        "Authorization": f"Bearer {octopart_api_key}",
        "Content-Type": "application/json"
    }
    query = """
    query Search($mpn: String!) {
      supSearch(q: $mpn, limit: 1) {
        results {
          part {
            mpn
            manufacturer { name }
            sellers {
              company { name }
              offers {
                clickUrl
                prices {
                  quantity
                  price
                  currency
                }
              }
            }
          }
        }
      }
    }
    """
    variables = {"mpn": article}
    try:
        response = requests.post(url, json={"query": query, "variables": variables}, headers=headers, timeout=10)
        if response.status_code != 200:
            return None
        data = response.json()
        offers = data["data"]["supSearch"]["results"][0]["part"]["sellers"][0]["offers"]

        price_data = None
        # 1. Versuche EUR
        for offer in offers:
            for p in offer["prices"]:
                if p["currency"] == "EUR":
                    price_data = p
                    break
            if price_data:
                break
        # 2. Falls kein EUR, nehme USD und rechne um
        usd_converted = False
        if not price_data:
            for offer in offers:
                for p in offer["prices"]:
                    if p["currency"] == "USD":
                        price_data = p
                        usd_converted = True
                        break
                if price_data:
                    break

        if price_data:
            qty = price_data["quantity"]
            price = price_data["price"] * 0.7  
            if usd_converted:
                usd_eur = get_usd_to_eur()
                price_eur = price * usd_eur
                preis_formatiert = f"{price_eur:.4f} €".replace(".", ",")
            else:
                preis_formatiert = f"{price:.4f} €".replace(".", ",")
            return {
                "Datum": date.today().strftime("%d.%m.%Y"),
                "Preis": preis_formatiert,
                "Losgröße": qty,
                "Quelle": "Octopart (-30%)"
            }
        else:
            return None
    except Exception as e:
        print("Octopart-Parsing-Fehler:", e)
        return None
