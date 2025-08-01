# Dieses Modul ist für die Abfrage von Preisen über die Octopart/Nexar-API zuständig.

from datetime import date
import requests
import os
from dotenv import load_dotenv
import ssl
from requests.adapters import HTTPAdapter
from urllib3.poolmanager import PoolManager

# Lade Umgebungsvariablen
load_dotenv(dotenv_path=os.path.join("venv", ".env"))
OCTOPART_API_KEY = os.getenv("OCTOPART_API_KEY")

# --- SSL/TLS FIX START ---
# Eigene Adapter-Klasse, um eine bestimmte TLS-Version zu erzwingen
class Tls12Adapter(HTTPAdapter):
    def init_poolmanager(self, connections, maxsize, block=False):
        self.poolmanager = PoolManager(
            num_pools=connections,
            maxsize=maxsize,
            block=block,
            ssl_version=ssl.PROTOCOL_TLSv1_2,
        )

# Funktion, um eine Session mit dem TLS-Fix zu erstellen
def get_tls_session():
    session = requests.Session()
    session.mount("https://", Tls12Adapter())
    return session
# --- SSL/TLS FIX END ---


def get_usd_to_eur():
    url = "https://api.exchangerate.host/latest?base=USD&symbols=EUR"
    try:
        # Verwende die Session mit dem TLS-Fix
        session = get_tls_session()
        response = session.get(url, timeout=5)
        response.raise_for_status()
        data = response.json()
        return data["rates"]["EUR"]
    except Exception as e:
        print(f"Konnte Wechselkurs nicht abrufen: {e}")
        return 0.92 # Fallback-Wert

def octopart_price_nexar(article, octopart_api_key=OCTOPART_API_KEY):
    if not octopart_api_key:
        return None
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
    try:
        response = requests.post(url, headers=headers, json={"query": query, "variables": {"mpn": article}}, timeout=10)
        response.raise_for_status()
        data = response.json()
        
        if not data.get("data", {}).get("supSearch", {}).get("results"):
            return None
            
        sellers = data["data"]["supSearch"]["results"][0]["part"]["sellers"]
        if not sellers:
            return None

        offers = sellers[0]["offers"]
        if not offers:
            return None

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
            price = price_data["price"]
            
            if usd_converted:
                usd_eur = get_usd_to_eur()
                price = price * usd_eur

            return {
                "Datum": date.today().strftime("%d.%m.%Y"),
                "Preis": price * 0.7,
                "Losgröße": qty,
                "Quelle": "Octopart (-30%)"
            }
        else:
            return None
    except Exception as e:
        print(f"Octopart-Parsing-Fehler: {e}")
        return None