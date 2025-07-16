from datetime import date
import requests, re
from bs4 import BeautifulSoup

def ac_price(article):
    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        url     = f"https://www.automotive-connectors.com/en/search?search={article}"
        soup    = BeautifulSoup(requests.get(url, headers=headers, timeout=10).text, "html.parser")
        box = soup.find("div", class_="product-box")
        if not box:
            return None
        rows = box.select("tr.product-block-prices-row")
        if not rows:
            link = box.find("a", class_="product-name")["href"]
            if not link.startswith("http"):
                link = "https://www.automotive-connectors.com" + link
            detail_html = requests.get(link, headers=headers, timeout=10).text
            rows = BeautifulSoup(detail_html, "html.parser").select("tr.product-block-prices-row")
            if not rows:
                return None
        last = rows[-1]
        qty_txt = last.select_one(".product-block-prices-quantity").get_text(strip=True)
        qty     = re.sub(r"[^\d]", "", qty_txt) or ""
        price_cell = last.find_all("td", class_="product-block-prices-cell")[-1]
        price_txt  = price_cell.get_text(" ", strip=True)
        # Preis extrahieren (nur Zahl, Komma/Punkt, optional Währungszeichen/Extra entfernen)
        price_float = None
        price_search = re.search(r"(\d+[.,]?\d*)", price_txt)
        if price_search:
            price_float = float(price_search.group(1).replace(",", "."))
            price_fmt = "{:.2f} €".format(price_float).replace(".", ",")
        else:
            price_fmt = price_txt  # Fallback: Original-Text
        return {
            "Datum": date.today().strftime("%d.%m.%Y"),
            "Preis": price_fmt,
            "Losgröße": qty,
            "Quelle": "Automotive-Connectors"
        }
    except Exception:
        return None
