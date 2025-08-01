# Dieses Modul ist für das Web-Scraping von Preisen von der Automotive-Connectors Webseite zuständig.

from datetime import date
import requests, re
from bs4 import BeautifulSoup

def ac_price(article):
    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        url     = f"https://www.automotive-connectors.com/en/search?search={article}"
        resp    = requests.get(url, headers=headers, timeout=10)
        soup    = BeautifulSoup(resp.text, "html.parser")

        # --- Interne Hilfsfunktion, um Preisdaten zu extrahieren und zu formatieren ---
        def extract_price_data(price_rows):
            if not price_rows:
                return None
            
            last = price_rows[-1]
            qty_txt = last.select_one(".product-block-prices-quantity").get_text(strip=True)
            qty     = int(re.sub(r"[^\d]", "", qty_txt) or "0")
            price_cell = last.find_all("td", class_="product-block-prices-cell")[-1]
            price_txt  = price_cell.get_text(" ", strip=True)
            
            discounted_price = 0.0
            price_search = re.search(r"(\d+[.,]?\d*)", price_txt)
            if price_search:
                price_float = float(price_search.group(1).replace(",", "."))
                discounted_price = price_float * 0.7
            
            # KORREKTUR: Gib die berechneten Zahlen zurück
            return {
                "Datum": date.today().strftime("%d.%m.%Y"),
                "Preis": discounted_price,
                "Losgröße": qty,
                "Quelle": "Automotive-Connectors (-30%)"
            }

        # 1. Prüfe, ob wir auf einer Produkt-Detailseite sind
        detail = soup.find("div", class_="product-detail-main")
        if detail:
            result = extract_price_data(detail.select("tr.product-block-prices-row"))
            if result:
                return result

        # 2. Suchergebnisseite
        boxes = soup.find_all("div", class_="product-box")
        if not boxes:
            return None

        box = boxes[0]
        
        # Hole Preisdaten direkt aus der Box
        result = extract_price_data(box.select("tr.product-block-prices-row"))
        if result:
            return result
        
        # Wenn keine Preisdaten in der Box sind, lade die Detailseite
        link_tag = box.find("a", class_="product-name")
        if link_tag and link_tag.has_attr('href'):
            link = link_tag["href"]
            if not link.startswith("http"):
                link = "https://www.automotive-connectors.com" + link
            
            detail_resp = requests.get(link, headers=headers, timeout=10)
            detail_soup = BeautifulSoup(detail_resp.text, "html.parser")
            detail_result = extract_price_data(detail_soup.select("tr.product-block-prices-row"))
            if detail_result:
                return detail_result

        return None # Wenn nirgendwo Preisdaten gefunden wurden
        
    except Exception as e:
        print(f"[AC-FEHLER] Unerwarteter Fehler für '{article}': {e}")
        return None