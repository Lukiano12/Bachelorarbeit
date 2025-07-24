from datetime import date
import requests, re
from bs4 import BeautifulSoup

def ac_price(article):
    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        url     = f"https://www.automotive-connectors.com/en/search?search={article}"
        resp    = requests.get(url, headers=headers, timeout=10)
        soup    = BeautifulSoup(resp.text, "html.parser")

        # 1. Prüfe, ob wir auf einer Produkt-Detailseite sind
        detail = soup.find("div", class_="product-detail-main")
        if detail:
            sku = detail.find("span", class_="product-detail-ordernumber")
            part_number = sku.get_text(strip=True) if sku else ""
            price_rows = detail.select("tr.product-block-prices-row")
            if not price_rows:
                return None
            last = price_rows[-1]
            qty_txt = last.select_one(".product-block-prices-quantity").get_text(strip=True)
            qty     = re.sub(r"[^\d]", "", qty_txt) or ""
            price_cell = last.find_all("td", class_="product-block-prices-cell")[-1]
            price_txt  = price_cell.get_text(" ", strip=True)
            price_search = re.search(r"(\d+[.,]?\d*)", price_txt)
            if price_search:
                price_float = float(price_search.group(1).replace(",", "."))
                discounted_price = price_float * 0.7
                price_fmt = "{:.2f} €".format(discounted_price).replace(".", ",")
            else:
                price_fmt = price_txt

            return {
                "Datum": date.today().strftime("%d.%m.%Y"),
                "Preis": price_fmt,
                "Losgröße": qty,
                "Quelle": "Automotive-Connectors (-30%)"
            }

        # 2. Suchergebnisseite: Suche gezielt nach Einzelstück (-L)
        boxes = soup.find_all("div", class_="product-box")
        if not boxes:
            return None

        # Suche nach Produktnummer mit -L
        preferred_box = None
        preferred_pn = f"{article}-L".replace("--", "-")
        for box in boxes:
            pn_tag = box.find(string=re.compile(r"Part number:"))
            if pn_tag:
                part_number = pn_tag.find_next().text.strip()
                if part_number.replace(" ", "").upper() == preferred_pn.replace(" ", "").upper():
                    preferred_box = box
                    break

        # Wenn gefunden, nutze diesen Eintrag, sonst wie bisher das erste Ergebnis
        box = preferred_box if preferred_box else boxes[0]

        # Hole ggf. Detailseite, falls Preisstaffel fehlt
        rows = box.select("tr.product-block-prices-row")
        if not rows:
            link = box.find("a", class_="product-name")["href"]
            if not link.startswith("http"):
                link = "https://www.automotive-connectors.com" + link
            detail_html = requests.get(link, headers=headers, timeout=10).text
            detail_soup = BeautifulSoup(detail_html, "html.parser")
            price_rows = detail_soup.select("tr.product-block-prices-row")
            if not price_rows:
                return None
            last = price_rows[-1]
            qty_txt = last.select_one(".product-block-prices-quantity").get_text(strip=True)
            qty     = re.sub(r"[^\d]", "", qty_txt) or ""
            price_cell = last.find_all("td", class_="product-block-prices-cell")[-1]
            price_txt  = price_cell.get_text(" ", strip=True)
            price_search = re.search(r"(\d+[.,]?\d*)", price_txt)
            if price_search:
                price_float = float(price_search.group(1).replace(",", "."))
                discounted_price = price_float * 0.7
                price_fmt = "{:.2f} €".format(discounted_price).replace(".", ",")
            else:
                price_fmt = price_txt

            return {
                "Datum": date.today().strftime("%d.%m.%Y"),
                "Preis": price_fmt,
                "Losgröße": qty,
                "Quelle": "Automotive-Connectors (-30%)"
            }
        last = rows[-1]
        qty_txt = last.select_one(".product-block-prices-quantity").get_text(strip=True)
        qty     = re.sub(r"[^\d]", "", qty_txt) or ""
        price_cell = last.find_all("td", class_="product-block-prices-cell")[-1]
        price_txt  = price_cell.get_text(" ", strip=True)
        price_search = re.search(r"(\d+[.,]?\d*)", price_txt)
        if price_search:
            price_float = float(price_search.group(1).replace(",", "."))
            discounted_price = price_float * 0.7
            price_fmt = "{:.2f} €".format(discounted_price).replace(".", ",")
        else:
            price_fmt = price_txt

        return {
            "Datum": date.today().strftime("%d.%m.%Y"),
            "Preis": price_fmt,
            "Losgröße": qty,
            "Quelle": "Automotive-Connectors (-30%)"
        }
    except Exception:
        return None