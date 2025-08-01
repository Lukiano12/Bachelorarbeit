import pandas as pd
import re

def format_value(value, field=None):
    if field and "Datum" in field and pd.notna(value) and str(value).strip() != "":
        # Wenn der Wert bereits ein korrekt formatierter String ist, direkt zurückgeben.
        if isinstance(value, str) and re.match(r"^\d{2}\.\d{2}\.\d{4}$", value):
            return value
        try:
            # KORREKTUR: dayfirst=True hinzugefügt, um die korrekte Reihenfolge zu erzwingen.
            date_val = pd.to_datetime(value, dayfirst=True, errors="coerce")
            if pd.notna(date_val):
                return date_val.strftime("%d.%m.%Y")
            else:
                return str(value)
        except Exception:
            return str(value)
    if field and "Preis" in field and pd.notna(value):
        try:
            val = re.sub(r"[^\d.,]", "", str(value)).replace(",", ".")
            price = float(val)
            return "{:.2f} €".format(price).replace(".", ",")
        except Exception:
            pass
    return "" if pd.isna(value) else str(value)

def sapnr_to_str(x):
    try:
        if pd.isna(x) or str(x).strip() == "" or str(x).lower() == "nan":
            return ""
        val = float(x)
        return str(int(val))
    except Exception:
        return str(x)