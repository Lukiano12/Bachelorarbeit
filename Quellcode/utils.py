# Diese Datei sammelt kleine, allgemeine Hilfsfunktionen, die in der gesamten Anwendung wiederverwendet werden.

import pandas as pd
import re
from config import ONLINE_KEYWORDS

def format_value(value, field=None):
    if field and "Datum" in field and pd.notna(value) and str(value).strip() != "":
        if isinstance(value, str) and re.match(r"^\d{2}\.\d{2}\.\d{4}$", value):
            return value
        try:
            date_val = pd.to_datetime(value, dayfirst=True, errors="coerce")
            return date_val.strftime("%d.%m.%Y") if pd.notna(date_val) else str(value)
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
        return str(int(float(x)))
    except (ValueError, TypeError):
        return str(x)

def clean_price(value):
    if isinstance(value, str):
        value = value.replace("€", "").replace(" ", "").replace(",", ".")
    try:
        return float(value)
    except (ValueError, TypeError):
        return value

def is_online_source(colname):
    return any(key in colname.lower() for key in ONLINE_KEYWORDS)