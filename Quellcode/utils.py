import pandas as pd
import re

def format_value(value, field=None):
    if field and "Datum" in field and pd.notna(value) and str(value).strip() != "":
        try:
            date_val = pd.to_datetime(value, errors="coerce")
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
