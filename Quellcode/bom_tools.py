import pandas as pd

def read_bom(bomfile, header=6):
    if bomfile.lower().endswith((".xls", ".xlsx")):
        bom_df = pd.read_excel(bomfile, header=header)
    elif bomfile.lower().endswith(".csv"):
        bom_df = pd.read_csv(bomfile, sep=None, engine='python', header=header)
    else:
        raise Exception("Nur Excel oder CSV unterstützt.")
    return bom_df

def detect_both_part_columns(bom_df):
    columns = [str(c).strip().replace("\n", "").replace("\r", "").replace(" ", "").lower() for c in bom_df.columns]
    sap_patterns = ["saparticleno", "wn_sap-articleno", "sap", "material", "artikel"]
    art_patterns = ["manufacturerorderno", "bestell", "orderno", "manu", "hersteller"]

    sap_col = None
    art_col = None

    for i, c in enumerate(columns):
        for s in sap_patterns:
            if s in c:
                sap_col = bom_df.columns[i]
        for s in art_patterns:
            if s in c:
                art_col = bom_df.columns[i]

    if not sap_col and not art_col:
        raise Exception("Keine Spalte mit Artikelnummer/SAP erkannt!")
    return sap_col, art_col

def is_valid_sapnr(s):
    """
    Prüft, ob die SAP-Nummer eine 'echte' 1000er ist (rein numerisch, meistens 7-10 Stellen).
    """
    s = str(s).strip()
    return s.isdigit() and 7 <= len(s) <= 10
