import pandas as pd
from utils import format_value, sapnr_to_str
from datetime import datetime, timedelta

def load_excel(file):
    sheet = "DB_4erDS"
    df = pd.read_excel(file, sheet_name=sheet, header=6)
    cols_to_drop = [
        "Unnamed: 0", "Unnamed: 17", "Unnamed: 18", "Unnamed: 19",
        "Unnamed: 20", "Unnamed: 21", "Unnamed: 22",
        "WN_PinClass", "WN_PolCount_NUM", "WN_Color", "WN_Min_CrossSection", "WN_Max_CrossSection" 
    ]
    if "WN_SAP-Artikel-NR" in df.columns:
        df["WN_SAP-Artikel-NR"] = df["WN_SAP-Artikel-NR"].apply(sapnr_to_str)
    df = df.drop(columns=[col for col in cols_to_drop if col in df.columns])
    return df

def search_and_show(df, search, search_cols):
    search = str(search).strip()
    search_df = df[search_cols].copy()
    if 'WN_SAP-Artikel-NR' in search_cols:
        search_df['WN_SAP-Artikel-NR'] = search_df['WN_SAP-Artikel-NR'].apply(
            lambda x: str(int(float(x))) if pd.notnull(x) and str(x).replace('.', '', 1).isdigit() else ""
        )
    if 'WN_HerstellerBestellnummer_1' in search_cols:
        search_df['WN_HerstellerBestellnummer_1'] = search_df['WN_HerstellerBestellnummer_1'].astype(str).map(
            lambda x: x.strip() if x != 'nan' else ""
        )
    mask = search_df.apply(lambda row: any(search == str(cell) for cell in row), axis=1)
    indices = mask[mask].index
    if len(indices) == 0:
        return None
    start = indices[0]
    end = min(start + 4, len(df))
    return df.iloc[start:end]

def merge_results(db_rows, online_results_list):
    from datetime import datetime, timedelta

    if db_rows is None or db_rows.empty:
        if online_results_list:
            data = {}
            for res in online_results_list:
                if res:
                    colname = res.get("Quelle", "Online")
                    data[colname] = [
                        format_value(res.get("Datum", ""), "Datum"),
                        format_value(res.get("Preis", ""), "Preis"),
                        res.get("Losgröße", ""),
                        res.get("Quelle", "")
                    ]
            if not data:
                return None, []
            return pd.DataFrame(data), []
        else:
            return None, []

    df = db_rows.copy()
    for col in df.columns:
        df[col] = df[col].astype(object)  # Typumwandlung für Kompatibilität
        df.iloc[0, df.columns.get_loc(col)] = format_value(df.iloc[0, df.columns.get_loc(col)], "Datum")
        df.iloc[1, df.columns.get_loc(col)] = format_value(df.iloc[1, df.columns.get_loc(col)], "Preis")
    for res in online_results_list:
        if res:
            colname = res.get("Quelle", "Online")
            df[colname] = [
                format_value(res.get("Datum", ""), "Datum"),
                format_value(res.get("Preis", ""), "Preis"),
                res.get("Losgröße", ""),
                res.get("Quelle", ""),
            ]
    block_size = 4
    today = datetime.today()
    veraltet_indices = []
    for i in range(0, len(df), block_size):
        block_veraltet = False
        for col in df.columns:
            for j in range(block_size):
                if i + j < len(df):
                    val = str(df.iloc[i + j][col]).strip()
                    # Unix-Timestamp prüfen
                    if val.isdigit() and len(val) >= 12:
                        try:
                            ts = int(val)
                            if ts > 1e12:
                                ts = ts // 1000
                            d = datetime.fromtimestamp(ts)
                            if d < today - timedelta(days=365):
                                block_veraltet = True
                        except Exception:
                            continue
                    else:
                        # Normales Datum prüfen
                        try:
                            d = datetime.strptime(val, "%d.%m.%Y")
                            if d < today - timedelta(days=365):
                                block_veraltet = True
                        except Exception:
                            continue
        if block_veraltet:
            veraltet_indices.extend(range(i, min(i + block_size, len(df))))
    return df, veraltet_indices