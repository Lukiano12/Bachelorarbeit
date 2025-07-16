import pandas as pd
from utils import format_value, sapnr_to_str

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
                return None
            return pd.DataFrame(data)
        else:
            return None

    df = db_rows.copy()
    for col in df.columns:
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
    return df
