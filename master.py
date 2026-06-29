"""
利用者マスタ管理 — CRUD・変更履歴・リタリコCSV取込
storage.py 経由で Google Sheets / ローカルCSV に二重保存する。
"""
import re
import pandas as pd

from storage import (
    load_df, save_df, save_history,
    load_history_list,
    load_history_df,
)

MASTER_COLUMNS = [
    "氏名", "フリガナ", "区分", "地区",
    "迎え先（平日", "迎え先（長期休み）",
    "迎え時刻（平日", "迎え時刻（長期休み）",
    "送り先", "送り時刻",
    "利用曜日", "住所", "特記事項", "備考"
]


def load_master(館: str) -> pd.DataFrame:
    df = load_df(館, "master", columns=MASTER_COLUMNS)
    if df.empty:
        return df
    rename_map = {
        "迎え先":   "迎え先（平日",
        "下校時刻": "迎え時刻（平日）",
        "自宅時刻": "送り時刻",
    }
    for old, new in rename_map.items():
        if old in df.columns and new not in df.columns:
            df.rename(columns={old: new}, inplace=True)
    for col in MASTER_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[MASTER_COLUMNS]


def save_master(館: str, df: pd.DataFrame):
    for col in MASTER_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    clean = df[MASTER_COLUMNS].copy()
    save_df(館, "master", clean)
    save_history(館, clean)


def load_history(館: str, timestamp: str) -> pd.DataFrame:
    return load_history_df(館, timestamp)


def import_from_ritalico(raw_df: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    result = pd.DataFrame(index=range(len(raw_df)))
    for app_col, csv_col in mapping.items():
        if csv_col and csv_col != "（未選択）" and csv_col in raw_df.columns:
            result[app_col] = raw_df[csv_col].fillna("").astype(str)
        else:
            result[app_col] = ""
    if result.get("住所", pd.Series([""]*len(raw_df))).eq("").all():
        parts = []
        for c in ["都道府県","市区町村","番地","ビル・マンション名"]:
            if c in raw_df.columns:
                parts.append(raw_df[c].fillna("").astype(str))
        if parts:
            result["住所"] = pd.concat(parts, axis=1).apply(
                lambda r: "".join(v for v in r if v and v != "None"), axis=1)
    if "住所" in result.columns:
        result["地区"] = result["住所"].apply(_extract_district)
    if "迎え先" in result.columns and "迎え先（平日）" not in result.columns:
        result.rename(columns={"迎え先":"迎え先（平日）"}, inplace=True)
    for col in MASTER_COLUMNS:
        if col not in result.columns:
            result[col] = ""
    return result[MASTER_COLUMNS].fillna("").replace("None","")


def _extract_district(address: str) -> str:
    import re
    if not address:
        return ""
    m = re.search(r"[市区町村郡](.+?)(?:[0-9０-９\-ー－]|$)", address)
    if m:
        return m.group(1).strip()
    m = re.search(r"([一-龥ぁ-んァ-ン]+)(?=[0-9０-９])", address)
    return m.group(1) if m else ""
