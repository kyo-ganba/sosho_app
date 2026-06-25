import pandas as pd
import json
from pathlib import Path
from datetime import datetime

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)

MASTER_COLUMNS = [
    "氏名", "フリガナ", "区分", "地区",
    "迎え先（平日）", "迎え先（長期休み）",
    "迎え時刻（平日）", "迎え時刻（長期休み）",
    "送り先", "送り時刻",
    "利用曜日", "住所", "備考"
]

def load_master(館: str) -> pd.DataFrame:
    p = DATA_DIR / f"{館}_master.csv"
    if p.exists():
        df = pd.read_csv(p, encoding="utf-8-sig", dtype=str).fillna("")
        # 旧カラム互換
        if "迎え先" in df.columns and "迎え先（平日）" not in df.columns:
            df["迎え先（平日）"] = df.pop("迎え先")
        if "下校時刻" in df.columns and "迎え時刻（平日）" not in df.columns:
            df["迎え時刻（平日）"] = df.pop("下校時刻")
        if "自宅時刻" in df.columns and "送り時刻" not in df.columns:
            df["送り時刻"] = df.pop("自宅時刻")
        for col in MASTER_COLUMNS:
            if col not in df.columns:
                df[col] = ""
        return df[MASTER_COLUMNS]
    return pd.DataFrame(columns=MASTER_COLUMNS)

def save_master(館: str, df: pd.DataFrame):
    p = DATA_DIR / f"{館}_master.csv"
    for col in MASTER_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df[MASTER_COLUMNS].to_csv(p, index=False, encoding="utf-8-sig")
    _save_history(館, df)

def _save_history(館: str, df: pd.DataFrame):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    hist_dir = DATA_DIR / "history" / 館
    hist_dir.mkdir(parents=True, exist_ok=True)
    df[MASTER_COLUMNS].to_csv(
        hist_dir / f"{ts}.csv", index=False, encoding="utf-8-sig")
    # 最新20件のみ保持
    files = sorted(hist_dir.glob("*.csv"))
    for old in files[:-20]:
        old.unlink()

def load_history_list(館: str):
    hist_dir = DATA_DIR / "history" / 館
    if not hist_dir.exists():
        return []
    return sorted([f.stem for f in hist_dir.glob("*.csv")], reverse=True)

def load_history(館: str, timestamp: str) -> pd.DataFrame:
    p = DATA_DIR / "history" / 館 / f"{timestamp}.csv"
    if p.exists():
        return pd.read_csv(p, encoding="utf-8-sig", dtype=str).fillna("")
    return pd.DataFrame(columns=MASTER_COLUMNS)

def import_from_ritalico(raw_df: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    result = pd.DataFrame(index=range(len(raw_df)))
    for app_col, csv_col in mapping.items():
        if csv_col and csv_col != "（未選択）" and csv_col in raw_df.columns:
            result[app_col] = raw_df[csv_col].fillna("").astype(str)
        else:
            result[app_col] = ""

    # 住所の自動結合
    if result.get("住所", pd.Series([""] * len(raw_df))).eq("").all():
        parts = []
        for c in ["都道府県", "市区町村", "番地", "ビル・マンション名"]:
            if c in raw_df.columns:
                parts.append(raw_df[c].fillna("").astype(str))
        if parts:
            result["住所"] = pd.concat(parts, axis=1).apply(
                lambda r: "".join(v for v in r if v and v != "None"), axis=1)

    # 地区自動抽出
    if "住所" in result.columns:
        result["地区"] = result["住所"].apply(_extract_district)

    # 迎え先の列名統一
    if "迎え先" in result.columns:
        result["迎え先（平日）"] = result.pop("迎え先")

    for col in MASTER_COLUMNS:
        if col not in result.columns:
            result[col] = ""
    return result[MASTER_COLUMNS].fillna("").replace("None", "")

def _extract_district(address: str) -> str:
    if not address:
        return ""
    import re
    m = re.search(r"[市区町村郡](.+?)(?:[0-9０-９\-ー－]|$)", address)
    if m:
        return m.group(1).strip()
    m = re.search(r"([\u4e00-\u9fa5ぁ-んァ-ン]+)(?=[0-9０-９])", address)
    return m.group(1) if m else ""
