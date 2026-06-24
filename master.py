import pandas as pd
from pathlib import Path

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)

MASTER_COLUMNS = [
    "氏名", "フリガナ", "区分", "地区", "住所",
    "迎え先", "下校時刻", "自宅時刻", "利用曜日", "備考"
]

def load_master(館: str) -> pd.DataFrame:
    p = DATA_DIR / f"{館}_master.csv"
    if p.exists():
        return pd.read_csv(p, encoding="utf-8-sig", dtype=str).fillna("")
    return pd.DataFrame(columns=MASTER_COLUMNS)

def save_master(館: str, df: pd.DataFrame):
    p = DATA_DIR / f"{館}_master.csv"
    for col in MASTER_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df[MASTER_COLUMNS].to_csv(p, index=False, encoding="utf-8-sig")

def import_from_ritalico(raw_df: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    """
    リタリコCSVの列をアプリ内項目にマッピングして変換する。
    mapping = {"氏名": "csv列名", "住所": "csv列名", ...}
    """
    result = pd.DataFrame(columns=MASTER_COLUMNS)

    for app_col, csv_col in mapping.items():
        if csv_col and csv_col != "（未選択）" and csv_col in raw_df.columns:
            result[app_col] = raw_df[csv_col].fillna("").astype(str)

    # 住所から地区を自動推定（町名を地区コードとして使う）
    if "住所" in result.columns and result["住所"].any():
        result["地区"] = result["住所"].apply(_extract_district)

    # 学校名を迎え先のデフォルトとして設定
    if "学校名" in raw_df.columns and "迎え先" not in result.columns:
        result["迎え先"] = raw_df.get("学校名", "").fillna("").astype(str)

    return result.fillna("")

def _extract_district(address: str) -> str:
    """住所から町丁目レベルの地区名を抽出する簡易ロジック"""
    if not address:
        return ""
    # 市区町村以降の最初のトークンを地区とする
    # 例: "千葉県松戸市小金原1-2-3" → "小金原"
    import re
    match = re.search(r"[市区町村郡](.+?)[0-9０-９\-ー－]", address)
    if match:
        return match.group(1).strip()
    # フォールバック: 数字前の文字列
    match = re.search(r"[\u4e00-\u9fa5ぁ-んァ-ン]+(?=[0-9０-９])", address)
    if match:
        return match.group(0)
    return ""
