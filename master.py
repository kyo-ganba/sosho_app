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
    リタリコCSVをアプリ内マスタ形式に変換する。
    住所は複数列（都道府県・市区町村・番地・ビル）を自動結合。
    """
    result = pd.DataFrame(index=range(len(raw_df)))

    for app_col, csv_col in mapping.items():
        if csv_col and csv_col != "（未選択）" and csv_col in raw_df.columns:
            result[app_col] = raw_df[csv_col].fillna("").astype(str)
        else:
            result[app_col] = ""

    # ── 住所の自動結合 ────────────────────────────────────────
    # 「住所」列が未選択 or 空の場合、都道府県+市区町村+番地+ビルを結合
    addr_filled = result.get("住所", pd.Series([""] * len(raw_df)))
    if addr_filled.eq("").all():
        addr_parts = []
        for col_candidate in ["都道府県", "市区町村", "番地", "ビル・マンション名"]:
            if col_candidate in raw_df.columns:
                addr_parts.append(raw_df[col_candidate].fillna("").astype(str))
        if addr_parts:
            result["住所"] = pd.concat(addr_parts, axis=1).apply(
                lambda row: "".join(v for v in row if v and v != "None"), axis=1
            )

    # ── 地区を住所から自動抽出 ───────────────────────────────
    if "住所" in result.columns:
        result["地区"] = result["住所"].apply(_extract_district)

    # ── フリガナ未選択時は保護者カナを流用 ──────────────────
    if result.get("フリガナ", pd.Series([""])).eq("").all():
        for cand in ["保護者（カナ）", "保護者カナ", "カナ"]:
            if cand in raw_df.columns:
                result["フリガナ"] = raw_df[cand].fillna("").astype(str)
                break

    # ── 区分が未選択の場合は空欄（手動入力） ────────────────
    # 後でマスタ編集画面から放デイ/児発を手入力してもらう

    # 空文字に統一
    for col in MASTER_COLUMNS:
        if col not in result.columns:
            result[col] = ""
    result = result.fillna("").replace("None", "")

    return result[MASTER_COLUMNS]


def _extract_district(address: str) -> str:
    """住所から町丁目レベルの地区名を抽出する"""
    if not address:
        return ""
    import re
    # 「松戸市小金原」→「小金原」のように市区町村以降の地名を取る
    match = re.search(r"[市区町村郡](.+?)(?:[0-9０-９\-ー－]|$)", address)
    if match:
        return match.group(1).strip()
    # フォールバック: 末尾の漢字地名
    match = re.search(r"([\u4e00-\u9fa5ぁ-んァ-ン]+)(?=[0-9０-９])", address)
    if match:
        return match.group(1)
    return ""
