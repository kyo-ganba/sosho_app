import re
import pandas as pd
from pathlib import Path
from datetime import datetime

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)

MASTER_COLUMNS = [
    "受給者証番号",
    "氏名", "フリガナ", "区分", "地区",
    "医ケア", "重心",
    "迎え先（平日）", "迎え先（長期休み）",
    "迎え時刻（平日）", "迎え時刻（長期休み）",
    "送り先", "送り時刻",
    "利用曜日", "通所区分", "住所",
    "迎え先住所", "送り先住所",
    "契約上限", "契約月",
    "利用開始日", "利用終了日", "状態",
    "特記事項", "備考",
]

HALLS     = ["Ⅰ番館", "Ⅱ番館", "Ⅲ番館", "Ⅴ番館"]
DAY_COLS  = ["月_固定", "火_固定", "水_固定", "木_固定", "金_固定", "土_固定"]
DAY_NAMES = ["月", "火", "水", "木", "金", "土"]


def load_master(館: str) -> pd.DataFrame:
    p = DATA_DIR / f"{館}_master.csv"
    if p.exists():
        df = pd.read_csv(p, encoding="utf-8-sig", dtype=str).fillna("")
        rename = {
            "迎え先":   "迎え先（平日）",
            "下校時刻": "迎え時刻（平日）",
            "自宅時刻": "送り時刻",
        }
        for old, new in rename.items():
            if old in df.columns and new not in df.columns:
                df[new] = df.pop(old)
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
    df[MASTER_COLUMNS].to_csv(hist_dir / f"{ts}.csv", index=False, encoding="utf-8-sig")
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


def normalize_hall(v: str) -> str:
    v = str(v).strip()
    tbl = {
        "1": "Ⅰ番館", "Ⅰ": "Ⅰ番館",
        "2": "Ⅱ番館", "Ⅱ": "Ⅱ番館",
        "3": "Ⅲ番館", "Ⅲ": "Ⅲ番館",
        "5": "Ⅴ番館", "Ⅴ": "Ⅴ番館",
    }
    return tbl.get(v, v)


def normalize_kubun(v: str) -> str:
    v = str(v).strip()
    if v in ("放デイ", "放課後等デイサービス"):
        return "放デイ"
    if v in ("児発", "児童発達支援"):
        return "児発"
    return v


def normalize_place(v: str) -> str:
    v = str(v).strip()
    return "なし" if v in ("", "なし", "送迎なし", "None") else v


def is_temp_juki_no(v: str) -> bool:
    v = str(v).strip()
    if not v:
        return True
    if v.startswith("申請") or v in ("申請中", "申請予定", "申請済み"):
        return True
    digits = v.lstrip("0")
    if not digits:
        return True
    try:
        if int(digits) < 1000:
            return True
    except ValueError:
        return True
    return False


def _extract_district(address: str) -> str:
    if not address:
        return ""
    m = re.search(r"[市区町村郡](.+?)(?:[0-9０-９\-ー－]|$)", address)
    if m:
        return m.group(1).strip()
    m = re.search(r"([一-龥ぁ-んァ-ン]+)(?=[0-9０-９])", address)
    return m.group(1) if m else ""


def import_from_internal_csv(raw_df: pd.DataFrame) -> dict:
    df = raw_df.copy().fillna("").astype(str)
    for col in DAY_COLS:
        if col in df.columns:
            df[col] = df[col].str.strip().apply(normalize_hall)

    results = {}
    for hall in HALLS:
        mask = pd.Series(False, index=df.index)
        for col in DAY_COLS:
            if col in df.columns:
                mask |= (df[col] == hall)
        if not mask.any():
            results[hall] = pd.DataFrame(columns=MASTER_COLUMNS)
            continue

        filtered = df[mask].copy()

        def _weekdays(row, _h=hall):
            return "".join(
                d for d, c in zip(DAY_NAMES, DAY_COLS)
                if c in row.index and row[c] == _h
            )

        def _get(col_name):
            if col_name in filtered.columns:
                return filtered[col_name]
            return pd.Series("", index=filtered.index)

        out = pd.DataFrame(index=filtered.index)
        out["受給者証番号"]         = _get("受給者証番号").str.strip()
        out["氏名"]                 = _get("氏名")
        out["フリガナ"]             = _get("カナ")
        out["区分"]                 = _get("区分").apply(normalize_kubun)
        out["地区"]                 = ""
        out["医ケア"]               = _get("医ケア")
        out["重心"]                 = _get("重心")
        out["迎え先（平日）"]       = _get("迎え").apply(normalize_place)
        out["迎え先（長期休み）"]   = out["迎え先（平日）"]
        out["迎え時刻（平日）"]     = ""
        out["迎え時刻（長期休み）"] = ""
        out["送り先"]               = _get("送り").apply(normalize_place)
        out["送り時刻"]             = ""
        out["利用曜日"]             = filtered.apply(_weekdays, axis=1)
        out["通所区分"]             = _get("通所区分")
        out["住所"]                 = ""
        out["迎え先住所"]           = ""
        out["送り先住所"]           = ""
        out["契約上限"]             = _get("契約上限")
        out["契約月"]               = _get("契約月")
        out["利用開始日"]           = _get("利用開始日")
        out["利用終了日"]           = _get("利用終了日")
        out["状態"]                 = _get("状態")
        out["特記事項"]             = ""
        out["備考"]                 = _get("備考")

        for col in MASTER_COLUMNS:
            if col not in out.columns:
                out[col] = ""

        results[hall] = (
            out[MASTER_COLUMNS].fillna("").replace("None", "").reset_index(drop=True)
        )

    return results


def get_facilities_needing_address(master_df: pd.DataFrame) -> list:
    skip = {"", "自宅", "なし", "送迎なし", "自力", "希望"}
    facilities = set()
    for col in ("迎え先（平日）", "迎え先（長期休み）", "送り先"):
        if col not in master_df.columns:
            continue
        for v in master_df[col].astype(str):
            v = v.strip()
            if v and v not in skip:
                facilities.add(v)
    return sorted(facilities)


def lookup_address_google(place_name: str, api_key: str,
                          city_hint: str = "松戸市") -> list:
    try:
        import requests
        r = requests.get(
            "https://maps.googleapis.com/maps/api/place/findplacefromtext/json",
            params={
                "input": f"{place_name} {city_hint}",
                "inputtype": "textquery",
                "fields": "formatted_address,name",
                "key": api_key,
                "language": "ja",
            },
            timeout=8,
        )
        return r.json().get("candidates", [])
    except Exception:
        return []


def import_from_ritalico(raw_df: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    result = pd.DataFrame(index=range(len(raw_df)))
    for app_col, csv_col in mapping.items():
        if csv_col and csv_col != "（未選択）" and csv_col in raw_df.columns:
            result[app_col] = raw_df[csv_col].fillna("").astype(str)
        else:
            result[app_col] = ""

    if result.get("住所", pd.Series([""] * len(raw_df))).eq("").all():
        parts = []
        for c in ("都道府県", "市区町村", "番地", "ビル・マンション名"):
            if c in raw_df.columns:
                parts.append(raw_df[c].fillna("").astype(str))
        if parts:
            result["住所"] = pd.concat(parts, axis=1).apply(
                lambda r: "".join(v for v in r if v and v not in ("None", "nan")), axis=1)

    if "住所" in result.columns:
        result["地区"] = result["住所"].apply(_extract_district)
    if "迎え先" in result.columns and "迎え先（平日）" not in result.columns:
        result.rename(columns={"迎え先": "迎え先（平日）"}, inplace=True)

    for col in MASTER_COLUMNS:
        if col not in result.columns:
            result[col] = ""
    return result[MASTER_COLUMNS].fillna("").replace("None", "")


def import_address_from_hogosha_csv(hogosha_df: pd.DataFrame,
                                    master_df: pd.DataFrame) -> tuple:
    """
    保護者一覧CSVから自宅住所を取込み、master_df の「住所」列を更新する。
    児童列が「青木 佑都、青木 佑愛」のように複数名を含む場合も対応。
    Returns: (updated_master_df, matched_count, unmatched_names)
    """
    hdf = hogosha_df.copy().fillna("")

    def _build_addr(row):
        parts = [
            str(row.get("都道府県", "") or "").strip(),
            str(row.get("市区町村", "") or "").strip(),
            str(row.get("番地",     "") or "").strip(),
        ]
        bld = str(row.get("ビル・マンション名", "") or "").strip()
        if bld and bld.lower() not in ("nan", "none", ""):
            parts.append(bld)
        return "".join(p for p in parts if p)

    name_to_addr: dict = {}
    for _, hrow in hdf.iterrows():
        addr = _build_addr(hrow)
        if not addr:
            continue
        child_raw = str(hrow.get("児童", "") or "").strip()
        child_raw = re.sub(r"[,，・]", "、", child_raw)
        for cn in child_raw.split("、"):
            cn = cn.strip()
            if cn:
                name_to_addr[cn] = addr

    def _normalize(s: str) -> str:
        return str(s).strip().replace(" ", "").replace("　", "")

    lookup = {_normalize(k): v for k, v in name_to_addr.items()}

    updated = master_df.copy()
    if "住所" not in updated.columns:
        updated["住所"] = ""

    matched = 0
    unmatched_names = []
    for idx, row in updated.iterrows():
        key = _normalize(str(row.get("氏名", "")))
        if key and key in lookup:
            updated.at[idx, "住所"] = lookup[key]
            district = _extract_district(lookup[key])
            if district and "地区" in updated.columns:
                if not str(updated.at[idx, "地区"]).strip():
                    updated.at[idx, "地区"] = district
            matched += 1
        else:
            if key:
                unmatched_names.append(str(row.get("氏名", "")))

    return updated, matched, unmatched_names
