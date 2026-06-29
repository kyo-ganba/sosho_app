"""
送迎ルート自動生成ロジック
地区ベースのgreedy割当 + 時刻ソート
"""
import pandas as pd
from typing import List, Dict, Any

TIME_GAP_MINUTES  = 20       # 同一グループとみなす時間幅（将来の時間×距離チェック用）
HOME_TIME_DEFAULT = "15:00"  # 送りのデフォルト出発時刻


# ── ユーティリティ ────────────────────────────────────────────
def _parse_time(s: Any, default: str = HOME_TIME_DEFAULT) -> str:
    """文字列を HH:MM 形式にノーマライズ"""
    try:
        parts = str(s).strip().split(":")
        return f"{int(parts[0]):02d}:{int(parts[1][:2]):02d}"
    except Exception:
        return default


def _time_to_min(t: str) -> int:
    try:
        h, m = t.split(":")
        return int(h) * 60 + int(m)
    except Exception:
        return _time_to_min(HOME_TIME_DEFAULT)


# ── メイン関数 ────────────────────────────────────────────────
def generate_routes(
    parts: pd.DataFrame,
    vehicles: List[Dict],
    drivers: List[str],
    is_long_holiday: bool = False,
) -> Dict[str, List[Dict]]:
    """
    参加者リストから送迎ルートを生成する。

    Parameters:
        parts           : 参加者DataFrame（マスタの部分集合）
        vehicles        : 車両リスト [{"車両名": str, "定員": int}, ...]
        drivers         : 運転者名リスト
        is_long_holiday : 長期休みフラグ（迎え先・時刻列を切り替え）

    Returns:
        {車両名: [{"type":"迎え"|"送り", "time":"HH:MM",
                   "name": str, "place": str, "driver": str}, ...]}
    """
    place_col = "迎え先（長期休み）" if is_long_holiday else "迎え先（平日）"
    time_col  = "迎え時刻（長期休み）" if is_long_holiday else "迎え時刻（平日）"

    vehicle_names = [v["車両名"] for v in vehicles]
    # ドライバー1名分を引いた乗客定員
    capacity = {v["車両名"]: max(1, int(v.get("定員", 7)) - 1) for v in vehicles}

    driver_pool = list(drivers) if drivers else []
    _di = [0]

    def next_driver() -> str:
        if not driver_pool:
            return ""
        d = driver_pool[_di[0] % len(driver_pool)]
        _di[0] += 1
        return d

    result: Dict[str, List[Dict]] = {v: [] for v in vehicle_names}

    def _assign(rows: List[Dict], trip_type: str):
        """
        地区ベースのgreedy割当。
        同地区の利用者は同じ車両に乗せる。定員を超えたら次の車両へ。
        """
        load:   Dict[str, int] = {v: 0 for v in vehicle_names}
        dist_v: Dict[str, str] = {}   # district -> 割当済み車両名

        for row in rows:
            dist = row["district"]

            # 同地区の車両があり、かつ空きがある場合は流用
            v = dist_v.get(dist)
            if v is None or load[v] >= capacity[v]:
                # 空き車両（まだ誰も乗っていない）を優先して探す
                v = next((vn for vn in vehicle_names if load[vn] == 0), None)
                if v is None:
                    # すべて満員 → 最も余裕のある車両（溢れても割り当てる）
                    v = min(vehicle_names, key=lambda vn: load[vn])
                dist_v[dist] = v

            load[v] += 1
            result[v].append({
                "type":   trip_type,
                "time":   row["time"],
                "name":   row["name"],
                "place":  row["place"],
                "driver": next_driver(),
            })

    # ── 迎え ──────────────────────────────────────────────────
    pick: List[Dict] = []
    for _, row in parts.iterrows():
        name  = str(row.get("氏名", "")).strip()
        ku    = str(row.get("区分", "")).strip()
        place = (str(row.get(place_col, "") or row.get("住所", "") or "自宅").strip()
                 or "自宅")
        t_str = _parse_time(row.get(time_col, HOME_TIME_DEFAULT))
        dist  = str(row.get("地区", "")).strip()
        if "児発" in ku:
            name += "(児)"
        pick.append({
            "name": name, "place": place, "time": t_str,
            "district": dist, "time_min": _time_to_min(t_str),
        })

    pick.sort(key=lambda x: x["time_min"])
    _assign(pick, "迎え")

    # ── 送り ──────────────────────────────────────────────────
    send: List[Dict] = []
    for _, row in parts.iterrows():
        name  = str(row.get("氏名", "")).strip()
        ku    = str(row.get("区分", "")).strip()
        place = (str(row.get("送り先", "") or row.get("住所", "") or "自宅").strip()
                 or "自宅")
        t_str = _parse_time(row.get("送り時刻", HOME_TIME_DEFAULT),
                             default=HOME_TIME_DEFAULT)
        dist  = str(row.get("地区", "")).strip()
        if "児発" in ku:
            name += "(児)"
        send.append({
            "name": name, "place": place, "time": t_str,
            "district": dist, "time_min": _time_to_min(t_str),
        })

    send.sort(key=lambda x: x["time_min"])
    _assign(send, "送り")

    # 空の車両を除外して返す
    return {v: trips for v, trips in result.items() if trips}
