"""
送迎ルート自動生成ロジック

方針:
    1. 迎え: 下校時刻でソートし、時刻が近い + 地区が同じ利用者を同じ車にまとめる
    2. 送り: 自宅時刻（または事業所終了時刻）でソートし、地区でグループ化
    3. 車両定員と運転者数の制約を考慮して振り分ける
"""

from datetime import time, datetime, timedelta
import pandas as pd
from typing import List, Dict, Any


TIME_GAP_MINUTES = 20   # この時間差以内なら同一便候補
HOME_TIME_DEFAULT = "17:00"  # 送りのデフォルト出発時刻


def generate_routes(
    participants: pd.DataFrame,
    vehicles: List[Dict],
    drivers: List[str],
) -> Dict[str, List[Dict]]:
    """
    Returns:
        {
            "セレナ": [
                {"type": "迎え", "time": "14:30", "name": "山田太郎", "place": "松戸小", "driver": "田中"},
                {"type": "送り", "time": "17:00", "name": "山田太郎", "place": "自宅", "driver": "田中"},
                ...
            ],
            ...
        }
    """
    if participants.empty or not vehicles:
        return {}

    # ── 迎えルートを生成 ──────────────────────────────────────
    mukae_trips = _build_pickup_trips(participants)

    # ── 送りルートを生成 ──────────────────────────────────────
    okuri_trips = _build_dropoff_trips(participants)

    all_trips = mukae_trips + okuri_trips

    # ── 車両・運転者に振り分け ────────────────────────────────
    routes = _assign_to_vehicles(all_trips, vehicles, drivers)

    return routes


# ─────────────────────────────────────────────────────────────
def _build_pickup_trips(df: pd.DataFrame) -> List[Dict]:
    """下校時刻順に迎え便を組む"""
    rows = []
    for _, r in df.iterrows():
        t = _parse_time(r.get("下校時刻", "15:00"))
        place = r.get("迎え先", "") or "自宅"
        rows.append({
            "type": "迎え",
            "time_obj": t,
            "time": t.strftime("%H:%M"),
            "name": r.get("氏名", ""),
            "place": place,
            "district": r.get("地区", ""),
            "driver": "",
            "vehicle": "",
        })
    rows.sort(key=lambda x: x["time_obj"])
    return rows


def _build_dropoff_trips(df: pd.DataFrame) -> List[Dict]:
    """送り便を組む（地区でグループ化して時刻順）"""
    rows = []
    for _, r in df.iterrows():
        t = _parse_time(r.get("自宅時刻", HOME_TIME_DEFAULT))
        rows.append({
            "type": "送り",
            "time_obj": t,
            "time": t.strftime("%H:%M"),
            "name": r.get("氏名", ""),
            "place": r.get("住所", "自宅") or "自宅",
            "district": r.get("地区", ""),
            "driver": "",
            "vehicle": "",
        })
    rows.sort(key=lambda x: (x["district"], x["time_obj"]))
    return rows


def _assign_to_vehicles(
    trips: List[Dict],
    vehicles: List[Dict],
    drivers: List[str],
) -> Dict[str, List[Dict]]:
    """
    地区ベースのグリーディ割当:
      - 同地区の連続する迎え/送りを同一車にまとめる
      - 定員を超えたら次の車に移す
      - 運転者を順番に割り当てる
    """
    if not vehicles:
        return {}

    # 車両ごとのバケット
    route_map: Dict[str, List[Dict]] = {v["車両名"]: [] for v in vehicles}
    capacities: Dict[str, int] = {v["車両名"]: int(v.get("定員", 6)) - 1 for v in vehicles}  # -1 for driver
    load: Dict[str, int] = {v["車両名"]: 0 for v in vehicles}

    vehicle_names = [v["車両名"] for v in vehicles]
    driver_idx = 0

    # 地区→担当車のキャッシュ（同地区は同じ車にまとめる）
    district_vehicle: Dict[str, str] = {}

    for trip in trips:
        district = trip["district"]
        v_name = None

        # 同地区の既割当車があればそちらを優先
        if district and district in district_vehicle:
            candidate = district_vehicle[district]
            if load[candidate] < capacities[candidate]:
                v_name = candidate

        # なければ空きのある最初の車を選ぶ
        if v_name is None:
            for vn in vehicle_names:
                if load[vn] < capacities[vn]:
                    v_name = vn
                    break

        # すべて満席なら先頭車両に強制
        if v_name is None:
            v_name = vehicle_names[0]

        # 運転者を割当（迎えの最初の便のみ）
        trip = trip.copy()
        if drivers:
            trip["driver"] = drivers[driver_idx % len(drivers)]
            # 次の便は同じ運転者が担当（1車1人制）
            # ※ここでは簡易的に車ごとに固定
            driver_idx_for_vehicle = vehicle_names.index(v_name) % len(drivers)
            trip["driver"] = drivers[driver_idx_for_vehicle]

        trip["vehicle"] = v_name
        route_map[v_name].append(trip)

        if district:
            district_vehicle[district] = v_name

        load[v_name] += 1

    # 不要な空車両を除く
    return {k: v for k, v in route_map.items() if v}


# ─── ユーティリティ ──────────────────────────────────────────
def _parse_time(val) -> time:
    if isinstance(val, time):
        return val
    if isinstance(val, str) and ":" in val:
        try:
            h, m = val.strip().split(":")
            return time(int(h), int(m[:2]))
        except Exception:
            pass
    return time(15, 0)
