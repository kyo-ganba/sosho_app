"""
送迎ルート自動生成ロジック
迎え先（平日）/（長期休み）を区別し、時刻+地区でグループ化して車両割当
"""
from datetime import time
from typing import List, Dict
import pandas as pd

TIME_GAP = 20   # 同一便にまとめる時間差（分）

def generate_routes(participants: pd.DataFrame, vehicles: List[Dict],
                    drivers: List[str], is_long_holiday: bool = False) -> Dict:
    if participants.empty or not vehicles:
        return {}

    mukae_col = "迎え先（長期休み）" if is_long_holiday else "迎え先（平日）"
    time_col  = "迎え時刻（長期休み）" if is_long_holiday else "迎え時刻（平日）"

    mukae = _build_trips(participants, "迎え", mukae_col, time_col, "地区")
    okuri = _build_trips(participants, "送り", "送り先", "送り時刻", "地区")
    all_trips = mukae + okuri

    return _assign(all_trips, vehicles, drivers)

def _build_trips(df, trip_type, place_col, time_col, district_col):
    rows = []
    for _, r in df.iterrows():
        place = r.get(place_col, "") or r.get("住所", "自宅") or "自宅"
        t = _parse(r.get(time_col, "15:00"))
        rows.append({
            "type": trip_type,
            "time_obj": t,
            "time": t.strftime("%H:%M"),
            "name": r.get("氏名", ""),
            "place": place,
            "district": r.get(district_col, ""),
            "driver": "",
            "vehicle": "",
        })
    rows.sort(key=lambda x: (x["district"], x["time_obj"]))
    return rows

def _assign(trips, vehicles, drivers):
    v_names = [v["車両名"] for v in vehicles]
    caps    = {v["車両名"]: int(v.get("定員", 6)) - 1 for v in vehicles}
    load    = {v: 0 for v in v_names}
    dist_v  = {}
    route   = {v: [] for v in v_names}

    for trip in trips:
        d = trip["district"]
        chosen = None
        if d and d in dist_v and load[dist_v[d]] < caps[dist_v[d]]:
            chosen = dist_v[d]
        if chosen is None:
            for v in v_names:
                if load[v] < caps[v]:
                    chosen = v
                    break
        if chosen is None:
            chosen = v_names[0]

        t = trip.copy()
        if drivers:
            t["driver"] = drivers[v_names.index(chosen) % len(drivers)]
        t["vehicle"] = chosen
        route[chosen].append(t)
        if d:
            dist_v[d] = chosen
        load[chosen] += 1

    return {k: v for k, v in route.items() if v}

def _parse(val) -> time:
    if isinstance(val, time):
        return val
    if isinstance(val, str) and ":" in val:
        try:
            parts = val.strip().split(":")
            return time(int(parts[0]), int(parts[1][:2]))
        except Exception:
            pass
    return time(15, 0)
