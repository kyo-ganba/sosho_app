"""
送迎表Excel出力 — タイムライン縦型・午前/午後2シート構成
列順: 迎/送 | 氏名 | 送迎先 | 運転者
色設定は color_config.py から読み込む
"""
import io
from datetime import date
from typing import Dict, List, Optional
import xlsxwriter

# 列構成（車両あたり4列）
COL_TIME    = 0
COLS_PER_V  = 4   # 迎送, 氏名, 送迎先, 運転者
# オフセット
OFF_MS   = 0   # 迎/送
OFF_NAME = 1   # 氏名
OFF_DEST = 2   # 送迎先
OFF_DRV  = 3   # 運転者

ROW_TITLE  = 0
ROW_VEHDR  = 1
ROW_SUBHDR = 2
ROW_DATA   = 3

AM_BORDER = "13:00"


def export_schedule(
    routes: Dict[str, Dict[str, List[Dict]]],
    target_date: date,
    館: str,
    staff_on_duty: Dict,
    jiriki_users: List[Dict] = None,
    master_df=None,
    colors: Dict = None,
) -> bytes:
    from color_config import load_colors
    if colors is None:
        colors = load_colors(館)

    buf = io.BytesIO()
    wb  = xlsxwriter.Workbook(buf, {"in_memory": True})
    wday = ["月","火","水","木","金","土","日"][target_date.weekday()]

    all_trips  = _collect_trips(routes)
    note_map   = _build_note_map(master_df)
    kubun_map  = _build_kubun_map(master_df)
    vehicles   = list(routes.keys())
    fmt        = _make_formats(wb, colors)

    # 午前シート
    ws_am = wb.add_worksheet(
        f"午前_{target_date.month}月{target_date.day}日({wday})")
    _build_sheet(wb, ws_am, vehicles, all_trips, "午前（〜12:59）",
                 target_date, wday, 館, note_map, kubun_map, fmt, colors,
                 time_filter=lambda t: t < AM_BORDER)

    # 午後シート
    ws_pm = wb.add_worksheet(
        f"午後_{target_date.month}月{target_date.day}日({wday})")
    _build_sheet(wb, ws_pm, vehicles, all_trips, "午後・送り（13:00〜）",
                 target_date, wday, 館, note_map, kubun_map, fmt, colors,
                 time_filter=lambda t: t >= AM_BORDER)

    wb.close()
    buf.seek(0)
    return buf.read()




# ── 児発判定マップ（氏名 → 区分） ────────────────────────────
def _build_kubun_map(master_df):
    """氏名をキーに区分（放デイ/児発）を返すマップを生成"""
    kubun_map = {}
    if master_df is None:
        return kubun_map
    for _, row in master_df.iterrows():
        name = str(row.get("氏名","")).strip()
        ku   = str(row.get("区分","")).strip()
        kubun_map[name] = ku
        # (児)付きの名前でも引けるように
        kubun_map[name.replace("(児)","")] = ku
    return kubun_map


def _count_childseats(trip_map, kubun_map):
    """
    1時刻帯の便（trip_map: {車両名: trip}）から
    車両ごとのチャイルドシート必要数を返す
    Returns: {車両名: int}
    """
    result = {}
    for vehicle, trip in trip_map.items():
        if trip.get("type","") != "迎え":
            result[vehicle] = 0
            continue
        names = trip.get("name","").split("\n")
        cnt = 0
        for raw_name in names:
            name = raw_name.strip()
            ku   = kubun_map.get(name, kubun_map.get(name.replace("(児)",""), ""))
            if "児発" in ku or "(児)" in name:
                cnt += 1
        result[vehicle] = cnt
    return result

# ── シート生成 ──────────────────────────────────────────────
def _build_sheet(wb, ws, vehicles, all_trips, period_label,
                 target_date, wday, 館, note_map, kubun_map, fmt, colors,
                 time_filter):
    n_v = len(vehicles)
    total_cols = 1 + n_v * COLS_PER_V

    _set_dimensions(ws, n_v)

    # タイトル
    ws.set_row(ROW_TITLE, 22)
    ws.merge_range(ROW_TITLE, 0, ROW_TITLE, total_cols - 1,
        f"{館}　送迎表　{target_date.year}年{target_date.month}月"
        f"{target_date.day}日（{wday}）　{period_label}",
        fmt["title"])

    # 車両名ヘッダー
    ws.set_row(ROW_VEHDR, 18)
    ws.write(ROW_VEHDR, COL_TIME, "", fmt["hdr"])
    for vi, vehicle in enumerate(vehicles):
        base = 1 + vi * COLS_PER_V
        ws.merge_range(ROW_VEHDR, base, ROW_VEHDR, base + COLS_PER_V - 1,
                       vehicle, fmt["car_hdr"])

    # サブヘッダー
    ws.set_row(ROW_SUBHDR, 14)
    ws.write(ROW_SUBHDR, COL_TIME, "時刻", fmt["sub_hdr"])
    for vi in range(n_v):
        base = 1 + vi * COLS_PER_V
        ws.write(ROW_SUBHDR, base + OFF_MS,   "迎/送",  fmt["sub_hdr"])
        ws.write(ROW_SUBHDR, base + OFF_NAME, "氏名",   fmt["sub_hdr"])
        ws.write(ROW_SUBHDR, base + OFF_DEST, "送迎先", fmt["sub_hdr"])
        ws.write(ROW_SUBHDR, base + OFF_DRV,  "運転者", fmt["sub_hdr"])

    # データ行
    time_rows = _build_time_rows(vehicles, all_trips, time_filter)
    current_row = ROW_DATA
    notes_rendered = set()

    for row_def in time_rows:
        if row_def["type"] == "time":
            ws.set_row(current_row, 18)
            ws.write(current_row, COL_TIME,
                     row_def["time"], fmt["time_cell"])

            for vi, vehicle in enumerate(vehicles):
                base = 1 + vi * COLS_PER_V
                trip = row_def["trips"].get(vehicle)
                if trip:
                    _write_trip_cells(ws, fmt, current_row, base, trip)
                else:
                    _write_empty_cells(ws, fmt, current_row, base)

            current_row += 1

            # 特記事項を該当車両列の直下に表示
            for vi, vehicle in enumerate(vehicles):
                base = 1 + vi * COLS_PER_V
                trip = row_def["trips"].get(vehicle)
                if not trip:
                    continue
                raw_name = trip.get("name", "")
                clean = raw_name.replace("(児)", "").strip()
                if clean in note_map and clean not in notes_rendered:
                    ws.set_row(current_row, 24)
                    # 該当車両列に特記を表示、他は空
                    ws.write(current_row, COL_TIME, "", fmt["note_cell"])
                    for vi2, v2 in enumerate(vehicles):
                        b2 = 1 + vi2 * COLS_PER_V
                        if vi2 == vi:
                            ws.write(current_row, b2 + OFF_MS,
                                     "!", fmt["note_cell"])
                            ws.merge_range(
                                current_row, b2 + OFF_NAME,
                                current_row, b2 + COLS_PER_V - 1,
                                note_map[clean], fmt["note_cell"])
                        else:
                            for c in range(COLS_PER_V):
                                ws.write(current_row, b2 + c,
                                         "", fmt["note_cell"])
                    current_row += 1
                    notes_rendered.add(clean)

            # チャイルドシート必要数の表示（迎え便のみ・1脚以上の車両）
            cs_counts = _count_childseats(row_def["trips"], kubun_map)
            has_cs = any(v > 0 for v in cs_counts.values())
            if has_cs:
                ws.set_row(current_row, 13)
                ws.write(current_row, COL_TIME, "CS", fmt["cs_cell"])
                for vi2, v2 in enumerate(vehicles):
                    b2 = 1 + vi2 * COLS_PER_V
                    cnt = cs_counts.get(v2, 0)
                    trip2 = row_def["trips"].get(v2)
                    if cnt > 0 and trip2 and trip2.get("type","") == "迎え":
                        ws.merge_range(
                            current_row, b2,
                            current_row, b2 + COLS_PER_V - 1,
                            f"チャイルドシート {cnt}脚", fmt["cs_cell"])
                    else:
                        for c in range(COLS_PER_V):
                            ws.write(current_row, b2 + c, "", fmt["cs_empty"])
                current_row += 1

        elif row_def["type"] == "arrow":
            ws.set_row(current_row, 10)
            ws.write(current_row, COL_TIME, "↓", fmt["arrow"])
            for vi, vehicle in enumerate(vehicles):
                base = 1 + vi * COLS_PER_V
                label = row_def["labels"].get(vehicle, "")
                if label:
                    ws.merge_range(current_row, base,
                                   current_row, base + COLS_PER_V - 1,
                                   label, fmt["arrow"])
                else:
                    for c in range(COLS_PER_V):
                        ws.write(current_row, base + c, "",
                                 fmt["arrow"])
            current_row += 1

    # 印刷設定
    ws.set_landscape()
    ws.set_paper(9)
    ws.fit_to_pages(1, 0)
    ws.set_margins(left=0.5, right=0.5, top=0.6, bottom=0.6)
    ws.repeat_rows(ROW_TITLE, ROW_SUBHDR)
    ws.freeze_panes(ROW_DATA, 1)


# ── セル書き込み ─────────────────────────────────────────────
def _write_trip_cells(ws, fmt, row, base, trip):
    ms     = trip.get("type", "")
    name   = trip.get("name", "")
    place  = trip.get("place", "")
    driver = trip.get("driver", "")
    is_jis = "(児)" in name
    is_add = trip.get("is_add", False)

    if is_add:
        ws.write(row, base + OFF_MS,   "追加", fmt["ms_add"])
        ws.write(row, base + OFF_NAME, name,   fmt["name_add"])
        ws.write(row, base + OFF_DEST, place,  fmt["dest"])
        ws.write(row, base + OFF_DRV,  driver, fmt["drv_add"])
    elif ms == "迎え":
        ms_f   = fmt["ms_jis"] if is_jis else fmt["ms_mu"]
        name_f = fmt["name_jis"] if is_jis else fmt["name_mu"]
        ws.write(row, base + OFF_MS,   "迎", ms_f)
        ws.write(row, base + OFF_NAME, name,   name_f)
        ws.write(row, base + OFF_DEST, place,  fmt["dest"])
        ws.write(row, base + OFF_DRV,  driver, fmt["drv_on"])
    else:
        ws.write(row, base + OFF_MS,   "送", fmt["ms_ok"])
        ws.write(row, base + OFF_NAME, name,   fmt["name_ok"])
        ws.write(row, base + OFF_DEST, place,  fmt["dest"])
        ws.write(row, base + OFF_DRV,  driver, fmt["drv_on"])


def _write_empty_cells(ws, fmt, row, base):
    ws.write(row, base + OFF_MS,   "", fmt["empty_ms"])
    ws.write(row, base + OFF_NAME, "", fmt["empty"])
    ws.write(row, base + OFF_DEST, "", fmt["empty"])
    ws.write(row, base + OFF_DRV,  "", fmt["empty_drv"])


# ── フォーマット生成（色設定を反映）────────────────────────
def _make_formats(wb, colors):
    def f(bg, fc="#000000", bold=False, size=10, align="center",
          border=1, italic=False, wrap=True):
        d = {"font_name":"Meiryo UI","font_size":size,
             "valign":"vcenter","align":align,"text_wrap":wrap,
             "bg_color":bg,"font_color":fc,"border":border}
        if bold:   d["bold"]   = True
        if italic: d["italic"] = True
        return wb.add_format(d)

    c = colors   # 短縮
    return {
        # ヘッダー系
        "title":    f(c["タイトル"],      "#FFFFFF", bold=True, size=12),
        "hdr":      f(c["車両ヘッダー"],  "#FFFFFF", bold=True),
        "car_hdr":  f(c["車両ヘッダー"],  "#FFFFFF", bold=True, size=11),
        "sub_hdr":  f(c["サブヘッダー"],  "#004D40", size=9, bold=True),
        # 時刻
        "time_cell":f(c["時刻"],          "#00695C", bold=True, size=11),
        # 迎/送バッジ
        "ms_mu":    f(c["迎え_放デイ_バッジ"], "#B71C1C", bold=True, size=10),
        "ms_jis":   f(c["迎え_児発_バッジ"],   "#01579B", bold=True, size=10),
        "ms_ok":    f(c["送り_バッジ"],        "#1B5E20", bold=True, size=10),
        "ms_add":   f(c["急遽追加_バッジ"],    "#E65100", bold=True, size=9),
        # 氏名
        "name_mu":  f(c["迎え_放デイ_氏名"], "#212121", size=10, align="left"),
        "name_jis": f(c["迎え_児発_氏名"],   "#212121", size=10, align="left"),
        "name_ok":  f(c["送り_氏名"],        "#212121", size=10, align="left"),
        "name_add": f(c["急遽追加_氏名"],    "#E65100", size=10, align="left"),
        # 送迎先（色なし）
        "dest":     f(c["送迎先"],   "#333333", size=9, align="left"),
        # 運転者
        "drv_on":   f("#FFFFFF",     "#333333", size=9, bold=True),
        "drv_add":  f("#FFFFFF",          "#E65100", size=9),
        # 空欄
        "empty":    f("#F5F5F5",     "#DDDDDD", size=9),
        "empty_ms": f("#F5F5F5",     "#DDDDDD", size=9),
        "empty_drv":f("#F5F5F5",     "#DDDDDD", size=9),
        # 特記事項
        "note_cell":f(c["特記事項"], "#5D4037", size=9, align="left"),
        # チャイルドシート
        "cs_cell":  f("#E8EAF6", "#283593", size=9, bold=True, align="center"),
        "cs_empty": f("#F5F5F5", "#DDDDDD", size=9),
        # はしご矢印
        "arrow":    f(c["はしご矢印"], "#9E9D24", size=9, italic=True),
    }


# ── 列幅・行高 ───────────────────────────────────────────────
def _set_dimensions(ws, n_v):
    ws.set_column(COL_TIME, COL_TIME, 7)
    for vi in range(n_v):
        base = 1 + vi * COLS_PER_V
        ws.set_column(base + OFF_MS,   base + OFF_MS,   5)   # 迎/送
        ws.set_column(base + OFF_NAME, base + OFF_NAME, 12)  # 氏名
        ws.set_column(base + OFF_DEST, base + OFF_DEST, 9)   # 送迎先
        ws.set_column(base + OFF_DRV,  base + OFF_DRV,  7)   # 運転者


# ── ユーティリティ ───────────────────────────────────────────
def _collect_trips(routes):
    result = {}
    for vehicle, type_map in routes.items():
        trips = []
        for trip_type, trip_list in type_map.items():
            for t in trip_list:
                trips.append(dict(t))
        trips.sort(key=lambda x: x.get("time","00:00")[:5])
        result[vehicle] = trips
    return result


def _build_note_map(master_df):
    note_map = {}
    if master_df is not None:
        for _, row in master_df.iterrows():
            n = str(row.get("特記事項","")).strip()
            if n:
                note_map[str(row.get("氏名","")).replace("(児)","").strip()] = n
    return note_map


def _build_time_rows(vehicles, all_trips, time_filter):
    entries = []
    for vehicle in vehicles:
        for trip in all_trips.get(vehicle, []):
            t = trip.get("time","")[:5]
            if time_filter(t):
                entries.append((t, vehicle, trip))
    entries.sort(key=lambda x: x[0])

    from collections import OrderedDict
    time_dict = OrderedDict()
    for t, vehicle, trip in entries:
        if t not in time_dict:
            time_dict[t] = {}
        if vehicle in time_dict[t]:
            existing = time_dict[t][vehicle]
            existing["name"]  += f"\n{trip.get('name','')}"
        else:
            time_dict[t][vehicle] = dict(trip)

    prev_time = {v: None for v in vehicles}
    prev_type = {v: None for v in vehicles}
    rows = []

    for t, trip_map in time_dict.items():
        arrow_labels = {}
        for vehicle, trip in trip_map.items():
            if (prev_time[vehicle] is not None
                    and trip.get("type","") == "迎え"
                    and prev_type[vehicle] == "迎え"):
                arrow_labels[vehicle] = f"→ はしご"
        if any(arrow_labels.values()):
            rows.append({"type": "arrow", "labels": arrow_labels})

        rows.append({"type": "time", "time": t, "trips": trip_map})

        for vehicle in trip_map:
            prev_time[vehicle] = t
            prev_type[vehicle] = trip_map[vehicle].get("type","")

    return rows
