"""
送迎表Excel出力 — xlsxwriter版
時刻を HH:MM 形式で出力、列幅・行高を元フォーマットに合わせて統一
"""
import io
from datetime import date
from typing import Dict, List
import xlsxwriter
import xlsxwriter.utility as xu

PINK       = "#FF7C80"
GREEN_C    = "#00CC99"
YELLOW     = "#FFFF00"
LIGHT_BLUE = "#D9E1F2"
ORANGE     = "#FFD966"
YELLOW_JIS = "#FFCC00"

VEHICLE_COLS = {
    "セレナ":    (1,  2,  3,  4,  5),
    "ボクシー":  (9,  10, 11, 12, 13),
    "白フリード":(14, 15, 16, 17, 18),
    "銀フリ地下":(20, 21, 22, 23, 24),
    "銀フリ浅香":(36, 37, 38, 39, 40),
}
VEHICLE_NAME_CELLS = {
    "セレナ":    (3,  4),
    "ボクシー":  (10, 11),
    "白フリード":(16, 17),
    "銀フリ地下":(22, 23),
    "銀フリ浅香":(28, 29),
}
STAFF_COLS  = [3, 10, 16, 22, 30]
COUNT_COLS  = {"セレナ":3,"ボクシー":11,"白フリード":16,"銀フリ地下":22,"銀フリ浅香":38}
JIRIKI_TIME = 49
JIRIKI_NAME = 50
JIRIKI_NOTE = 51

ROW_DATE = 1; ROW_WTIME = 2; ROW_COUNT = 3; ROW_HDR = 4; ROW_DATA = 5

def export_schedule(routes, target_date, 館, staff_on_duty, jiriki_users=None):
    buf = io.BytesIO()
    wb  = xlsxwriter.Workbook(buf, {"in_memory": True})
    wday = ["月","火","水","木","金","土","日"][target_date.weekday()]
    ws  = wb.add_worksheet(f"{target_date.month}月{target_date.day}日({wday})")
    fmt = _fmts(wb)
    _dims(ws)
    _header(ws, fmt, target_date, wday, 館, staff_on_duty)
    _veh_headers(ws, fmt)
    _counts(ws, fmt)
    _trips(ws, fmt, routes)
    _jiriki(ws, fmt, jiriki_users or [])
    wb.close()
    buf.seek(0)
    return buf.read()

def _fmts(wb):
    def f(**k):
        b={"font_name":"Meiryo UI","font_size":11,"valign":"vcenter",
           "align":"center","text_wrap":True}; b.update(k); return wb.add_format(b)
    def fb(**k):
        b={"font_name":"Meiryo UI","font_size":11,"valign":"vcenter",
           "align":"center","text_wrap":True,"border":1}; b.update(k); return wb.add_format(b)
    return {
        "normal":f(),"bold12":f(bold=True,font_size=12),
        "bold10":f(bold=True,font_size=10),"small":f(font_size=9),
        "mukae_hd":fb(bg_color=PINK),"okuri_hd":fb(bg_color=GREEN_C),
        "mukae_jis":fb(bg_color=LIGHT_BLUE),"okuri_jis":fb(bg_color=ORANGE),
        "name_hd":fb(bg_color=YELLOW),"name_jis":fb(bg_color=YELLOW_JIS),
        "time_f":fb(font_size=11),"place_f":fb(align="left",font_size=10),
        "driver_f":fb(font_size=10),"count_f":f(font_size=9),
        "jiriki_n":f(font_size=10,align="left"),
    }

def _header(ws, fmt, target_date, wday, 館, staff):
    ws.merge_range(ROW_DATE,0,ROW_DATE,2,
        f"{target_date.year}年{target_date.month}月{target_date.day}日（{wday}）",fmt["bold12"])
    on = [(n,i) for n,i in staff.items() if i.get("on")]
    for idx,(name,info) in enumerate(on):
        if idx>=len(STAFF_COLS): break
        col=STAFF_COLS[idx]
        ws.merge_range(ROW_DATE, col,ROW_DATE, col+1,name,fmt["normal"])
        ws.merge_range(ROW_WTIME,col,ROW_WTIME,col+1,info.get("work_time","9:30-18:30"),fmt["small"])

def _veh_headers(ws, fmt):
    for v,(c1,c2) in VEHICLE_NAME_CELLS.items():
        ws.merge_range(ROW_HDR,c1,ROW_HDR,c2,v,fmt["bold12"])
        cols=VEHICLE_COLS.get(v)
        if cols:
            ws.write(ROW_HDR,cols[4],"運転",fmt["small"])
            ws.write(ROW_HDR,cols[4]+1,"LINE",fmt["small"])
    ws.write(ROW_HDR,JIRIKI_NAME,"自力送迎",fmt["bold10"])

def _counts(ws, fmt):
    for v,col in COUNT_COLS.items():
        cl=xu.xl_col_to_name(col)
        ws.write_formula(ROW_COUNT,col,
            f'="チャイルド"&COUNTIF({cl}7:{cl}51,"*(児)")',fmt["count_f"])

def _trips(ws, fmt, routes):
    cur={v:ROW_DATA for v in VEHICLE_COLS}
    all_t=[]
    for v,tm in routes.items():
        for tt,trips in tm.items():
            for t in trips: all_t.append((v,t))
    all_t.sort(key=lambda x:x[1].get("time","00:00"))

    for v,trip in all_t:
        if v not in VEHICLE_COLS: continue
        cols=VEHICLE_COLS[v]; r=cur[v]
        ms=trip.get("type",""); name=trip.get("name","")
        place=trip.get("place","自宅"); driver=trip.get("driver","")
        # 時刻を HH:MM に正規化
        raw_t = trip.get("time","15:00")
        t_str = raw_t[:5] if len(raw_t)>=5 else raw_t
        is_jis="(児)" in name
        if ms=="迎え":
            mf=fmt["mukae_jis"] if is_jis else fmt["mukae_hd"]
            nf=fmt["name_jis"]  if is_jis else fmt["name_hd"]
            ml="迎"
        else:
            mf=fmt["okuri_jis"] if is_jis else fmt["okuri_hd"]
            nf=fmt["name_jis"]  if is_jis else fmt["name_hd"]
            ml="送"
        ws.write(r,cols[0],t_str,  fmt["time_f"])
        ws.write(r,cols[1],ml,     mf)
        ws.write(r,cols[2],name,   nf)
        ws.write(r,cols[3],place,  fmt["place_f"])
        ws.write(r,cols[4],driver, fmt["driver_f"])
        cur[v]+=1

    lr=max(cur.values())+2
    ws.write(lr,3,"リーダー",fmt["bold10"])
    ws.write(lr,4,"児発：",  fmt["normal"])
    ws.write(lr,9,"放デイ：",fmt["normal"])

def _jiriki(ws, fmt, users):
    r=ROW_DATA
    for u in users:
        ws.write(r,JIRIKI_TIME,u.get("time",""),    fmt["normal"])
        ws.write(r,JIRIKI_NAME,u.get("name",""),    fmt["jiriki_n"])
        ws.write(r,JIRIKI_NOTE,u.get("note","自力"),fmt["small"])
        r+=1

def _dims(ws):
    W=[22.4,7.5,6.0,13.0,9.7,8.7,8.9,13.0,7.6,4.9,
       13.0,10.2,8.9,8.0,7.6,5.0,13.0,9.0,8.9,8.6,
       7.6,5.0,13.0,9.0,8.9,8.0,8.0,8.0,8.0,8.0]
    for i,w in enumerate(W): ws.set_column(i,i,w)
    for i in range(30,55):   ws.set_column(i,i,8.5)
    ws.set_row(0,45.75); ws.set_row(1,48.75); ws.set_row(2,41.25)
    ws.set_row(3,31.5);  ws.set_row(4,30.0)
    for r in range(5,55): ws.set_row(r,22.0)   # データ行は22ptに統一
