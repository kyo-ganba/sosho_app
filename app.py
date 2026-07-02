"""
送迎表作成ツール v1.0
- 社内スタッフのみアクセス可能（パスワード認証）
- 保護者一覧CSVから自宅住所を一括取込 → 迎え先/送り先「自宅」に自動紐づけ
- 自力通所者の登所・退所時刻管理
- 時刻入力: プルダウン（5分刻み）＋手入力 両対応
- スペース無視のあいまい名前検索（全画面対応）
- スタッフ勤務: A/B/C勤・半休・有休・カスタム勤務時間対応
"""
import re
import streamlit as st
import pandas as pd
import json
from pathlib import Path
from datetime import datetime, time, date

# set_page_config must be the first Streamlit call
# (before any module that uses @st.cache_resource at import time)
st.set_page_config(page_title="送迎表ツール", page_icon="🚐", layout="wide")

from color_config import (load_colors, save_colors, reset_colors,
                          COLOR_LABELS, COLOR_GROUPS, DEFAULT_COLORS)
from master import (load_master, save_master, import_from_ritalico,
                    import_from_internal_csv, is_temp_juki_no,
                    get_facilities_needing_address, lookup_address_google,
                    import_address_from_hogosha_csv,
                    load_history_list, load_history, MASTER_COLUMNS)
from routing import generate_routes
from excel_export import export_schedule
from storage import load_json_data, save_json_data, is_gsheet_configured

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)

WEEKDAYS = ["月","火","水","木","金","土","日"]
HALLS    = ["Ⅰ番館","Ⅱ番館","Ⅲ番館","Ⅴ番館"]

SHIFT_PRESETS = {
    "A勤":      ("09:00", "18:00"),
    "B勤":      ("09:30", "18:30"),
    "C勤":      ("10:00", "19:00"),
    "半休AM":   ("09:00", "13:00"),
    "半休PM":   ("13:00", "18:30"),
    "カスタム": None,
    "有休":     None,
    "欠勤":     None,
}
_OFF_SHIFTS = {"有休", "欠勤"}

_QUICK_TIMES = [f"{h:02d}:{m:02d}"
                for h in range(6, 20)
                for m in range(0, 60, 5)]

_JIRIKI_VALUES = {"自力", "なし", "送迎なし", "自力通所"}


def _fuzzy_match(query: str, *texts: str) -> bool:
    """スペース・全角スペースを無視した部分一致検索。複数テキストのいずれかに含まれれば True。"""
    q = re.sub(r'[\s　]+', '', str(query)).lower()
    if not q:
        return True
    for text in texts:
        t = re.sub(r'[\s　]+', '', str(text)).lower()
        if q in t:
            return True
    return False


# ════ 認証 ════════════════════════════════════════════════════
def _check_auth() -> bool:
    """社内スタッフ向けパスワード認証。認証済みなら True を返す。"""
    if st.session_state.get("authenticated"):
        return True

    try:
        valid_pws = list(st.secrets.get("auth", {}).get("passwords", []))
    except Exception:
        valid_pws = []
    if not valid_pws:
        valid_pws = ["sosho2024"]   # ← Streamlit Secrets で必ず上書きしてください

    st.title("🚐 送迎表ツール")
    st.markdown("### キッズフロンティア 社内専用システム")
    st.divider()
    st.markdown("**🔐 スタッフ認証**")
    pw = st.text_input("パスワードを入力してください", type="password", key="login_pw",
                       placeholder="パスワード")
    if st.button("ログイン →", type="primary", use_container_width=False):
        if pw in valid_pws:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("❌ パスワードが違います。もう一度入力してください。")
    st.caption("🔑 パスワードが分からない場合は管理者にお問い合わせください。")
    return False


# ════ ヘルパー ════════════════════════════════════════════════
def _str_to_time(s, default="15:00"):
    try:
        parts = str(s).strip().split(":")
        return time(int(parts[0]), int(parts[1][:2]))
    except Exception:
        try:
            parts = default.split(":")
            return time(int(parts[0]), int(parts[1]))
        except Exception:
            return time(15, 0)

def _validate_hhmm(s):
    try:
        parts = str(s).strip().split(":")
        h, m = int(parts[0]), int(parts[1][:2])
        if 0 <= h <= 23 and 0 <= m <= 59:
            return f"{h:02d}:{m:02d}"
    except Exception:
        pass
    return None

def load_vehicles(館):
    return load_json_data(館, "vehicles", default=[
        {"車両名":"セレナ","定員":7},{"車両名":"ボクシー","定員":7},
        {"車両名":"白フリード","定員":6},{"車両名":"銀フリ地下","定員":6},
        {"車両名":"銀フリ浅香","定員":6}])

def save_vehicles(館, data):
    save_json_data(館, "vehicles", data)

def load_staff(館):
    return load_json_data(館, "staff", default=[])

def save_staff(館, data):
    save_json_data(館, "staff", data)

def load_custom_places(館):
    return load_json_data(館, "custom_places", default={"pickup": [], "dropoff": []})

def save_custom_places(館, data):
    save_json_data(館, "custom_places", data)

def _is_jiriki(row) -> bool:
    """自力通所かどうかを判定する"""
    tku = str(row.get("通所区分", "")).strip()
    if tku in _JIRIKI_VALUES:
        return True
    pickup  = str(row.get("迎え先（平日）", "") or row.get("迎え先（長期休み）", "")).strip()
    dropoff = str(row.get("送り先", "")).strip()
    if pickup in ("なし", "送迎なし") and dropoff in ("なし", "送迎なし", ""):
        return True
    return False

def _time_cell(col, default_str, base_key):
    opts = _QUICK_TIMES + ["✍️ 手入力"]
    if default_str in _QUICK_TIMES:
        def_idx = _QUICK_TIMES.index(default_str)
    else:
        def_idx = len(opts) - 1
    sel = col.selectbox("", opts, index=def_idx,
                        key=f"{base_key}_sel", label_visibility="collapsed")
    if sel == "✍️ 手入力":
        raw = col.text_input("", value=default_str, key=f"{base_key}_txt",
                             label_visibility="collapsed", placeholder="HH:MM", max_chars=5)
        v = _validate_hhmm(raw)
        return v if v else default_str
    return sel


# ════ ページ: 当日入力 ════════════════════════════════════════
def page_daily(館):
    with st.expander("📖 使い方ガイド（クリックして開く）", expanded=False):
        st.markdown("""
| ステップ | やること | ポイント |
|----------|----------|----------|
| **① 日付を選ぶ** | カレンダーで今日の日付を確認 | 長期休みは「長期休み」をONに |
| **② 参加者チェック** | 固定利用者は自動でチェック済み | 追加参加・休み・変更がある場合だけ修正 |
| **③ 迎え先・時刻を確認** | デフォルト値が自動入力される | 変更がある場合はプルダウンで選択 |
| **④ 自力通所者の時刻** | 送迎不要の方の登所・退所時刻を入力 | Excelに別欄で記載されます |
| **⑤ スタッフ確認** | 本日のスタッフ出勤状況を確認 | 🚗マークが運転担当者 |
| **⑥ 自動生成** | 「送迎ルートを自動生成」ボタンを押す | AIが最適ルートを自動で振り分け |
| **⑦ ルート確認** | 生成されたルートを確認・手動修正 | 車両・ドライバー・順番を変更できます |
| **⑧ Excel出力** | 「Excelをダウンロード」で完成！ | 送迎表が印刷用Excelで保存されます |
        """)

    st.header(f"📅 当日入力 — {館}")

    master_df = load_master(館)
    if master_df.empty:
        st.warning("⚠️ 利用者マスタが未登録です。「利用者マスタ管理」でデータを登録してください。")
        return

    # ── 日付・モード ──
    col_d, col_w, col_m = st.columns([2, 1, 1])
    with col_d:
        target_date = st.date_input("対象日", value=datetime.today())
    with col_w:
        st.metric("曜日", WEEKDAYS[target_date.weekday()])
    with col_m:
        is_long = st.toggle("長期休み", value=False)

    wday       = WEEKDAYS[target_date.weekday()]
    place_col  = "迎え先（長期休み）" if is_long else "迎え先（平日）"
    time_col_p = "迎え時刻（長期休み）" if is_long else "迎え時刻（平日）"

    # ── ソート：固定利用者→非固定→自力通所、各グループ内は児発→放デイ→時刻順 ──
    df_work = master_df.copy()
    df_work["_fixed"]  = df_work["利用曜日"].apply(lambda x: wday in str(x))
    df_work["_jiriki"] = df_work.apply(_is_jiriki, axis=1)

    def _eff_pickup_time_for_sort(row):
        """実効的な迎え時刻（セッション値 > 曜日設定 > デフォルト）— ソート用"""
        pos = row.name
        t_key = f"t_{pos}"
        if t_key in st.session_state:
            v = str(st.session_state[t_key] or "")[:5]
            if v and ":" in v:
                return v
        d_t = str(row.get(f"{wday}_迎え時刻", "") or "").strip()
        if d_t:
            return d_t[:5]
        return str(row.get(time_col_p, "15:00") or "15:00")[:5]

    def _sub_sort(grp: pd.DataFrame) -> pd.DataFrame:
        if grp.empty:
            return grp
        g = grp.copy()
        g["_ku"] = g["区分"].apply(lambda x: 0 if str(x).strip() == "児発" else 1)
        g["_et"] = g.apply(_eff_pickup_time_for_sort, axis=1)
        return g.sort_values(["_ku", "_et"]).drop(columns=["_ku", "_et"])

    df_sorted = pd.concat([
        _sub_sort(df_work[~df_work["_jiriki"] & df_work["_fixed"]]),
        _sub_sort(df_work[~df_work["_jiriki"] & ~df_work["_fixed"]]),
        df_work[df_work["_jiriki"]],
    ], ignore_index=True)

    # ── カスタム送迎先 ──
    custom_places = load_custom_places(館)

    def _build_opts(*cols, extras=None):
        s = set(extras or [])
        for c in cols:
            if c in master_df.columns:
                for v in master_df[c].dropna():
                    v = str(v).strip()
                    if v and v not in ("None", "なし", ""):
                        s.add(v)
        return ["自宅", "なし"] + sorted(s - {"自宅"})

    pickup_opts_base  = _build_opts("迎え先（平日）", "迎え先（長期休み）",
                                    extras=custom_places.get("pickup", []))
    dropoff_opts_base = _build_opts("送り先", extras=custom_places.get("dropoff", []))

    with st.expander("🏫 送迎先を追加・管理", expanded=False):
        p_col, d_col, btn_col = st.columns([3, 3, 1])
        new_pickup_val  = p_col.text_input("迎え先を追加", key="new_pickup_add", placeholder="例: ○○小学校")
        new_dropoff_val = d_col.text_input("送り先を追加", key="new_dropoff_add", placeholder="例: ○○デイ")
        if btn_col.button("＋ 追加", key="add_custom_places", use_container_width=True):
            changed = False
            if new_pickup_val.strip() and new_pickup_val.strip() not in custom_places["pickup"]:
                custom_places["pickup"].append(new_pickup_val.strip()); changed = True
            if new_dropoff_val.strip() and new_dropoff_val.strip() not in custom_places["dropoff"]:
                custom_places["dropoff"].append(new_dropoff_val.strip()); changed = True
            if changed:
                save_custom_places(館, custom_places)
                st.success("✅ 追加しました"); st.rerun()
        if custom_places["pickup"] or custom_places["dropoff"]:
            pl, dl = st.columns(2)
            pl.markdown("**迎え先**"); dl.markdown("**送り先**")
            for p in list(custom_places["pickup"]):
                if pl.button(f"🗑️ {p}", key=f"rm_pickup_{p}"):
                    custom_places["pickup"].remove(p)
                    save_custom_places(館, custom_places); st.rerun()
            for d in list(custom_places["dropoff"]):
                if dl.button(f"🗑️ {d}", key=f"rm_dropoff_{d}"):
                    custom_places["dropoff"].remove(d)
                    save_custom_places(館, custom_places); st.rerun()

    # ══ 送迎あり参加者チェック ══
    st.subheader("📋 参加者チェック（送迎あり）")
    _total_trans = int((~df_work["_jiriki"]).sum())
    _fixed_count = int((~df_work["_jiriki"] & df_work["_fixed"]).sum())
    _jiriki_count = int(df_work["_jiriki"].sum())
    mc1, mc2, mc3 = st.columns(3)
    mc1.metric("送迎対象（全）", f"{_total_trans}名")
    mc2.metric("固定参加", f"{_fixed_count}名", help="本日の曜日が利用曜日に含まれる利用者")
    mc3.metric("自力通所", f"{_jiriki_count}名", help="送迎不要（自力通所・通所区分「自力」）")

    search_q = st.text_input(
        "🔍 名前・地区で検索（スペース・全角無視）",
        placeholder="例：山田　/ やまだ / 2年生",
        key="daily_search",
    )

    df_transport = df_sorted[~df_sorted["_jiriki"]].copy()
    if search_q:
        mask = df_transport.apply(
            lambda r: _fuzzy_match(search_q,
                                   str(r.get("氏名", "")),
                                   str(r.get("地区", "")),
                                   str(r.get("フリガナ", ""))),
            axis=1,
        )
        df_transport = df_transport[mask]

    attendance          = {}
    time_overrides      = {}
    send_time_overrides = {}
    pickup_overrides    = {}
    dropoff_overrides   = {}
    route_orders        = {}

    with st.container(border=True):
        # ヘッダー行：参加 | 状態 | 氏名 | 区分 | 迎え先 | 迎え時刻 | 地区 | 送り先 | 送り時刻 | 順 | 休校 | 備考
        _COL_W = [0.04, 0.03, 0.12, 0.06, 0.14, 0.08, 0.05, 0.14, 0.08, 0.04, 0.06, 0.09]
        hcols = st.columns(_COL_W)
        for c, h in zip(hcols, ["参加","","氏名","区分","迎え先","迎え時刻","地区","送り先","送り時刻","順","休校","備考"]):
            if h:
                c.markdown(f"**{h}**")
        hcols[1].markdown('<span style="font-size:10px" title="🔶固定外 ⏱️時刻変更">状態</span>', unsafe_allow_html=True)

        prev_fixed = None
        for enum_pos, (pos, row) in enumerate(df_transport.iterrows()):
            is_fixed = bool(row.get("_fixed", False))
            if prev_fixed is True and not is_fixed:
                st.markdown("---")
                st.caption("⬇️ 固定曜日外（追加参加の場合はチェック）")
            elif prev_fixed is not None:
                st.markdown('<hr style="border:none;border-top:1px solid #EBEBEB;margin:2px 0">',
                            unsafe_allow_html=True)

            # 曜日別設定があればデフォルト値を上書き
            _day_pp = str(row.get(f"{wday}_迎え先",  "") or "").strip()
            _day_pt = str(row.get(f"{wday}_迎え時刻","") or "").strip()
            _day_dp = str(row.get(f"{wday}_送り先",  "") or "").strip()
            _day_dt = str(row.get(f"{wday}_送り時刻","") or "").strip()

            place_def  = _day_pp or str(row.get(place_col, "") or "自宅").strip() or "自宅"
            t_def_str  = (_day_pt or str(row.get(time_col_p, "15:00") or "15:00"))[:5]
            send_place = _day_dp or str(row.get("送り先", "") or "自宅").strip() or "自宅"
            send_t_def = (_day_dt or str(row.get("送り時刻", "17:00") or "17:00"))[:5]
            ika        = str(row.get("医ケア", "")).strip()

            pickup_opts  = list(pickup_opts_base)
            if place_def not in pickup_opts:
                pickup_opts.insert(2, place_def)
            dropoff_opts = list(dropoff_opts_base)
            if send_place not in dropoff_opts:
                dropoff_opts.insert(2, send_place)

            rc = st.columns(_COL_W)
            attend = rc[0].checkbox("", value=is_fixed, key=f"att_{pos}", label_visibility="collapsed")

            # 状態インジケーター（🔶固定外 / ⏱️時刻変更）
            _icons = []
            if not is_fixed:
                _icons.append("🔶")
            _cur_t = str(st.session_state.get(f"t_{pos}", "") or "")[:5]
            if _cur_t and _cur_t != t_def_str:
                _icons.append("⏱️")
            rc[1].markdown("".join(_icons) or " ", help="🔶=固定曜日外  ⏱️=時刻変更済")

            if ika:
                rc[2].markdown(f"⚕️ **{row.get('氏名','')}**", help=f"医療的ケア: {ika}")
            else:
                rc[2].write(str(row.get("氏名", "")))
            ku = str(row.get("区分", ""))
            badge_color = "#D9E1F2" if ku == "児発" else "#E2EFDA"
            rc[3].markdown(f'<span style="background:{badge_color};padding:2px 4px;border-radius:3px;font-size:11px">{ku}</span>',
                           unsafe_allow_html=True)

            school_hol = rc[10].checkbox("", key=f"hol_{pos}", label_visibility="collapsed", help="学校が休校")
            if school_hol:
                rc[4].caption("🏠 自宅（休校）"); effective_pickup = "自宅"
            else:
                try:
                    pick_idx = pickup_opts.index(place_def)
                except ValueError:
                    pick_idx = 0
                effective_pickup = rc[4].selectbox("", options=pickup_opts, index=pick_idx,
                                                    key=f"pick_{pos}", label_visibility="collapsed")

            t_new_str      = _time_cell(rc[5], t_def_str, f"t_{pos}")
            rc[6].caption(str(row.get("地区", "")))
            try:
                drop_idx = dropoff_opts.index(send_place)
            except ValueError:
                drop_idx = 0
            effective_dropoff  = rc[7].selectbox("", options=dropoff_opts, index=drop_idx,
                                                  key=f"drop_{pos}", label_visibility="collapsed")
            send_t_new_str = _time_cell(rc[8], send_t_def, f"st_{pos}")
            order_val = rc[9].number_input("", min_value=1, max_value=99,
                                           value=enum_pos + 1,
                                           key=f"ord_{pos}", label_visibility="collapsed",
                                           help="同時刻のはしご送迎順")
            rc[11].text_input("", key=f"note_{pos}", label_visibility="collapsed", placeholder="メモ")

            prev_fixed                = is_fixed
            attendance[pos]           = attend
            time_overrides[pos]       = t_new_str
            send_time_overrides[pos]  = send_t_new_str
            pickup_overrides[pos]     = effective_pickup
            dropoff_overrides[pos]    = effective_dropoff
            route_orders[pos]         = order_val

    # ══ 自力通所セクション ══
    df_jiriki  = df_sorted[df_sorted["_jiriki"]].copy()
    jiriki_users = []

    if not df_jiriki.empty:
        st.subheader("🚶 自力通所")
        st.caption("送迎不要の利用者です。登所・退所時刻を入力してください（送迎表に記載されます）。")
        with st.container(border=True):
            jh = st.columns([0.05, 0.20, 0.10, 0.18, 0.18, 0.29])
            for c, h in zip(jh, ["参加","氏名","区分","登所時刻","退所時刻","備考"]):
                c.markdown(f"**{h}**")
            for pos, row in df_jiriki.iterrows():
                is_fixed_j = bool(row.get("_fixed", False))
                jc = st.columns([0.05, 0.20, 0.10, 0.18, 0.18, 0.29])
                attend_j = jc[0].checkbox("", value=is_fixed_j,
                                           key=f"jiriki_att_{pos}", label_visibility="collapsed")
                jc[1].write(str(row.get("氏名", "")))
                ku = str(row.get("区分", ""))
                badge_color = "#D9E1F2" if ku == "児発" else "#E2EFDA"
                jc[2].markdown(f'<span style="background:{badge_color};padding:2px 4px;border-radius:3px;font-size:11px">{ku}</span>',
                               unsafe_allow_html=True)
                arrive_def = str(row.get("迎え時刻（平日）", "15:00") or "15:00")[:5]
                depart_def = str(row.get("送り時刻", "17:00") or "17:00")[:5]
                arrive_str = _time_cell(jc[3], arrive_def, f"jiriki_arrive_{pos}")
                depart_str = _time_cell(jc[4], depart_def, f"jiriki_depart_{pos}")
                note_j     = jc[5].text_input("", key=f"jiriki_note_{pos}",
                                               label_visibility="collapsed", placeholder="メモ")
                if attend_j:
                    jiriki_users.append({
                        "name":   str(row.get("氏名", "")),
                        "kubun":  ku,
                        "arrive": arrive_str,
                        "depart": depart_str,
                        "note":   note_j,
                    })

    # ── スタッフ ──
    st.subheader("👤 本日のスタッフ勤務状況")
    staff_data    = load_staff(館)
    staff_on_duty = {}

    if not staff_data:
        st.info("スタッフ未登録 →「車両・スタッフ設定」で登録してください")
    else:
        shift_opts = list(SHIFT_PRESETS.keys())
        h0,h1,h2,h3,h4,h5 = st.columns([0.04,0.16,0.14,0.10,0.10,0.08])
        for col, lbl in [(h0,"出勤"),(h1,"氏名"),(h2,"勤務区分"),(h3,"開始"),(h4,"終了"),(h5,"運転")]:
            col.markdown(f"**{lbl}**")
        st.markdown('<hr style="margin:4px 0">', unsafe_allow_html=True)

        for i, s in enumerate(staff_data):
            name      = s.get("氏名", f"staff_{i}")
            can_drive = bool(s.get("運転可", False))
            def_shift = s.get("default_shift", "B勤")
            if def_shift not in shift_opts:
                def_shift = "B勤"

            c0,c1,c2,c3,c4,c5 = st.columns([0.04,0.16,0.14,0.10,0.10,0.08])
            default_on = def_shift not in _OFF_SHIFTS
            attend_s = c0.checkbox("", value=default_on, key=f"stf_on_{i}", label_visibility="collapsed")
            icon = "🚗" if can_drive else "👤"
            c1.write(f"{icon} {name}")
            def_idx = shift_opts.index(def_shift)
            shift = c2.selectbox("", shift_opts, index=def_idx,
                                 key=f"stf_shift_{i}", label_visibility="collapsed",
                                 disabled=not attend_s)

            is_off = (shift in _OFF_SHIFTS) or (not attend_s)
            preset = SHIFT_PRESETS.get(shift)

            if is_off:
                c3.caption("—"); c4.caption("—")
                work_start = work_end = ""
            elif shift == "カスタム":
                prev_s = s.get("custom_start", "09:30")
                prev_e = s.get("custom_end",   "18:30")
                ws = c3.text_input("", value=prev_s, key=f"stf_start_{i}",
                                   label_visibility="collapsed", max_chars=5, placeholder="HH:MM")
                we = c4.text_input("", value=prev_e, key=f"stf_end_{i}",
                                   label_visibility="collapsed", max_chars=5, placeholder="HH:MM")
                work_start = _validate_hhmm(ws) or prev_s
                work_end   = _validate_hhmm(we) or prev_e
            else:
                work_start, work_end = preset
                c3.caption(work_start); c4.caption(work_end)

            drv_default = can_drive and attend_s and not is_off
            drive_on = c5.checkbox("", value=drv_default, key=f"stf_drv_{i}",
                                   label_visibility="collapsed",
                                   disabled=(not can_drive or not attend_s or is_off))
            staff_on_duty[name] = {
                "on":         attend_s and not is_off,
                "drive":      drive_on,
                "shift":      shift,
                "work_start": work_start,
                "work_end":   work_end,
            }

    st.divider()
    if st.button("🚐 送迎ルートを自動生成", type="primary", use_container_width=True):
        attending_rows = []
        for pos, row in df_transport.iterrows():
            if not attendance.get(pos, False):
                continue
            r = {c: row.get(c, "") for c in master_df.columns if c in row}
            r[time_col_p] = time_overrides.get(pos, r.get(time_col_p, "15:00"))
            r["送り時刻"] = send_time_overrides.get(pos, r.get("送り時刻", "17:00"))
            r[place_col]  = pickup_overrides.get(pos, r.get(place_col, "自宅"))
            r["送り先"]   = dropoff_overrides.get(pos, r.get("送り先", "自宅"))
            r["送迎順"]   = route_orders.get(pos, 99)
            # ── 自宅住所の紐づけ：「自宅」の場合は住所列から実際の住所をセット ──
            home_addr = str(row.get("住所", "")).strip()
            if r.get(place_col, "") in ("自宅", "") and home_addr:
                r["迎え先住所"] = home_addr
            if r.get("送り先", "") in ("自宅", "") and home_addr:
                r["送り先住所"] = home_addr
            attending_rows.append(r)

        if not attending_rows and not jiriki_users:
            st.error("参加者が選択されていません")
            return

        parts    = pd.DataFrame(attending_rows) if attending_rows else pd.DataFrame()
        vehicles = load_vehicles(館)
        drivers  = [n for n, i in staff_on_duty.items() if i.get("drive")]
        routes   = generate_routes(parts, vehicles, drivers, is_long_holiday=is_long) if not parts.empty else {}

        st.session_state.routes        = routes
        st.session_state.jiriki_users  = jiriki_users
        st.session_state.target_date   = target_date
        st.session_state.staff_on_duty = staff_on_duty
        st.session_state["selected_kan"] = 館
        _attend_count = len(attending_rows)
        _vehicle_count = len(routes)
        st.success(f"✅ ルートを生成しました！ 【送迎あり {_attend_count}名 / 車両 {_vehicle_count}台】 ↓ 下にスクロールして確認・修正してください")

    # ── ルート確認・修正 ──
    if st.session_state.get("routes") is not None:
        routes         = st.session_state.routes
        stored_jiriki  = st.session_state.get("jiriki_users", [])

        if routes:
            st.divider()
            st.subheader("🗺️ 生成ルート（確認・修正）")
            sod             = st.session_state.get("staff_on_duty", {})
            drivers_on_duty = [n for n, i in sod.items() if i.get("drive")]
            edited = {}

            for vehicle, trips in routes.items():
                with st.container(border=True):
                    col_vh, col_drv = st.columns([3, 2])
                    col_vh.markdown(f"#### 🚐 {vehicle}")
                    drv_opts   = ["（便ごとに設定）"] + drivers_on_duty
                    veh_driver = col_drv.selectbox("担当ドライバー（一括適用）",
                                                   drv_opts, key=f"veh_drv_{vehicle}")
                    effective_trips = []
                    for t in trips:
                        t2 = dict(t)
                        if veh_driver != "（便ごとに設定）":
                            t2["driver"] = veh_driver
                        effective_trips.append(t2)

                    tab_m, tab_o = st.tabs(["迎え", "送り"])
                    with tab_m:
                        edited.setdefault(vehicle, {})["迎え"] = _trip_editor(
                            [t for t in effective_trips if t["type"]=="迎え"],
                            vehicle, "迎え", drivers=drivers_on_duty)
                    with tab_o:
                        edited.setdefault(vehicle, {})["送り"] = _trip_editor(
                            [t for t in effective_trips if t["type"]=="送り"],
                            vehicle, "送り", drivers=drivers_on_duty)
        else:
            edited = {}

        if stored_jiriki:
            st.divider()
            st.subheader("🚶 自力通所（確認）")
            jc_h = st.columns([0.25, 0.10, 0.15, 0.15, 0.35])
            for c, h in zip(jc_h, ["氏名","区分","登所時刻","退所時刻","備考"]):
                c.markdown(f"**{h}**")
            for u in stored_jiriki:
                jc_r = st.columns([0.25, 0.10, 0.15, 0.15, 0.35])
                jc_r[0].write(u["name"]); jc_r[1].write(u["kubun"])
                jc_r[2].write(u["arrive"]); jc_r[3].write(u["depart"])
                jc_r[4].write(u.get("note",""))

        st.divider()
        td    = st.session_state.target_date
        fname = f"{td.strftime('%Y%m%d')}_{館}送迎表.xlsx"
        excel_bytes = export_schedule(
            edited, td, 館,
            st.session_state.get("staff_on_duty", {}),
            jiriki_users=stored_jiriki,
            master_df=load_master(館),
            colors=load_colors(館))
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                "📥 Excelをダウンロード", data=excel_bytes, file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, type="primary")
        with col2:
            if st.button("🔄 再生成", use_container_width=True):
                st.session_state.routes       = None
                st.session_state.jiriki_users = []
                st.rerun()


def _trip_editor(trips, vehicle, trip_type, drivers=None):
    if not trips:
        st.caption("（なし）"); return []
    edited      = []
    driver_opts = [""] + (drivers or [])
    for i, trip in enumerate(trips):
        c1, c2, c3, c4 = st.columns([0.15, 0.25, 0.3, 0.3])
        t = c1.time_input("", value=_str_to_time(trip.get("time","15:00")),
                          key=f"{vehicle}_{trip_type}_{i}_t", step=300,
                          label_visibility="collapsed")
        name  = c2.text_input("", value=trip.get("name",""),
                               key=f"{vehicle}_{trip_type}_{i}_n", label_visibility="collapsed")
        place = c3.text_input("", value=trip.get("place",""),
                               key=f"{vehicle}_{trip_type}_{i}_p", label_visibility="collapsed")
        cur_drv = trip.get("driver", "")
        if drivers:
            try:
                drv_idx = driver_opts.index(cur_drv)
            except ValueError:
                drv_idx = 0
            driver = c4.selectbox("", options=driver_opts, index=drv_idx,
                                   key=f"{vehicle}_{trip_type}_{i}_d", label_visibility="collapsed")
        else:
            driver = c4.text_input("", value=cur_drv,
                                    key=f"{vehicle}_{trip_type}_{i}_d", label_visibility="collapsed")
        edited.append({"type":trip_type,"time":str(t)[:5],"name":name,"place":place,"driver":driver})
    return edited


# ════ ページ: 利用者マスタ ════════════════════════════════════
def page_master(館):
    with st.expander("📖 使い方ガイド（クリックして開く）", expanded=False):
        st.markdown("""
### 👤 利用者マスタ管理の使い方

| タブ | 使うとき |
|------|----------|
| ✏️ **手動編集** | 利用者を1人ずつ追加・修正する |
| 🏠 **保護者住所取込** | リタリコ等から書き出した保護者一覧CSVを使って住所をまとめて登録 |
| 📋 **内部CSV取込** | 社内の利用者一覧CSVからまとめて登録 |
| 📂 **リタリコCSV** | リタリコの形式に合わせた列マッピングで取込 |
| 🗺️ **住所検索** | Google Maps APIで送迎先施設の住所を検索して登録 |
| 🕐 **変更履歴** | 間違えて保存した場合に過去の状態に戻せる |

**💡 よくある手順（はじめて使う場合）**
1. 「内部CSV取込」か「リタリコCSV」タブで利用者を一括登録
2. 「保護者住所取込」タブで自宅住所を一括登録
3. 「手動編集」タブで細かい情報を調整
        """)

    st.header(f"👥 利用者マスタ管理 — {館}")

    tab_edit, tab_hogosha, tab_csv, tab_ritalico, tab_addr, tab_hist = st.tabs([
        "✏️ 手動編集", "🏠 保護者住所取込", "📋 内部CSV取込", "📂 リタリコCSV", "🗺️ 住所検索", "🕐 変更履歴",
    ])

    # ── 手動編集 ──────────────────────────────────────────────
    with tab_edit:
        master_df = load_master(館)
        if master_df.empty:
            master_df = pd.DataFrame(columns=MASTER_COLUMNS)

        # 統計カード
        _n_total  = len(master_df)
        _n_active = int((master_df.get("状態", pd.Series([""] * len(master_df), index=master_df.index))
                         .apply(lambda x: str(x).strip() not in ("退所","終了","inactive"))).sum())
        _n_jiriki = int(master_df.apply(
            lambda r: str(r.get("通所区分","")).strip() in _JIRIKI_VALUES, axis=1).sum())
        _n_ika    = int((master_df.get("医ケア", pd.Series([""] * len(master_df), index=master_df.index))
                         .apply(lambda x: bool(str(x).strip()))).sum())
        sc1, sc2, sc3, sc4 = st.columns(4)
        sc1.metric("登録人数", f"{_n_total}名")
        sc2.metric("在籍中", f"{_n_active}名")
        sc3.metric("自力通所", f"{_n_jiriki}名")
        sc4.metric("医療的ケア", f"{_n_ika}名")

        st.info("💡 **通所区分**「自力」→ 当日入力で自力通所扱い。**住所**は保護者住所取込タブから一括登録できます。")

        # ── 名前一覧（クリックで詳細編集）──────────────────────
        sq = st.text_input("🔍 名前で検索（スペース無視・部分一致）",
                           key="master_search", placeholder="例：山田 / やまだ / ヤマダ")

        if sq:
            _mask = master_df.apply(
                lambda r: _fuzzy_match(sq, str(r.get("氏名","")), str(r.get("フリガナ",""))), axis=1)
            disp_df = master_df[_mask]
        else:
            disp_df = master_df

        disp_indices = list(disp_df.index)
        _name_tbl = disp_df[["氏名","フリガナ","区分","利用曜日","状態"]].reset_index(drop=True)

        _lc, _rc = st.columns([5, 1])
        _lc.caption(f"📋 {len(_name_tbl)}名 — 行をクリックして詳細を編集")
        if _rc.button("➕ 新規追加", key="add_new_master_btn", use_container_width=True):
            _new_row = {col: "" for col in MASTER_COLUMNS}
            master_df = pd.concat([master_df, pd.DataFrame([_new_row])], ignore_index=True)
            save_master(館, master_df)
            st.rerun()

        _ev = st.dataframe(
            _name_tbl,
            use_container_width=True,
            hide_index=True,
            on_select="rerun",
            selection_mode="single-row",
            column_config={
                "氏名":    st.column_config.TextColumn("氏名",    width="medium"),
                "フリガナ": st.column_config.TextColumn("フリガナ", width="medium"),
                "区分":    st.column_config.TextColumn("区分",    width="small"),
                "利用曜日": st.column_config.TextColumn("利用曜日", width="small"),
                "状態":    st.column_config.TextColumn("状態",    width="small"),
            },
        )

        # ── 詳細編集パネル（行選択時に表示）──────────────────────
        _sel = _ev.selection.rows if hasattr(_ev, "selection") else []
        if _sel:
            _orig_idx = disp_indices[_sel[0]]
            _row = master_df.loc[_orig_idx]
            _ek  = _orig_idx  # widget key suffix

            st.divider()
            st.subheader(f"✏️ {_row.get('氏名','新規')}")

            _dtab_b, _dtab_s, _dtab_o = st.tabs(["📋 基本情報", "🚐 曜日別送迎設定", "📝 その他"])
            _nv = {}  # new values

            with _dtab_b:
                _c1, _c2, _c3 = st.columns(3)
                _nv["氏名"]        = _c1.text_input("氏名", value=str(_row.get("氏名","") or ""), key=f"ev_nm_{_ek}")
                _nv["フリガナ"]     = _c2.text_input("フリガナ", value=str(_row.get("フリガナ","") or ""), key=f"ev_kn_{_ek}")
                _nv["受給者証番号"] = _c3.text_input("受給者証番号", value=str(_row.get("受給者証番号","") or ""), key=f"ev_rn_{_ek}")

                _c4, _c5, _c6, _c7 = st.columns(4)
                _KU  = ["放デイ","児発"]
                _ST  = ["","在籍","退所","体験"]
                _TC  = ["","送迎","自力"]
                _IKA = ["","医ケア1","医ケア2","医ケア3"]
                _ku_v  = str(_row.get("区分","放デイ") or "放デイ")
                _st_v  = str(_row.get("状態","") or "")
                _tc_v  = str(_row.get("通所区分","") or "")
                _ika_v = str(_row.get("医ケア","") or "")
                _nv["区分"]     = _c4.selectbox("区分",     _KU,  index=_KU.index(_ku_v)   if _ku_v  in _KU  else 0, key=f"ev_ku_{_ek}")
                _nv["状態"]     = _c5.selectbox("状態",     _ST,  index=_ST.index(_st_v)   if _st_v  in _ST  else 0, key=f"ev_st_{_ek}")
                _nv["利用曜日"] = _c6.text_input("利用曜日（月水金など）", value=str(_row.get("利用曜日","") or ""), key=f"ev_wd_{_ek}")
                _nv["地区"]     = _c7.text_input("地区", value=str(_row.get("地区","") or ""), key=f"ev_ar_{_ek}")

                _c8, _c9, _c10, _c11 = st.columns(4)
                _nv["医ケア"]   = _c8.selectbox("医ケア",   _IKA, index=_IKA.index(_ika_v) if _ika_v in _IKA else 0, key=f"ev_ik_{_ek}")
                _nv["重心"]     = _c9.text_input("重心", value=str(_row.get("重心","") or ""), key=f"ev_js_{_ek}")
                _nv["通所区分"] = _c10.selectbox("通所区分", _TC,  index=_TC.index(_tc_v)   if _tc_v  in _TC  else 0, key=f"ev_tc_{_ek}")
                try:
                    _cnt_v = int(str(_row.get("契約上限","0") or "0").strip() or 0)
                except Exception:
                    _cnt_v = 0
                _nv["契約上限"] = _c11.number_input("契約上限", value=_cnt_v, min_value=0, key=f"ev_cn_{_ek}")

                st.markdown("**― デフォルト送迎設定（曜日設定がない場合に使用）―**")
                _da1, _da2, _da3 = st.columns(3)
                _nv["迎え先（平日）"]    = _da1.text_input("迎え先（平日）",       value=str(_row.get("迎え先（平日）","")    or ""), key=f"ev_pp_{_ek}")
                _nv["迎え時刻（平日）"]  = _da1.text_input("迎え時刻（平日）HH:MM", value=str(_row.get("迎え時刻（平日）","")  or ""), key=f"ev_pt_{_ek}", max_chars=5, placeholder="15:00")
                _nv["迎え先（長期休み）"] = _da2.text_input("迎え先（長期休み）",   value=str(_row.get("迎え先（長期休み）","") or ""), key=f"ev_lp_{_ek}")
                _nv["迎え時刻（長期休み）"] = _da2.text_input("迎え時刻（長休）HH:MM", value=str(_row.get("迎え時刻（長期休み）","") or ""), key=f"ev_lt_{_ek}", max_chars=5, placeholder="10:00")
                _nv["送り先"]   = _da3.text_input("送り先",       value=str(_row.get("送り先","")   or ""), key=f"ev_dp_{_ek}")
                _nv["送り時刻"] = _da3.text_input("送り時刻 HH:MM", value=str(_row.get("送り時刻","") or ""), key=f"ev_dt_{_ek}", max_chars=5, placeholder="17:00")
                _nv["住所"]    = st.text_input("住所（自宅）", value=str(_row.get("住所","") or ""), key=f"ev_ad_{_ek}")

            with _dtab_s:
                st.caption("各曜日の送迎設定 — 空白の場合はデフォルト設定を使用します")
                _wday_tabs = st.tabs(["月", "火", "水", "木", "金"])
                for _wt, _wd in zip(_wday_tabs, ["月","火","水","木","金"]):
                    with _wt:
                        _pw1, _pw2 = st.columns(2)
                        _nv[f"{_wd}_迎え先"]  = _pw1.text_input("迎え先（空白=デフォルト）",       value=str(_row.get(f"{_wd}_迎え先","")  or ""), key=f"ev_{_wd}_pp_{_ek}")
                        _nv[f"{_wd}_迎え時刻"] = _pw1.text_input("迎え時刻 HH:MM（空白=デフォルト）", value=str(_row.get(f"{_wd}_迎え時刻","") or ""), key=f"ev_{_wd}_pt_{_ek}", max_chars=5, placeholder="HH:MM")
                        _nv[f"{_wd}_送り先"]  = _pw2.text_input("送り先（空白=デフォルト）",         value=str(_row.get(f"{_wd}_送り先","")  or ""), key=f"ev_{_wd}_dp_{_ek}")
                        _nv[f"{_wd}_送り時刻"] = _pw2.text_input("送り時刻 HH:MM（空白=デフォルト）", value=str(_row.get(f"{_wd}_送り時刻","") or ""), key=f"ev_{_wd}_dt_{_ek}", max_chars=5, placeholder="HH:MM")

            with _dtab_o:
                _nv["特記事項"] = st.text_area("特記事項", value=str(_row.get("特記事項","") or ""), key=f"ev_nt_{_ek}", height=80)
                _nv["備考"]    = st.text_area("備考",    value=str(_row.get("備考","")    or ""), key=f"ev_bk_{_ek}", height=80)
                _oc1, _oc2 = st.columns(2)
                _nv["利用開始日"] = _oc1.text_input("利用開始日（YYYY-MM-DD）", value=str(_row.get("利用開始日","") or ""), key=f"ev_sd_{_ek}")
                _nv["利用終了日"] = _oc2.text_input("利用終了日（YYYY-MM-DD）", value=str(_row.get("利用終了日","") or ""), key=f"ev_ed_{_ek}")
                _nv["契約月"]    = _oc1.text_input("契約月", value=str(_row.get("契約月","") or ""), key=f"ev_cm_{_ek}")
                _nv["迎え先住所"] = _oc2.text_input("迎え先住所（学校など）", value=str(_row.get("迎え先住所","") or ""), key=f"ev_pa_{_ek}")
                _nv["送り先住所"] = st.text_input("送り先住所", value=str(_row.get("送り先住所","") or ""), key=f"ev_da_{_ek}")

            # 保存・削除ボタン
            _sc, _dc = st.columns([4, 1])
            if _sc.button("💾 保存する", type="primary", key=f"save_detail_{_ek}", use_container_width=True):
                for _k, _v in _nv.items():
                    master_df.at[_orig_idx, _k] = _v
                save_master(館, master_df)
                st.success(f"✅ 「{_nv.get('氏名','')}」を保存しました！")
                st.rerun()

            if _dc.button("🗑️ 削除", key=f"del_btn_{_ek}", use_container_width=True):
                st.session_state[f"_cdel_{_ek}"] = True

            if st.session_state.get(f"_cdel_{_ek}"):
                st.warning(f"⚠️ 「{_row.get('氏名','')}」を削除しますか？この操作は元に戻せません。")
                _yc, _nc = st.columns(2)
                if _yc.button("はい、削除する", key=f"del_yes_{_ek}", type="primary"):
                    master_df = master_df.drop(index=_orig_idx).reset_index(drop=True)
                    save_master(館, master_df)
                    st.session_state.pop(f"_cdel_{_ek}", None)
                    st.success("削除しました")
                    st.rerun()
                if _nc.button("キャンセル", key=f"del_no_{_ek}"):
                    st.session_state.pop(f"_cdel_{_ek}", None)
                    st.rerun()

    # ── 保護者一覧 住所取込 ───────────────────────────────────
    with tab_hogosha:
        st.subheader("🏠 保護者一覧CSVから自宅住所を取込")
        st.info("""
リタリコ等からエクスポートした **保護者一覧CSV** をアップロードしてください。
「児童」列の名前でマスタと照合し、「住所」列（都道府県＋市区町村＋番地）を自動更新します。
複数児童が「、」区切りで入っている場合も対応しています。
        """)

        uploaded_h = st.file_uploader("保護者一覧CSVを選択", type=["csv"], key="hogosha_csv")
        if uploaded_h is not None:
            try:
                import io as _io
                import traceback as _tb
                raw_bytes = uploaded_h.read()
                if not raw_bytes:
                    st.error("ファイルが空です。再度アップロードしてください。")
                else:
                    hdf = None
                    used_enc = None
                    for enc in ("utf-8-sig", "cp932", "shift-jis", "utf-8"):
                        try:
                            hdf = pd.read_csv(_io.BytesIO(raw_bytes), encoding=enc, dtype=str)
                            used_enc = enc
                            break
                        except Exception:
                            continue

                    if hdf is None:
                        st.error("文字コードを判定できませんでした。UTF-8 または Shift-JIS (cp932) の CSV をアップロードしてください。")
                    else:
                        st.caption(f"文字コード: {used_enc} で読込成功 ({len(hdf)}件)")
                        preview_cols = [c for c in ("児童","都道府県","市区町村","番地","ビル・マンション名") if c in hdf.columns]
                        st.write("**プレビュー（先頭3行）**")
                        st.dataframe(hdf[preview_cols].head(3) if preview_cols else hdf.head(3),
                                     use_container_width=True)

                        master_df = load_master(館)
                        if master_df.empty:
                            st.warning("マスタが未登録です。先に利用者マスタを登録してください。")
                        else:
                            updated_df, matched, unmatched = import_address_from_hogosha_csv(hdf, master_df)

                            col_a, col_b = st.columns(2)
                            col_a.metric("照合一致", f"{matched}名")
                            col_b.metric("未一致",   f"{len(unmatched)}名")

                            if unmatched:
                                with st.expander(f"⚠️ 未一致の児童名（{len(unmatched)}名）"):
                                    st.write("、".join(unmatched))

                            changed_mask = updated_df["住所"] != master_df["住所"].reindex(
                                updated_df.index, fill_value="")
                            if changed_mask.any():
                                st.write(f"**更新される住所（{changed_mask.sum()}件）**")
                                st.dataframe(updated_df.loc[changed_mask, ["氏名","地区","住所"]],
                                             use_container_width=True)

                            if st.button("💾 住所を保存する", type="primary", key="hogosha_save"):
                                save_master(館, updated_df)
                                st.success(f"✅ {matched}名の住所を保存しました")
                                st.rerun()
            except Exception as _e:
                st.error(f"処理中にエラーが発生しました: {_e}")
                st.code(_tb.format_exc())

    # ── 内部CSV取込 ───────────────────────────────────────────
    with tab_csv:
        st.subheader("📋 内部利用者一覧CSV を取込")
        st.info("受給者証番号を主キーとして管理します。曜日列の「2」は自動的に「Ⅱ番館」に変換されます。")
        uploaded_csv = st.file_uploader("利用者一覧CSVを選択", type=["csv"], key="internal_csv")
        if uploaded_csv is not None:
            import io as _io_csv
            try:
                raw_bytes = uploaded_csv.read()
                raw_df = None
                for enc in ("utf-8-sig", "shift-jis", "cp932", "utf-8"):
                    try:
                        raw_df = pd.read_csv(_io_csv.BytesIO(raw_bytes), encoding=enc, dtype=str)
                        st.caption(f"文字コード: {enc} で読込成功")
                        break
                    except Exception:
                        continue
                if raw_df is None:
                    st.error("文字コードを判定できませんでした。UTF-8 または Shift-JIS の CSV をアップロードしてください。")
                else:
                    st.write(f"**読込件数: {len(raw_df)}名**")
                    st.dataframe(raw_df.head(5), use_container_width=True)

                    hall_dfs = import_from_internal_csv(raw_df)
                    st.subheader("番館別 振り分けプレビュー")
                    cols4 = st.columns(4)
                    for i, h in enumerate(HALLS):
                        cols4[i].metric(h, f"{len(hall_dfs.get(h, []))}名")

                    all_rows = pd.concat([d for d in hall_dfs.values() if not d.empty], ignore_index=True)
                    if not all_rows.empty and "受給者証番号" in all_rows.columns:
                        temp_mask  = all_rows["受給者証番号"].apply(is_temp_juki_no)
                        temp_users = all_rows[temp_mask]
                        if len(temp_users):
                            with st.expander(f"⚠️ 受給者証番号が未確定: {len(temp_users)}名"):
                                st.dataframe(temp_users[["氏名","受給者証番号"]], use_container_width=True)

                    df_cur = hall_dfs.get(館, pd.DataFrame())
                    if not df_cur.empty:
                        st.subheader(f"{館} の取込データ（{len(df_cur)}名）")
                        preview_c = ["氏名","区分","通所区分","利用曜日","迎え先（平日）","送り先"]
                        st.dataframe(df_cur[[c for c in preview_c if c in df_cur.columns]],
                                     use_container_width=True)

                    c_left, c_right = st.columns(2)
                    import_target = c_left.radio("取込対象", [f"現在の番館のみ（{館}）","全番館（Ⅰ〜Ⅴ）"])
                    merge_mode    = c_right.radio("既存データとの統合", ["上書き（全件置換）","追加・更新（受給者証番号で照合）"])

                    if st.button("📥 取込・保存", type="primary"):
                        targets     = [館] if "現在" in import_target else HALLS
                        saved_total = 0
                        for h in targets:
                            new_df = hall_dfs.get(h, pd.DataFrame())
                            if new_df.empty:
                                continue
                            if "追加・更新" in merge_mode:
                                existing = load_master(h)
                                if not existing.empty and "受給者証番号" in existing.columns:
                                    merged = existing.copy()
                                    for _, row in new_df.iterrows():
                                        jid      = str(row.get("受給者証番号","")).strip()
                                        idx_list = merged.index[merged["受給者証番号"].str.strip()==jid].tolist()
                                        if idx_list:
                                            merged.loc[idx_list[0]] = row
                                        else:
                                            merged = pd.concat([merged, row.to_frame().T], ignore_index=True)
                                    save_master(h, merged)
                                else:
                                    save_master(h, new_df)
                            else:
                                save_master(h, new_df)
                            saved_total += len(new_df)
                        st.success(f"✅ {saved_total}件を保存しました（{', '.join(targets)}）")
                        st.rerun()
            except Exception as _e_csv:
                import traceback as _tb_csv
                st.error(f"CSVの処理中にエラーが発生しました: {_e_csv}")
                st.code(_tb_csv.format_exc())

    # ── リタリコCSV ───────────────────────────────────────────
    with tab_ritalico:
        st.subheader("📂 リタリコCSVをインポート")
        uploaded = st.file_uploader("CSVファイルを選択", type=["csv"], key="ritalico_csv")
        if uploaded is not None:
            import io as _io_rl
            try:
                raw_bytes = uploaded.read()
                raw_df = None
                for enc in ("utf-8-sig", "shift-jis", "cp932", "utf-8"):
                    try:
                        raw_df = pd.read_csv(_io_rl.BytesIO(raw_bytes), encoding=enc)
                        st.caption(f"文字コード: {enc} で読込成功")
                        break
                    except Exception:
                        continue
                if raw_df is None:
                    st.error("文字コードを判定できませんでした。UTF-8 または Shift-JIS の CSV をアップロードしてください。")
                else:
                    st.dataframe(raw_df.head(), use_container_width=True)
                    cols_csv = ["（未選択）"] + list(raw_df.columns)

                    def _g(*cands):
                        for c in cands:
                            if c in raw_df.columns: return c
                        return "（未選択）"

                    st.subheader("列マッピング")
                    c1, c2 = st.columns(2)
                    col_name   = c1.selectbox("氏名列（必須）", cols_csv,
                        index=cols_csv.index(_g("児童","氏名","利用者名")), key="m_name")
                    col_kana   = c2.selectbox("フリガナ列", cols_csv,
                        index=cols_csv.index(_g("児童カナ","フリガナ")), key="m_kana")
                    col_addr   = st.selectbox("住所列", cols_csv, index=0, key="m_addr")
                    c3, c4, c5 = st.columns(3)
                    col_school = c3.selectbox("迎え先列", cols_csv,
                        index=cols_csv.index(_g("学校名","通学先","迎え先")), key="m_school")
                    col_type   = c4.selectbox("区分列", cols_csv,
                        index=cols_csv.index(_g("サービス種別","区分")), key="m_type")
                    col_day    = c5.selectbox("利用曜日列", cols_csv,
                        index=cols_csv.index(_g("利用曜日","曜日")), key="m_day")

                    if st.button("📥 取込・保存", type="primary", key="ritalico_save"):
                        mapping   = {"氏名":col_name,"フリガナ":col_kana,"住所":col_addr,
                                     "迎え先（平日）":col_school,"区分":col_type,"利用曜日":col_day}
                        result_df = import_from_ritalico(raw_df, mapping)
                        save_master(館, result_df)
                        st.success(f"✅ {len(result_df)}件を取込みました！")
            except Exception as _e_rl:
                import traceback as _tb_rl
                st.error(f"処理中にエラーが発生しました: {_e_rl}")
                st.code(_tb_rl.format_exc())

    # ── 住所検索 ──────────────────────────────────────────────
    with tab_addr:
        st.subheader("🗺️ 送迎先 住所検索")
        master_df  = load_master(館)
        facilities = get_facilities_needing_address(master_df)
        if not facilities:
            st.info("送迎先の施設がありません。先に利用者マスタを取り込んでください。")
        else:
            api_key = ""
            try:
                api_key = st.secrets.get("google_maps", {}).get("api_key", "")
            except Exception:
                pass
            if not api_key:
                st.warning("⚠️ Google Maps APIキーが未設定です（Secrets に [google_maps] api_key を追加してください）")

            addr_map = load_json_data(館, "address_map", default={})

            st.caption(f"{len(facilities)}件の送迎先の住所を設定できます")
            changed = False
            for fac in facilities:
                c1, c2, c3 = st.columns([3, 5, 1])
                c1.markdown(f"**{fac}**")
                cur = addr_map.get(fac, "")
                new_val = c2.text_input("住所", value=cur, key=f"addr_{fac}",
                                        label_visibility="collapsed",
                                        placeholder="例: 東京都新宿区西新宿1-1-1")
                if new_val != cur:
                    addr_map[fac] = new_val
                    changed = True
                if api_key:
                    if c3.button("🔍", key=f"look_{fac}", help="Google Mapsで検索"):
                        with st.spinner(f"{fac} を検索中…"):
                            found = lookup_address_google(fac, api_key)
                        if found:
                            addr_map[fac] = found
                            save_json_data(館, "address_map", addr_map)
                            st.success(f"✅ {fac}: {found}")
                            st.rerun()
                        else:
                            st.warning(f"⚠️ {fac} の住所が見つかりませんでした")

            if st.button("💾 住所マップを保存", type="primary", key="addr_save"):
                save_json_data(館, "address_map", addr_map)
                st.success("✅ 住所マップを保存しました")

    # ── 変更履歴 ──────────────────────────────────────────────
    with tab_hist:
        st.subheader("🕐 変更履歴")
        hist_list = load_history_list(館)
        if not hist_list:
            st.info("変更履歴がありません。マスタを保存すると履歴が作成されます。")
        else:
            sel = st.selectbox("復元する履歴を選択", hist_list, key="hist_sel")
            if sel:
                hist_df = load_history(館, sel)
                st.dataframe(hist_df, use_container_width=True)
                if st.button("⏪ この履歴に戻す", type="primary", key="hist_restore"):
                    save_master(館, hist_df)
                    st.success(f"✅ {sel} の状態に戻しました")
                    st.rerun()


# ════ ページ: 車両・スタッフ・カラー設定 ══════════════════════════
def page_settings(館):
    st.header(f"⚙️ 設定 — {館}")
    tab_v, tab_s, tab_p, tab_c = st.tabs(["🚐 車両", "👤 スタッフ", "📍 カスタム送迎先", "🎨 カラー"])

    with tab_v:
        vdf = pd.DataFrame(load_vehicles(館))
        ev = st.data_editor(vdf, num_rows="dynamic", use_container_width=True,
                            column_config={
                                "定員": st.column_config.NumberColumn("定員", min_value=1, max_value=20)
                            }, key="veh_ed")
        if st.button("💾 車両を保存", type="primary", key="sv_v"):
            save_vehicles(館, ev.to_dict("records"))
            st.success("✅ 保存しました")

    with tab_s:
        raw_s = load_staff(館)
        sdf = pd.DataFrame(raw_s) if raw_s else pd.DataFrame(columns=["氏名","運転可","備考"])
        es = st.data_editor(sdf, num_rows="dynamic", use_container_width=True,
                            column_config={
                                "運転可": st.column_config.CheckboxColumn("運転可")
                            }, key="stf_ed")
        if st.button("💾 スタッフを保存", type="primary", key="sv_s"):
            save_staff(館, es.to_dict("records"))
            st.success("✅ 保存しました")

    with tab_p:
        st.caption("ここで登録した送迎先は「迎え先」「送り先」のプルダウンに表示されます。")
        raw_cp = load_custom_places(館)
        st.markdown("**迎え先**")
        new_p = st.text_input("追加する迎え先", key="cp_new_p", placeholder="例: ○○小学校")
        if st.button("＋ 迎え先を追加", key="cp_add_p"):
            if new_p.strip() and new_p.strip() not in raw_cp.get("pickup", []):
                raw_cp.setdefault("pickup", []).append(new_p.strip())
                save_custom_places(館, raw_cp)
                st.rerun()
        for pp in list(raw_cp.get("pickup", [])):
            if st.button(f"🗑️ {pp}", key=f"cp_rm_p_{pp}"):
                raw_cp["pickup"].remove(pp)
                save_custom_places(館, raw_cp)
                st.rerun()
        st.markdown("**送り先**")
        new_d = st.text_input("追加する送り先", key="cp_new_d", placeholder="例: ○○デイ")
        if st.button("＋ 送り先を追加", key="cp_add_d"):
            if new_d.strip() and new_d.strip() not in raw_cp.get("dropoff", []):
                raw_cp.setdefault("dropoff", []).append(new_d.strip())
                save_custom_places(館, raw_cp)
                st.rerun()
        for dd in list(raw_cp.get("dropoff", [])):
            if st.button(f"🗑️ {dd}", key=f"cp_rm_d_{dd}"):
                raw_cp["dropoff"].remove(dd)
                save_custom_places(館, raw_cp)
                st.rerun()

    with tab_c:
        st.caption("送迎表Excelの各セルの色を設定できます。")
        colors = load_colors(館)
        updated = dict(colors)
        for group_name, keys in COLOR_GROUPS.items():
            st.subheader(group_name)
            cols_c = st.columns(len(keys))
            for i, key in enumerate(keys):
                label = COLOR_LABELS.get(key, key)
                val = colors.get(key, "#FFFFFF")
                updated[key] = cols_c[i].color_picker(label, value=val, key=f"cp_{key}")
        st.divider()
        c1c, c2c = st.columns(2)
        if c1c.button("💾 保存", type="primary", use_container_width=True, key="col_save"):
            save_colors(館, updated)
            st.success("✅ カラー設定を保存しました")
        if c2c.button("🔄 デフォルトに戻す", use_container_width=True, key="col_reset"):
            reset_colors(館)
            st.success("✅ デフォルト色に戻しました")
            st.rerun()


# ════ セッション初期化 ════════════════════════════════════════════
if "routes" not in st.session_state:
    st.session_state.routes = None

# ════ 認証 ═══════════════════════════════════════════════════════
if not _check_auth():
    st.stop()

# ════ サイドバー ══════════════════════════════════════════════════
_kan_from_url = st.query_params.get("kan", "Ⅴ番館")
if _kan_from_url not in HALLS:
    _kan_from_url = "Ⅴ番館"

with st.sidebar:
    st.markdown("## 🚐 送迎表ツール")
    st.caption("v1.0")
    館 = st.selectbox("事業所", HALLS,
                      index=HALLS.index(_kan_from_url), key="館_sel")
    try:
        st.query_params["kan"] = 館
    except Exception:
        pass
    st.divider()
    if is_gsheet_configured():
        st.caption("☁️ Supabase連携中")
    nav = st.radio("", [
        "📅 当日入力・送迎表",
        "👥 利用者マスタ管理",
        "⚙️ 車両・スタッフ・色設定",
    ], label_visibility="collapsed", key="nav_sel")

# ════ ルーティング ════════════════════════════════════════════════
if nav == "📅 当日入力・送迎表":
    page_daily(館)
elif nav == "👥 利用者マスタ管理":
    page_master(館)
elif nav == "⚙️ 車両・スタッフ・色設定":
    page_settings(館)
