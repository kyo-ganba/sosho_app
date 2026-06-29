"""
送迎表作成ツール v0.7
- 運転者を便・車両ごとに手動設定できるUI
- Google Sheets永続化（storage.py経由）
- データ保存状態をサイドバーに表示
"""
import streamlit as st
import pandas as pd
import json
from pathlib import Path
from datetime import datetime, time, date

from color_config import (load_colors, save_colors, reset_colors,
                          COLOR_LABELS, COLOR_GROUPS, DEFAULT_COLORS)
from master import (load_master, save_master, import_from_ritalico,
                    load_history_list, load_history, MASTER_COLUMNS)
from routing import generate_routes
from excel_export import export_schedule
from storage import load_json_data, save_json_data, is_gsheet_configured

st.set_page_config(page_title="送迎表ツール", page_icon="🚐", layout="wide")

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)

WEEKDAYS = ["月","火","水","木","金","土","日"]

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

# ════ ページ: 当日入力 ════════════════════════════════════════
def page_daily(館):
    with st.expander("📖 使い方ガイド（クリックで開く）", expanded=False):
        st.markdown("""
**【STEP 1】** 対象日を選ぶ → 曜日が自動表示されます
**【STEP 2】** 参加者チェック — その曜日の固定利用者に自動でチェックが入ります。欠席者はチェックを外し、追加参加者はチェックを入れてください
**【STEP 3】** 迎え時刻を確認・修正 — 当日の下校時刻が変わる場合はここで変更します
**【STEP 4】** 本日のスタッフ勤務状況を確認・修正
**【STEP 5】**「送迎ルートを自動生成」ボタンを押す
**【STEP 6】** 生成されたルートを確認・修正（車両ごとのドライバー設定も可能）
**【STEP 7】** 「Excelをダウンロード」
        """)

    st.header(f"📅 当日入力 — {館}")

    master_df = load_master(館)
    if master_df.empty:
        st.warning("⚠️ 利用者マスタが未登録です。左メニューの「利用者マスタ管理」でデータを登録してください。")
        return

    # ── 日付・モード選択 ──
    col_d, col_w, col_m = st.columns([2, 1, 1])
    with col_d:
        target_date = st.date_input("対象日", value=datetime.today())
    with col_w:
        st.metric("曜日", WEEKDAYS[target_date.weekday()])
    with col_m:
        is_long = st.toggle("長期休み", value=False, help="夏休み・冬休みなど長期休暇中はON")

    wday = WEEKDAYS[target_date.weekday()]

    # ── 参加者テーブル ──
    st.subheader("参加者チェック")
    st.caption("その曜日の固定利用者には自動でチェックが入っています。当日の変更に合わせて調整してください。")

    attendance = {}
    time_overrides = {}
    notes_map = {}

    with st.container(border=True):
        hcols = st.columns([0.05, 0.18, 0.1, 0.18, 0.17, 0.12, 0.12, 0.08])
        for c, h in zip(hcols, ["参加","氏名","区分","迎え先","迎え時刻","地区","送り先","備考"]):
            c.markdown(f"**{h}**")
        st.divider()

        for idx, row in master_df.iterrows():
            rdays = str(row.get("利用曜日",""))
            default_attend = wday in rdays if rdays else False

            place_col = "迎え先（長期休み）" if is_long else "迎え先（平日）"
            time_col  = "迎え時刻（長期休み）" if is_long else "迎え時刻（平日）"
            place = row.get(place_col,"") or row.get("住所","自宅") or "自宅"
            t_default = _str_to_time(row.get(time_col,"15:00"))

            rc = st.columns([0.05, 0.18, 0.1, 0.18, 0.17, 0.12, 0.12, 0.08])
            attend = rc[0].checkbox("", value=default_attend,
                                    key=f"att_{idx}", label_visibility="collapsed")
            rc[1].write(row.get("氏名",""))
            ku = row.get("区分","")
            rc[2].markdown(
                f'<span style="background:#D9E1F2;padding:2px 6px;border-radius:4px;font-size:12px">{ku}</span>'
                if ku == "児発" else
                f'<span style="background:#E2EFDA;padding:2px 6px;border-radius:4px;font-size:12px">{ku}</span>',
                unsafe_allow_html=True)
            rc[3].write(place)
            t_new = rc[4].time_input("", value=t_default, key=f"t_{idx}",
                                      label_visibility="collapsed", step=300)
            rc[5].write(row.get("地区",""))
            rc[6].write(row.get("送り先","") or row.get("住所","自宅") or "自宅")
            note = rc[7].text_input("", key=f"note_{idx}", label_visibility="collapsed",
                                     placeholder="メモ")

            attendance[idx] = attend
            time_overrides[idx] = str(t_new)
            notes_map[idx] = note

    # ── スタッフ ──
    st.subheader("本日のスタッフ")
    staff_data = load_staff(館)
    staff_on_duty = {}
    if staff_data:
        cols4 = st.columns(min(len(staff_data), 5))
        for i, s in enumerate(staff_data):
            drive = s.get("運転可", False)
            icon = "🚗" if drive else "👤"
            on = cols4[i % 5].checkbox(f"{icon} {s['氏名']}", value=True, key=f"st_{i}")
            staff_on_duty[s["氏名"]] = {
                "on": on, "drive": drive and on,
                "work_time": s.get("勤務時間", "9:30-18:30")}
    else:
        st.info("スタッフ未登録 →「車両・スタッフ設定」で登録してください")

    st.divider()
    if st.button("🚐 送迎ルートを自動生成", type="primary", use_container_width=True):
        flags = [attendance.get(i, False) for i in range(len(master_df))]
        parts = master_df[flags].copy()
        if parts.empty:
            st.error("参加者が選択されていません")
            return
        time_col = "迎え時刻（長期休み）" if is_long else "迎え時刻（平日）"
        for orig_idx, t_str in time_overrides.items():
            if orig_idx in parts.index:
                parts.loc[orig_idx, time_col] = t_str
        vehicles = load_vehicles(館)
        drivers = [n for n, i in staff_on_duty.items() if i.get("drive")]
        routes = generate_routes(parts, vehicles, drivers, is_long_holiday=is_long)
        st.session_state.routes = routes
        st.session_state.target_date = target_date
        st.session_state.staff_on_duty = staff_on_duty
        st.session_state["selected_kan"] = 館
        st.success("✅ ルートを生成しました！下にスクロールして確認・修正してください。")

    # ── ルート確認・修正 ──
    if st.session_state.get("routes"):
        st.divider()
        st.subheader("🗺️ 生成ルート（確認・修正）")
        routes = st.session_state.routes

        # 運転可能なスタッフ一覧
        sod = st.session_state.get("staff_on_duty", {})
        drivers_on_duty = [n for n, i in sod.items() if i.get("drive")]

        edited = {}
        for vehicle, trips in routes.items():
            with st.container(border=True):
                col_vh, col_drv = st.columns([3, 2])
                col_vh.markdown(f"#### 🚐 {vehicle}")

                # 車両ごとのデフォルトドライバー（全便一括設定）
                drv_opts = ["（便ごとに設定）"] + drivers_on_duty
                veh_driver = col_drv.selectbox(
                    "担当ドライバー（全便に一括適用）",
                    drv_opts,
                    key=f"veh_drv_{vehicle}",
                )

                # 一括適用の場合は全tripsのdriverを上書き
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

        st.divider()
        td = st.session_state.target_date
        fname = f"{td.strftime('%Y%m%d')}_{館}送迎表.xlsx"
        excel_bytes = export_schedule(
            edited, td, 館,
            st.session_state.get("staff_on_duty", {}),
            master_df=load_master(館),
            colors=load_colors(館))
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                "📥 Excelをダウンロード", data=excel_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, type="primary")
        with col2:
            if st.button("🔄 再生成", use_container_width=True):
                st.session_state.routes = None; st.rerun()


def _trip_editor(trips, vehicle, trip_type, drivers=None):
    """便リストの編集UI。drivers が指定されている場合はドライバーをselectboxで選択。"""
    if not trips:
        st.caption("（なし）")
        return []
    edited = []
    driver_opts = [""] + (drivers or [])
    for i, trip in enumerate(trips):
        c1, c2, c3, c4 = st.columns([0.15, 0.25, 0.3, 0.3])
        t = c1.time_input("", value=_str_to_time(trip.get("time","15:00")),
                          key=f"{vehicle}_{trip_type}_{i}_t", step=300,
                          label_visibility="collapsed")
        name  = c2.text_input("", value=trip.get("name",""),
                               key=f"{vehicle}_{trip_type}_{i}_n",
                               label_visibility="collapsed")
        place = c3.text_input("", value=trip.get("place",""),
                               key=f"{vehicle}_{trip_type}_{i}_p",
                               label_visibility="collapsed")
        cur_drv = trip.get("driver", "")
        if drivers:
            try:
                drv_idx = driver_opts.index(cur_drv)
            except ValueError:
                drv_idx = 0
            driver = c4.selectbox("", options=driver_opts, index=drv_idx,
                                   key=f"{vehicle}_{trip_type}_{i}_d",
                                   label_visibility="collapsed")
        else:
            driver = c4.text_input("", value=cur_drv,
                                    key=f"{vehicle}_{trip_type}_{i}_d",
                                    label_visibility="collapsed")
        edited.append({"type":trip_type,"time":str(t)[:5],"name":name,
                        "place":place,"driver":driver})
    return edited


# ════ ページ: 利用者マスタ ════════════════════════════════════
def page_master(館):
    with st.expander("📖 使い方ガイド（クリックで開く）", expanded=False):
        st.markdown("""
**【利用者の追加】**「手動編集」タブ → 表の一番下の行に入力 →「保存」ボタン
**【利用者の削除】**「手動編集」タブ → 削除したい行の左端のチェックボックスを選択 → ゴミ箱アイコン →「保存」ボタン
**【一括取込】**「リタリコCSV取込」タブ → CSVをアップロード → 列マッピングを確認 →「取込・保存」
**【変更履歴】**「変更履歴」タブ → 過去の状態に戻すことができます

⚠️ **迎え先（平日）**: 通常の下校先（学校名など）
⚠️ **迎え先（長期休み）**: 夏休み等は自宅になる場合が多いです
⚠️ **利用曜日**: 月水金 のように続けて入力（スペースなし）
        """)

    st.header(f"👥 利用者マスタ管理 — {館}")
    if is_gsheet_configured():
        st.info("☁️ Google Sheetsに自動バックアップされています", icon="✅")

    tab_edit, tab_import, tab_hist = st.tabs(["✏️ 手動編集", "📂 リタリコCSV取込", "🕐 変更履歴"])

    with tab_edit:
        master_df = load_master(館)
        if master_df.empty:
            master_df = pd.DataFrame(columns=MASTER_COLUMNS)

        st.caption(f"登録人数: {len(master_df)}名")
        st.info("💡 利用曜日は「月水金」のように漢字1文字で続けて入力してください。当日入力画面でその曜日の利用者に自動チェックが入ります。")

        edited = st.data_editor(
            master_df, num_rows="dynamic", use_container_width=True,
            column_config={
                "区分": st.column_config.SelectboxColumn(
                    "区分", options=["放デイ","児発"], required=False),
                "迎え時刻（平日）":    st.column_config.TextColumn("迎え時刻(平日) HH:MM"),
                "迎え時刻（長期休み）": st.column_config.TextColumn("迎え時刻(長休) HH:MM"),
                "送り時刻":            st.column_config.TextColumn("送り時刻 HH:MM"),
                "利用曜日":            st.column_config.TextColumn("利用曜日(例:月水金)"),
                "特記事項":           st.column_config.TextColumn("特記事項（連絡先・引き渡しルールなど）"),
                "備考":               st.column_config.TextColumn("備考"),
            },
            hide_index=False,
        )

        col1, col2 = st.columns([1,3])
        with col1:
            if st.button("💾 保存", type="primary", use_container_width=True):
                save_master(館, edited)
                st.success(f"✅ {len(edited)}件を保存しました")
                st.rerun()

    with tab_import:
        st.subheader("リタリコCSVをインポート")
        st.info("リタリコから「保護者一覧」などのCSVをエクスポートしてアップロードしてください。")
        uploaded = st.file_uploader("CSVファイルを選択", type=["csv"], key="ritalico_csv")
        if uploaded:
            raw_bytes = uploaded.read()
            raw_df = None
            for enc in ["utf-8-sig","shift-jis","cp932","utf-8"]:
                try:
                    import io as _io
                    raw_df = pd.read_csv(_io.BytesIO(raw_bytes), encoding=enc)
                    st.caption(f"文字コード: {enc} で読込成功")
                    break
                except Exception:
                    continue
            if raw_df is None:
                st.error("文字コードを判定できませんでした。UTF-8またはShift-JISで保存し直してください。")
                st.stop()

            st.write("**プレビュー（先頭5行）**")
            st.dataframe(raw_df.head(), use_container_width=True)

            cols_csv = ["（未選択）"] + list(raw_df.columns)
            def _g(*cands):
                for c in cands:
                    if c in raw_df.columns: return c
                return "（未選択）"

            st.subheader("列マッピング")
            st.caption("「氏名列」と「住所（市区町村・番地など複数列）」を選ぶだけで取込めます。他は後で手動入力できます。")

            c1,c2 = st.columns(2)
            col_name = c1.selectbox("氏名列（必須）", cols_csv,
                index=cols_csv.index(_g("児童","氏名","利用者名","お子様名")),
                key="m_name")
            col_kana = c2.selectbox("フリガナ列", cols_csv,
                index=cols_csv.index(_g("児童カナ","フリガナ","保護者（カナ）")),
                key="m_kana")

            col_addr = st.selectbox("住所列（単一列の場合のみ・複数列なら未選択でOK）",
                cols_csv, index=0, key="m_addr")

            c3,c4,c5 = st.columns(3)
            col_school = c3.selectbox("学校名/迎え先列", cols_csv,
                index=cols_csv.index(_g("学校名","学校","通学先","迎え先")), key="m_school")
            col_type = c4.selectbox("区分列（放デイ/児発）", cols_csv,
                index=cols_csv.index(_g("サービス種別","区分","サービス")), key="m_type")
            col_day = c5.selectbox("利用曜日列", cols_csv,
                index=cols_csv.index(_g("利用曜日","曜日")), key="m_day")

            if st.button("📥 取込・保存", type="primary"):
                mapping = {"氏名":col_name,"フリガナ":col_kana,"住所":col_addr,
                           "迎え先（平日）":col_school,"区分":col_type,"利用曜日":col_day}
                result_df = import_from_ritalico(raw_df, mapping)
                save_master(館, result_df)
                st.success(f"✅ {len(result_df)}件を取込みました！「手動編集」タブで迎え時刻・区分を追記してください。")
                st.dataframe(result_df[["氏名","地区","迎え先（平日）","区分","利用曜日"]],
                             use_container_width=True)

    with tab_hist:
        st.subheader("変更履歴")
        hist_list = load_history_list(館)
        if not hist_list:
            st.info("まだ変更履歴がありません。保存するたびに自動で記録されます。")
        else:
            sel = st.selectbox("履歴を選択", hist_list,
                format_func=lambda x: f"{x[:4]}/{x[4:6]}/{x[6:8]} {x[9:11]}:{x[11:13]}:{x[13:15]}")
            hist_df = load_history(館, sel)
            st.dataframe(hist_df, use_container_width=True)
            if st.button("⏪ この時点に戻す", type="secondary"):
                save_master(館, hist_df)
                st.success("✅ 復元しました")
                st.rerun()


# ════ ページ: 車両・スタッフ ═════════════════════════════════
def page_vehicles(館):
    with st.expander("📖 使い方ガイド（クリックで開く）", expanded=False):
        st.markdown("""
**【車両の追加・変更】**「車両」タブ → 表に入力 →「保存」ボタン
**【スタッフの追加】**「スタッフ」タブ → 氏名・運転可否・勤務時間を入力 →「保存」ボタン
"��️「運転可」にチェックがあるスタッフのみ、当日入力画面で運転者候補として表示されます
        """)

    st.header(f"🚗 車両・スタッフ設定 — {館}")
    tab_v, tab_s = st.tabs(["🚐 車両", "👤 スタッフ"])

    with tab_v:
        vdf = pd.DataFrame(load_vehicles(館))
        ev = st.data_editor(vdf, num_rows="dynamic", use_container_width=True,
                            column_config={"定員": st.column_config.NumberColumn("定員",min_value=1,max_value=20)})
        if st.button("💾 車両を保存", type="primary", key="sv"):
            save_vehicles(館, ev.to_dict("records")); st.success("✅ 保存しました")

    with tab_s:
        sdf = pd.DataFrame(load_staff(館)) if load_staff(館) else \
              pd.DataFrame(columns=["氏名","運転可","勤務時間","備考"])
        es = st.data_editor(sdf, num_rows="dynamic", use_container_width=True,
                            column_config={"運転可": st.column_config.CheckboxColumn("運転可")})
        if st.button("💾 スタッフを保存", type="primary", key="ss"):
            save_staff(館, es.to_dict("records")); st.success("✅ 保存しました")


# ════ ページ: カラー設定 ═══════════════════════════════════
def page_colors(館):
    with st.expander("📖 使い方ガイド（クリックで開く）", expanded=False):
        st.markdown("""
**色の変更方法：** 各項目のカラーピッカーで色を選択 →「保存」ボタンを押すと反映されます
**リシット：**「デフォルトに戻す」ボタンで初期色に戻せます
⚠️ 設定は館ごとに保存されます
        """)

    st.header(f"🎨 カラー設定 — {館}")
    st.caption("送迎表Excelの各セルの色を自由に設定できます。")

    colors = load_colors(館)
    updated = dict(colors)

    for group_name, keys in COLOR_GROUPS.items():
        st.subheader(group_name)
        cols = st.columns(len(keys))
        for i, key in enumerate(keys):
            label = COLOR_LABELS.get(key, key)
            val = colors.get(key, "#FFFFFF")
            picked = cols[i].color_picker(label, value=val, key=f"cp_{key}")
            updated[key] = picked

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        if st.button("💾 保存", type="primary", use_container_width=True):
            save_colors(館, updated)
            st.success("✅ カラー設定を保存しました。次回はExcel出力から反映されます。")
    with col2:
        if st.button("🔄 デフォルトに戻す", use_container_width=True):
            reset_colors(館)
            st.success("✅ デフォルト色に戻しました。")
            st.rerun()

    st.divider()
    st.subheader("プレビュー")
    st.caption("現在の色設定のイメージ")
    cols = st.columns(5)
    preview_items = [
        ("迎え（放デイ）", "迎え_放デイ_バッジ", "迎え_放デイ_氏名"),
        ("迎え（児発）",   "迎え_児発_バッジ",   "迎え_児発_氏名"),
        ("送り",           "送り_バッジ",         "送り_氏名"),
        ("特記事項",       "特記事項",            "特記事項"),
        ("急遽追加",       "急遽追加_バッジ",     "急遽追加_氏名"),
    ]
    for i, (label, badge_key, name_key) in enumerate(preview_items):
        badge_c = updated.get(badge_key, "#FFFFFF")
        name_c  = updated.get(name_key,  "#FFFFFF")
        cols[i].markdown(
            f'<div style="text-align:center;font-size:11px;margin-bottom:4px">{label}</div>'
            f'<div style="background:{badge_c};padding:4px;border-radius:4px 4px 0 0;'
            f'text-align:center;font-size:11px;font-weight:bold">バッジ</div>'
            f'<div style="background:{name_c};padding:4px;border-radius:0 0 4px 4px;'
            f'text-align:center;font-size:11px">氏名欄</div>',
            unsafe_allow_html=True)


# ════ セッション初期化 ════════════════════════════════════════
if "routes" not in st.session_state:
    st.session_state.routes = None

# ════ サイドバー ══════════════════════════════════════════════
with st.sidebar:
    st.title("🚐 送迎表ツール")
    st.caption("v0.7")
    館 = st.selectbox("事業所", ["Ⅰ番館","Ⅱ番館","Ⅲ番館","Ⅴ番館"], key="館")
    st.divider()

    # データ保存状態
    if is_gsheet_configured():
        st.success("☁️ Google Sheets連携 ON", icon="✅")
    else:
        st.warning("💾 ローカル保存（要設定）", icon="⚠️")
        with st.expander("Google Sheets設定方法"):
            st.markdown("""
1. Google Cloud Console でサービスアカウントを作成
2. Google Sheets API を有効化
3. スプレッドシートを作成してサービスアカウントと共有
4. Streamlit Cloud → App設定 → Secrets に認証情報を追加
            """)

    st.divider()
    nav = st.radio("", ["📅 当日入力・送迎表生成",
                        "👥 利用者マスタ管理",
                        "🚗 車両・スタッフ設定",
                        "🎨 カラー設定"],
                   label_visibility="collapsed")

# ════ ルーティング ════════════════════════════════════════════
if nav == "📅 当日入力・送迎表生成":
    page_daily(館)
elif nav == "👥 利用者マスタ管理":
    page_master(館)
elif nav == "🚗 車両・スタッフ設定":
    page_vehicles(館)
elif nav == "🎨 カラー設定":
    page_colors(館)
