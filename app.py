import streamlit as st
import pandas as pd
import json
from pathlib import Path
from datetime import datetime, time
import io

from master import load_master, save_master, import_from_ritalico
from routing import generate_routes
from excel_export import export_schedule

st.set_page_config(page_title="送迎表作成ツール", page_icon="🚐", layout="wide")

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)

# ─── セッション初期化 ───────────────────────────────────────
if "page" not in st.session_state:
    st.session_state.page = "daily"
if "routes" not in st.session_state:
    st.session_state.routes = None

# ─── サイドバーナビ ─────────────────────────────────────────
with st.sidebar:
    st.title("🚐 送迎表ツール")
    館 = st.selectbox("事業所", ["Ⅰ番館", "Ⅱ番館", "Ⅲ番館", "Ⅴ番館"], key="館")
    st.divider()
    nav = st.radio(
        "メニュー",
        ["📅 当日入力・送迎表生成", "👥 利用者マスタ管理", "🚗 車両・スタッフ設定"],
        label_visibility="collapsed",
    )
    st.divider()
    st.caption("v0.1 プロトタイプ")

# ─── ページ切り替え ─────────────────────────────────────────
if nav == "📅 当日入力・送迎表生成":
    page_daily(館)
elif nav == "👥 利用者マスタ管理":
    page_master(館)
elif nav == "🚗 車両・スタッフ設定":
    page_vehicles(館)


# ════════════════════════════════════════════════════════════
def page_daily(館):
    st.header(f"📅 当日入力 — {館}")

    master_df = load_master(館)
    if master_df.empty:
        st.warning("利用者マスタが未登録です。まず「利用者マスタ管理」でデータを登録してください。")
        return

    # 日付・曜日
    col_date, col_day = st.columns([2, 1])
    with col_date:
        target_date = st.date_input("対象日", value=datetime.today())
    with col_day:
        st.metric("曜日", ["月", "火", "水", "木", "金", "土", "日"][target_date.weekday()])

    st.subheader("参加者チェック")

    # 曜日ごとにデフォルト参加フラグを絞る
    weekday_map = {0: "月", 1: "火", 2: "水", 3: "木", 4: "金", 5: "土", 6: "日"}
    weekday_jp = weekday_map[target_date.weekday()]

    # 利用曜日列がある場合はフィルタ
    if "利用曜日" in master_df.columns:
        default_attend = master_df["利用曜日"].str.contains(weekday_jp, na=False)
    else:
        default_attend = pd.Series([True] * len(master_df))

    attendance = {}
    pickup_overrides = {}  # 下校時刻の当日変更

    with st.expander("▼ 参加者一覧（チェックで参加確定）", expanded=True):
        cols = st.columns([0.05, 0.2, 0.15, 0.2, 0.15, 0.15, 0.1])
        headers = ["参加", "氏名", "区分", "迎え先（学校等）", "迎え時刻", "地区", "備考"]
        for col, h in zip(cols, headers):
            col.markdown(f"**{h}**")

        for idx, row in master_df.iterrows():
            cols = st.columns([0.05, 0.2, 0.15, 0.2, 0.15, 0.15, 0.1])
            default = bool(default_attend.iloc[idx] if idx < len(default_attend) else True)
            attend = cols[0].checkbox("", value=default, key=f"att_{idx}", label_visibility="collapsed")
            cols[1].write(row.get("氏名", ""))
            cols[2].write(row.get("区分", ""))
            cols[3].write(row.get("迎え先", ""))

            # 下校時刻：当日上書き可能
            default_time = row.get("下校時刻", time(15, 0))
            if isinstance(default_time, str):
                try:
                    h_m = default_time.split(":")
                    default_time = time(int(h_m[0]), int(h_m[1]))
                except Exception:
                    default_time = time(15, 0)
            t = cols[4].time_input("", value=default_time, key=f"t_{idx}", label_visibility="collapsed", step=300)
            cols[5].write(row.get("地区", ""))
            note = cols[6].text_input("", key=f"note_{idx}", label_visibility="collapsed")

            attendance[idx] = attend
            pickup_overrides[idx] = {"time": t, "note": note}

    # 欠席・特記
    st.subheader("特記事項")
    special_notes = st.text_area("全体メモ（スタッフ不足・応援依頼など）", height=80)

    # 当日スタッフ
    st.subheader("本日のスタッフ勤務状況")
    staff_data = load_staff(館)
    staff_on_duty = {}
    if staff_data:
        s_cols = st.columns(4)
        for i, s in enumerate(staff_data):
            can_drive = s.get("運転可", False)
            icon = "🚗" if can_drive else "👤"
            on = s_cols[i % 4].checkbox(f"{icon} {s['氏名']}", value=True, key=f"staff_{i}")
            staff_on_duty[s["氏名"]] = {"on": on, "drive": can_drive and on}
    else:
        st.info("スタッフ情報が未登録です。「車両・スタッフ設定」から追加してください。")

    st.divider()
    if st.button("🚐 送迎ルートを自動生成", type="primary", use_container_width=True):
        participants = master_df[
            [attendance.get(i, False) for i in range(len(master_df))]
        ].copy()

        if participants.empty:
            st.error("参加者が選択されていません。")
            return

        # 下校時刻の上書き反映
        for orig_idx, new_vals in pickup_overrides.items():
            mask = participants.index == orig_idx
            if mask.any():
                participants.loc[mask, "下校時刻"] = str(new_vals["time"])
                participants.loc[mask, "備考"] = new_vals["note"]

        vehicles = load_vehicles(館)
        drivers = [name for name, info in staff_on_duty.items() if info.get("drive")]

        with st.spinner("ルートを計算中..."):
            routes = generate_routes(participants, vehicles, drivers)

        st.session_state.routes = routes
        st.session_state.target_date = target_date
        st.session_state.館 = 館
        st.success("ルートを生成しました！下にスクロールして確認・修正してください。")

    # ─── ルート表示・修正 ───────────────────────────────────
    if st.session_state.routes:
        st.divider()
        st.subheader("🗺️ 生成されたルート（確認・修正）")
        routes = st.session_state.routes

        edited_routes = {}
        for vehicle, trips in routes.items():
            st.markdown(f"### 🚐 {vehicle}")
            tab_mukae, tab_okuri = st.tabs(["迎え", "送り"])

            with tab_mukae:
                mukae = [t for t in trips if t["type"] == "迎え"]
                edited = _render_trip_editor(mukae, vehicle, "迎え", master_df)
                edited_routes.setdefault(vehicle, {})["迎え"] = edited

            with tab_okuri:
                okuri = [t for t in trips if t["type"] == "送り"]
                edited = _render_trip_editor(okuri, vehicle, "送り", master_df)
                edited_routes.setdefault(vehicle, {})["送り"] = edited

        st.divider()
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            if st.button("📋 送迎表Excelを出力", use_container_width=True, type="primary"):
                excel_bytes = export_schedule(
                    edited_routes,
                    st.session_state.target_date,
                    館,
                    staff_on_duty,
                )
                st.download_button(
                    "⬇️ ダウンロード",
                    data=excel_bytes,
                    file_name=f"送迎表_{館}_{target_date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        with col_dl2:
            if st.button("🔄 ルートを再生成", use_container_width=True):
                st.session_state.routes = None
                st.rerun()


def _render_trip_editor(trips, vehicle, trip_type, master_df):
    """迎え/送りのリストを表示・編集UIとして返す"""
    if not trips:
        st.caption("（なし）")
        return []

    edited = []
    for i, trip in enumerate(trips):
        with st.container():
            c1, c2, c3, c4 = st.columns([0.15, 0.25, 0.3, 0.3])
            t = c1.time_input(
                "時刻",
                value=_str_to_time(trip.get("time", "15:00")),
                key=f"{vehicle}_{trip_type}_{i}_t",
                step=300,
                label_visibility="collapsed",
            )
            name = c2.text_input(
                "氏名",
                value=trip.get("name", ""),
                key=f"{vehicle}_{trip_type}_{i}_n",
                label_visibility="collapsed",
            )
            place = c3.text_input(
                "場所",
                value=trip.get("place", ""),
                key=f"{vehicle}_{trip_type}_{i}_p",
                label_visibility="collapsed",
            )
            driver = c4.text_input(
                "担当",
                value=trip.get("driver", ""),
                key=f"{vehicle}_{trip_type}_{i}_d",
                label_visibility="collapsed",
            )
            edited.append({"type": trip_type, "time": str(t), "name": name, "place": place, "driver": driver})
    return edited


# ════════════════════════════════════════════════════════════
def page_master(館):
    st.header(f"👥 利用者マスタ管理 — {館}")

    tab_import, tab_edit = st.tabs(["リタリコCSV取込", "手動編集"])

    with tab_import:
        st.subheader("リタリコからCSVをインポート")
        st.info(
            "リタリコの利用者情報CSVをそのままアップロードしてください。\n"
            "氏名・住所・学校名などを自動マッピングします。"
        )
        uploaded = st.file_uploader("CSVファイルを選択", type=["csv"], key="ritalico_csv")
        if uploaded:
            raw_df = pd.read_csv(uploaded, encoding="utf-8-sig")
            st.write("**CSVプレビュー（先頭5行）**")
            st.dataframe(raw_df.head(), use_container_width=True)

            st.subheader("列マッピング設定")
            st.caption("リタリコCSVの列名 → アプリ内の項目に対応付けてください")
            cols_in_csv = ["（未選択）"] + list(raw_df.columns)
            map_cols = st.columns(3)
            col_name = map_cols[0].selectbox("氏名列", cols_in_csv, key="m_name")
            col_kana = map_cols[1].selectbox("フリガナ列", cols_in_csv, key="m_kana")
            col_addr = map_cols[2].selectbox("住所列", cols_in_csv, key="m_addr")
            map_cols2 = st.columns(3)
            col_school = map_cols2[0].selectbox("学校名列", cols_in_csv, key="m_school")
            col_type = map_cols2[1].selectbox("区分列（放デイ/児発）", cols_in_csv, key="m_type")
            col_day = map_cols2[2].selectbox("利用曜日列", cols_in_csv, key="m_day")

            if st.button("取込・保存", type="primary"):
                mapping = {
                    "氏名": col_name, "フリガナ": col_kana, "住所": col_addr,
                    "学校名": col_school, "区分": col_type, "利用曜日": col_day,
                }
                result_df = import_from_ritalico(raw_df, mapping)
                save_master(館, result_df)
                st.success(f"{len(result_df)}件を取り込みました。")
                st.dataframe(result_df, use_container_width=True)

    with tab_edit:
        st.subheader("利用者一覧・編集")
        master_df = load_master(館)

        if master_df.empty:
            st.info("データがありません。CSVから取込むか、下のボタンで手動追加してください。")
            master_df = pd.DataFrame(columns=[
                "氏名", "フリガナ", "区分", "地区", "住所", "迎え先", "下校時刻", "自宅時刻", "利用曜日", "備考"
            ])

        edited_df = st.data_editor(
            master_df,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "区分": st.column_config.SelectboxColumn("区分", options=["放デイ", "児発"]),
                "下校時刻": st.column_config.TextColumn("迎え時刻(HH:MM)"),
                "自宅時刻": st.column_config.TextColumn("送り時刻(HH:MM)"),
                "利用曜日": st.column_config.TextColumn("利用曜日(例:月水金)"),
            },
        )
        if st.button("💾 保存", type="primary"):
            save_master(館, edited_df)
            st.success("保存しました。")


# ════════════════════════════════════════════════════════════
def page_vehicles(館):
    st.header(f"🚗 車両・スタッフ設定 — {館}")

    tab_v, tab_s = st.tabs(["車両", "スタッフ"])

    with tab_v:
        st.subheader("車両一覧")
        vehicles = load_vehicles(館)
        v_df = pd.DataFrame(vehicles) if vehicles else pd.DataFrame(columns=["車両名", "定員", "備考"])
        edited_v = st.data_editor(v_df, num_rows="dynamic", use_container_width=True)
        if st.button("車両情報を保存", type="primary", key="save_v"):
            save_vehicles(館, edited_v.to_dict("records"))
            st.success("保存しました。")

    with tab_s:
        st.subheader("スタッフ一覧")
        staff = load_staff(館)
        s_df = pd.DataFrame(staff) if staff else pd.DataFrame(columns=["氏名", "運転可", "備考"])
        edited_s = st.data_editor(
            s_df,
            num_rows="dynamic",
            use_container_width=True,
            column_config={"運転可": st.column_config.CheckboxColumn("運転可")},
        )
        if st.button("スタッフ情報を保存", type="primary", key="save_s"):
            save_staff(館, edited_s.to_dict("records"))
            st.success("保存しました。")


# ─── データ永続化ヘルパー ────────────────────────────────────
def load_vehicles(館):
    p = DATA_DIR / f"{館}_vehicles.json"
    if p.exists():
        return json.loads(p.read_text(encoding="utf-8"))
    # デフォルト車両（Ⅴ番館のサンプルから）
    return [
        {"車両名": "セレナ", "定員": 7, "備考": ""},
        {"車両名": "ボクシー", "定員": 7, "備考": ""},
        {"車両名": "白フリード", "定員": 6, "備考": ""},
        {"車両名": "銀フリ地下", "定員": 6, "備考": ""},
        {"車両名": "銀フリ浅香", "定員": 6, "備考": ""},
    ]

def save_vehicles(館, data):
    p = DATA_DIR / f"{館}_vehicles.json"
    p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

def load_staff(館):
    p = DATA_DIR / f"{館}_staff.json"
    if p.exists():
        return json.loads(p.read_text(encoding="utf-8"))
    return []

def save_staff(館, data):
    p = DATA_DIR / f"{館}_staff.json"
    p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

def _str_to_time(s):
    try:
        h, m = s.split(":")
        return time(int(h), int(m))
    except Exception:
        return time(15, 0)
