"""
送迎表作成ツール v0.7
- 運転者を便・車両ごとに手動設定できるUI
- Supabase永続化（storage.py経由）
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
                    import_from_internal_csv, is_temp_juki_no,
                    get_facilities_needing_address, lookup_address_google,
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
**【STEP 1】** 対象日を選ぶ → 曜日が自動表示
**【STEP 2】** 固定利用者は上部に自動ソート。休校チェックで迎え先が「自宅」に切替
**【STEP 3】** 迎え先・送り先はプルダウンで選択、時刻も確認・修正
**【STEP 4】** スタッフ勤務状況を確認・修正
**【STEP 5】**「送迎ルートを自動生成」ボタンを押す
**【STEP 6】** 生成されたルートを確認・修正
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

    wday       = WEEKDAYS[target_date.weekday()]
    place_col  = "迎え先（長期休み）" if is_long else "迎え先（平日）"
    time_col_p = "迎え時刻（長期休み）" if is_long else "迎え時刻（平日）"

    # ── ソート：固定利用者を上部に ──
    df_work = master_df.copy()
    df_work["_fixed"] = df_work["利用曜日"].apply(lambda x: wday in str(x))
    df_sorted = pd.concat([
        df_work[df_work["_fixed"]],
        df_work[~df_work["_fixed"]],
    ], ignore_index=True)

    # ── 迎え先・送り先の選択肢 ──
    def _build_opts(*cols):
        s = set()
        for c in cols:
            if c in master_df.columns:
                for v in master_df[c].dropna():
                    v = str(v).strip()
                    if v and v not in ("None", "なし", ""):
                        s.add(v)
        return ["自宅", "なし"] + sorted(s - {"自宅"})

    pickup_opts_base  = _build_opts("迎え先（平日）", "迎え先（長期休み）")
    dropoff_opts_base = _build_opts("送り先")

    # ── 検索フィルター ──
    st.subheader("参加者チェック")
    col_srch, col_cnt = st.columns([4, 1])
    search_q = col_srch.text_input(
        "🔍", placeholder="氏名・地区で絞り込み",
        key="daily_search", label_visibility="collapsed")
    fixed_count = int(df_work["_fixed"].sum())
    col_cnt.metric("固定", f"{fixed_count}名")

    df_display = df_sorted.copy()
    if search_q:
        mask = (
            df_display["氏名"].str.contains(search_q, na=False, case=False) |
            df_display.get("地区", pd.Series("", index=df_display.index))
            .str.contains(search_q, na=False, case=False)
        )
        df_display = df_display[mask]

    st.caption("💡 固定曜日利用者が上部に表示 / 休校チェックで迎え先が自宅になります")

    attendance          = {}
    time_overrides      = {}
    send_time_overrides = {}
    pickup_overrides    = {}
    dropoff_overrides   = {}

    with st.container(border=True):
        hcols = st.columns([0.04, 0.13, 0.06, 0.16, 0.09, 0.05, 0.16, 0.09, 0.06, 0.10])
        for c, h in zip(hcols,
                        ["参加","氏名","区分","迎え先","迎え時刻",
                         "地区","送り先","送り時刻","休校","備考"]):
            c.markdown(f"**{h}**")

        prev_fixed = None
        for pos, row in df_display.iterrows():
            is_fixed = bool(row.get("_fixed", False))

            # グループ境界線
            if prev_fixed is True and not is_fixed:
                st.markdown("---")
                st.caption("⬇️ 固定曜日外（追加参加の場合はチェック）")
            elif prev_fixed is not None:
                st.markdown(
                    '<hr style="border:none;border-top:1px solid #EBEBEB;margin:2px 0">',
                    unsafe_allow_html=True)

            place_def      = str(row.get(place_col, "") or "自宅").strip() or "自宅"
            t_def          = _str_to_time(row.get(time_col_p, "15:00"))
            send_place_def = str(row.get("送り先", "") or "自宅").strip() or "自宅"
            send_t_def     = _str_to_time(row.get("送り時刻", "17:00"), default="17:00")
            ika            = str(row.get("医ケア", "")).strip()

            # 現在値が選択肢にない場合は追加
            pickup_opts = list(pickup_opts_base)
            if place_def not in pickup_opts:
                pickup_opts.insert(2, place_def)
            dropoff_opts = list(dropoff_opts_base)
            if send_place_def not in dropoff_opts:
                dropoff_opts.insert(2, send_place_def)

            rc = st.columns([0.04, 0.13, 0.06, 0.16, 0.09, 0.05, 0.16, 0.09, 0.06, 0.10])

            attend = rc[0].checkbox("", value=is_fixed,
                                    key=f"att_{pos}", label_visibility="collapsed")

            name_str = str(row.get("氏名", ""))
            if ika:
                rc[1].markdown(f"⚕️ **{name_str}**", help=f"医療的ケア: {ika}")
            else:
                rc[1].write(name_str)

            ku = str(row.get("区分", ""))
            badge_color = "#D9E1F2" if ku == "児発" else "#E2EFDA"
            rc[2].markdown(
                f'<span style="background:{badge_color};padding:2px 4px;'
                f'border-radius:3px;font-size:11px">{ku}</span>',
                unsafe_allow_html=True)

            # 休校チェック（迎え先を自宅に強制）
            school_hol = rc[8].checkbox("", key=f"hol_{pos}",
                                         label_visibility="collapsed", help="学校が休校")

            if school_hol:
                rc[3].caption("🏠 自宅（休校）")
                effective_pickup = "自宅"
            else:
                try:
                    pick_idx = pickup_opts.index(place_def)
                except ValueError:
                    pick_idx = 0
                effective_pickup = rc[3].selectbox(
                    "", options=pickup_opts, index=pick_idx,
                    key=f"pick_{pos}", label_visibility="collapsed")

            t_new = rc[4].time_input("", value=t_def, key=f"t_{pos}",
                                      step=300, label_visibility="collapsed")

            rc[5].caption(str(row.get("地区", "")))

            try:
                drop_idx = dropoff_opts.index(send_place_def)
            except ValueError:
                drop_idx = 0
            effective_dropoff = rc[6].selectbox(
                "", options=dropoff_opts, index=drop_idx,
                key=f"drop_{pos}", label_visibility="collapsed")

            send_t_new = rc[7].time_input("", value=send_t_def, key=f"st_{pos}",
                                           step=300, label_visibility="collapsed")

            note = rc[9].text_input("", key=f"note_{pos}",
                                     label_visibility="collapsed", placeholder="メモ")

            prev_fixed                = is_fixed
            attendance[pos]           = attend
            time_overrides[pos]       = str(t_new)[:5]
            send_time_overrides[pos]  = str(send_t_new)[:5]
            pickup_overrides[pos]     = effective_pickup
            dropoff_overrides[pos]    = effective_dropoff

    # ── スタッフ ──
    st.subheader("本日のスタッフ")
    staff_data = load_staff(館)
    staff_on_duty = {}
    if staff_data:
        cols4 = st.columns(min(len(staff_data), 5))
        for i, s in enumerate(staff_data):
            drive = s.get("運転可", False)
            icon = "🚗" if drive else "👤"
            on = cols4[i % 5].checkbox(f"{icon} {s['氏名']}", value=True, key=f"stf_{i}")
            staff_on_duty[s["氏名"]] = {
                "on": on, "drive": drive and on,
                "work_time": s.get("勤務時間", "9:30-18:30")}
    else:
        st.info("スタッフ未登録 →「車両・スタッフ設定」で登録してください")

    st.divider()
    if st.button("🚐 送迎ルートを自動生成", type="primary", use_container_width=True):
        attending_rows = []
        for pos, row in df_display.iterrows():
            if not attendance.get(pos, False):
                continue
            r = {c: row.get(c, "") for c in master_df.columns}
            r[time_col_p] = time_overrides.get(pos, r.get(time_col_p, "15:00"))
            r["送り時刻"] = send_time_overrides.get(pos, r.get("送り時刻", "17:00"))
            r[place_col]  = pickup_overrides.get(pos, r.get(place_col, "自宅"))
            r["送り先"]   = dropoff_overrides.get(pos, r.get("送り先", "自宅"))
            attending_rows.append(r)

        if not attending_rows:
            st.error("参加者が選択されていません")
            return

        parts = pd.DataFrame(attending_rows)
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

        sod = st.session_state.get("staff_on_duty", {})
        drivers_on_duty = [n for n, i in sod.items() if i.get("drive")]

        edited = {}
        for vehicle, trips in routes.items():
            with st.container(border=True):
                col_vh, col_drv = st.columns([3, 2])
                col_vh.markdown(f"#### 🚐 {vehicle}")

                drv_opts = ["（便ごとに設定）"] + drivers_on_duty
                veh_driver = col_drv.selectbox(
                    "担当ドライバー（全便に一括適用）",
                    drv_opts,
                    key=f"veh_drv_{vehicle}",
                )

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

# ════ ページ: 利用者マスタ ════════════════════════════════════════
def page_master(館):
    with st.expander("📖 使い方ガイド（クリックで開く）", expanded=False):
        st.markdown("""
**【利用者の追加】**「手動編集」タブ → 表の一番下の行に入力 →「保存」ボタン
**【CSV一括取込】**「内部CSV取込」タブ → 利用者一覧CSVをアップロード → 番館別に自動振り分け
**【住所検索】**「住所検索」タブ → 施設名からGoogle Mapsで住所を自動取得
**【変更履歴】**「変更履歴」タブ → 過去の状態に戻すことができます

⚠️ **迎え先（平日）**: 通常の下校先（学校名など）
⚠️ **迎え先（長期休み）**: 夏休み等は自宅になる場合が多いです
⚠️ **利用曜日**: 月水金 のように続けて入力（スペースなし）
        """)

    st.header(f"👥 利用者マスタ管理 — {館}")
    if is_gsheet_configured():
        st.info("☁️ Supabaseに自動バックアップされています", icon="✅")

    tab_edit, tab_csv, tab_ritalico, tab_addr, tab_hist = st.tabs([
        "✏️ 手動編集", "📋 内部CSV取込", "📂 リタリコCSV", "🗺️ 住所検索", "🕐 変更履歴",
    ])

    # ── 手動編集 ─────────────────────────────────────────────────────────
    with tab_edit:
        master_df = load_master(館)
        if master_df.empty:
            master_df = pd.DataFrame(columns=MASTER_COLUMNS)

        st.caption(f"登録人数: {len(master_df)}名")

        # 医ケア・重心バッジ
        if "医ケア" in master_df.columns:
            ika = master_df[master_df["医ケア"].str.strip().ne("")]
            if len(ika) > 0:
                st.warning(f"⚕️ 医療的ケア対象: {len(ika)}名", icon="⚕️")

        st.info("💡 利用曜日は「月水金」のように漢字1文字で続けて入力してください。")

        edited = st.data_editor(
            master_df, num_rows="dynamic", use_container_width=True,
            column_config={
                "受給者証番号":        st.column_config.TextColumn("受給者証番号"),
                "区分":               st.column_config.SelectboxColumn(
                                         "区分", options=["放デイ", "児発"]),
                "医ケア":             st.column_config.SelectboxColumn(
                                         "医ケア", options=["", "医ケア1", "医ケア2", "医ケア3"]),
                "重心":               st.column_config.TextColumn("重心"),
                "迎え時刻（平日）":    st.column_config.TextColumn("迎え時刻(平日) HH:MM"),
                "迎え時刻（長期休み）": st.column_config.TextColumn("迎え時刻(長休) HH:MM"),
                "送り時刻":            st.column_config.TextColumn("送り時刻 HH:MM"),
                "利用曜日":            st.column_config.TextColumn("利用曜日(例:月水金)"),
                "契約上限":            st.column_config.NumberColumn("契約上限", min_value=0),
                "契約月":              st.column_config.NumberColumn("契約月", min_value=0),
                "特記事項":            st.column_config.TextColumn("特記事項"),
                "備考":               st.column_config.TextColumn("備考"),
            },
            hide_index=False,
        )

        if st.button("💾 保存", type="primary", use_container_width=False):
            save_master(館, edited)
            st.success(f"✅ {len(edited)}件を保存しました")
            st.rerun()

    # ── 内部CSV取込 ───────────────────────────────────────────────────
    with tab_csv:
        st.subheader("内部利用者一覧CSV を取込")
        st.info("""
- 受給者証番号を主キーとして管理します
- 曜日列の `2` は自動的に `Ⅱ番館` に変換されます
- 各番館ごとに利用曜日が自動設定されます
- 児発 / 放デイ を自動識別します
- 受給者証番号が仮番号（99・申請中など）は⚠️で表示します
        """)

        uploaded_csv = st.file_uploader("利用者一覧CSVを選択", type=["csv"], key="internal_csv")
        if uploaded_csv:
            raw_bytes = uploaded_csv.read()
            raw_df = None
            for enc in ("utf-8-sig", "shift-jis", "cp932", "utf-8"):
                try:
                    import io as _io
                    raw_df = pd.read_csv(_io.BytesIO(raw_bytes), encoding=enc, dtype=str)
                    st.caption(f"文字コード: {enc} で読込成功")
                    break
                except Exception:
                    continue
            if raw_df is None:
                st.error("文字コードを判定できませんでした。")
                st.stop()

            st.write(f"**読込件数: {len(raw_df)}名**")
            st.dataframe(raw_df.head(5), use_container_width=True)

            # 番館別振り分けプレビュー
            hall_dfs = import_from_internal_csv(raw_df)
            st.subheader("番館別 振り分けプレビュー")
            cols4 = st.columns(4)
            for i, h in enumerate(["Ⅰ番館", "Ⅱ番館", "Ⅲ番館", "Ⅴ番館"]):
                cols4[i].metric(h, f"{len(hall_dfs.get(h, []))}名")

            # 仮番号警告
            all_rows = pd.concat(
                [d for d in hall_dfs.values() if not d.empty], ignore_index=True
            )
            if not all_rows.empty and "受給者証番号" in all_rows.columns:
                all_unique = all_rows.drop_duplicates("受給者証番号")
                temp_mask = all_unique["受給者証番号"].apply(is_temp_juki_no)
                temp_users = all_unique[temp_mask]
                if len(temp_users):
                    with st.expander(f"⚠️ 受給者証番号が未確定: {len(temp_users)}名"):
                        st.dataframe(
                            temp_users[["氏名", "受給者証番号", "状態", "備考"]],
                            use_container_width=True,
                        )

            # 現番館プレビュー
            df_cur = hall_dfs.get(館, pd.DataFrame())
            if not df_cur.empty:
                st.subheader(f"{館} の取込データ（{len(df_cur)}名）")
                preview_cols = ["氏名", "区分", "利用曜日", "迎え先（平日）",
                                "送り先", "医ケア", "受給者証番号"]
                st.dataframe(df_cur[[c for c in preview_cols if c in df_cur.columns]],
                             use_container_width=True)

            st.divider()
            c_left, c_right = st.columns(2)
            import_target = c_left.radio(
                "取込対象",
                [f"現在の番館のみ（{館}）", "全番館（Ⅰ〜Ⅴ）"],
                key="import_target",
            )
            merge_mode = c_right.radio(
                "既存データとの統合",
                ["上書き（全件置換）", "追加・更新（受給者証番号で照合）"],
                key="merge_mode",
            )

            if st.button("📥 取込・保存", type="primary"):
                targets = [館] if "現在" in import_target else ["Ⅰ番館", "Ⅱ番館", "Ⅲ番館", "Ⅴ番館"]
                saved_total = 0
                for h in targets:
                    new_df = hall_dfs.get(h, pd.DataFrame())
                    if new_df.empty:
                        continue
                    if "追加・更新" in merge_mode:
                        existing = load_master(h)
                        if (not existing.empty and "受給者証番号" in existing.columns):
                            existing_ids = set(existing["受給者証番号"].str.strip())
                            # 既存行を更新
                            merged = existing.copy()
                            for _, row in new_df.iterrows():
                                jid = str(row.get("受給者証番号", "")).strip()
                                idx_list = merged.index[
                                    merged["受給者証番号"].str.strip() == jid
                                ].tolist()
                                if idx_list:
                                    merged.loc[idx_list[0]] = row
                                else:
                                    merged = pd.concat(
                                        [merged, row.to_frame().T], ignore_index=True
                                    )
                            save_master(h, merged)
                        else:
                            save_master(h, new_df)
                    else:
                        save_master(h, new_df)
                    saved_total += len(new_df)
                st.success(f"✅ {saved_total}件を保存しました（{', '.join(targets)}）")
                st.rerun()

    # ── リタリコCSV ───────────────────────────────────────
    with tab_ritalico:
        st.subheader("リタリコCSVをインポート")
        st.info("リタリコから「保護者一覧」などのCSVをエクスポートしてアップロードしてください。")
        uploaded = st.file_uploader("CSVファイルを選択", type=["csv"], key="ritalico_csv")
        if uploaded:
            raw_bytes = uploaded.read()
            raw_df = None
            for enc in ("utf-8-sig", "shift-jis", "cp932", "utf-8"):
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

            st.dataframe(raw_df.head(), use_container_width=True)
            cols_csv = ["（未選択）"] + list(raw_df.columns)

            def _g(*cands):
                for c in cands:
                    if c in raw_df.columns:
                        return c
                return "（未選択）"

            st.subheader("列マッピング")
            c1, c2 = st.columns(2)
            col_name  = c1.selectbox("氏名列（必須）", cols_csv,
                index=cols_csv.index(_g("児童", "氏名", "利用者名", "お子様名")), key="m_name")
            col_kana  = c2.selectbox("フリガナ列", cols_csv,
                index=cols_csv.index(_g("児童カナ", "フリガナ", "保護者（カナ）")), key="m_kana")
            col_addr  = st.selectbox("住所列", cols_csv, index=0, key="m_addr")
            c3, c4, c5 = st.columns(3)
            col_school = c3.selectbox("迎え先列", cols_csv,
                index=cols_csv.index(_g("学校名", "通学先", "迎え先")), key="m_school")
            col_type  = c4.selectbox("区分列", cols_csv,
                index=cols_csv.index(_g("サービス種別", "区分", "サービス")), key="m_type")
            col_day   = c5.selectbox("利用曜日列", cols_csv,
                index=cols_csv.index(_g("利用曜日", "曜日")), key="m_day")

            if st.button("📥 取込・保存", type="primary", key="ritalico_save"):
                mapping = {"氏名": col_name, "フリガナ": col_kana, "住所": col_addr,
                           "迎え先（平日）": col_school, "区分": col_type, "利用曜日": col_day}
                result_df = import_from_ritalico(raw_df, mapping)
                save_master(館, result_df)
                st.success(f"✅ {len(result_df)}件を取込みました！")
                st.dataframe(result_df[["氏名", "地区", "迎え先（平日）", "区分", "利用曜日"]],
                             use_container_width=True)

    # ── 住所検索 ────────────────────────────────────────────
    with tab_addr:
        st.subheader("🗺️ 送迎先 住所検索")
        master_df = load_master(館)
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
                st.warning("""
⚠️ Google Maps APIキーが設定されていません。

Streamlit Cloud → Settings → Secrets に以下を追加してください：
```toml
[google_maps]
api_key = "AIza..."
```
APIキーなしでも住所を手動入力して保存できます。
                """)

            # 住所マップ（施設名 → 住所）をロード
            addr_map = load_json_data(館, "address_map", default={})

            st.write(f"**送迎先施設: {len(facilities)}件**")
            changed = False
            for facility in facilities:
                existing = addr_map.get(facility, "")
                row_a, row_b, row_c = st.columns([2, 3, 1])
                row_a.write(f"**{facility}**")

                new_addr = row_b.text_input(
                    "住所", value=existing,
                    key=f"addr_{facility}",
                    label_visibility="collapsed",
                    placeholder="住所を入力またはGoogle検索",
                )
                if new_addr != existing:
                    addr_map[facility] = new_addr
                    changed = True

                search_ok = bool(api_key)
                if row_c.button("🔍 検索", key=f"search_{facility}", disabled=not search_ok):
                    cands = lookup_address_google(facility, api_key)
                    if cands:
                        st.session_state[f"cands_{facility}"] = cands
                    else:
                        st.warning(f"「{facility}」の住所が見つかりませんでした")

                # 候補ボタン
                for cand in st.session_state.get(f"cands_{facility}", [])[:3]:
                    addr = cand.get("formatted_address", "")
                    name = cand.get("name", "")
                    if st.button(f"✅ {name} — {addr}", key=f"pick_{facility}_{addr}"):
                        addr_map[facility] = addr
                        # マスタの住所列も更新
                        updated = load_master(館)
                        for place_col, addr_col in (
                            ("迎え先（平日）", "迎え先住所"),
                            ("迎え先（長期休み）", "迎え先住所"),
                            ("送り先", "送り先住所"),
                        ):
                            if place_col in updated.columns and addr_col in updated.columns:
                                updated.loc[updated[place_col] == facility, addr_col] = addr
                        save_master(館, updated)
                        save_json_data(館, "address_map", addr_map)
                        del st.session_state[f"cands_{facility}"]
                        st.success(f"✅ 保存: {addr}")
                        st.rerun()

            if st.button("💾 住所マップを保存", type="primary"):
                save_json_data(館, "address_map", addr_map)
                st.success("✅ 保存しました")

    # ── 変更履歴 ────────────────────────────────────────────
    with tab_hist:
        st.subheader("変更履歴")
        hist_list = load_history_list(館)
        if not hist_list:
            st.info("まだ変更履歴がありません。保存するたびに自動で記録されます。")
        else:
            sel = st.selectbox(
                "履歴を選択", hist_list,
                format_func=lambda x: f"{x[:4]}/{x[4:6]}/{x[6:8]} {x[9:11]}:{x[11:13]}:{x[13:15]}",
            )
            hist_df = load_history(館, sel)
            st.dataframe(hist_df, use_container_width=True)
            if st.button("⏪ この時点に戻す", type="secondary"):
                save_master(館, hist_df)
                st.success("✅ 復元しました")
                st.rerun()



# ════ ページ: 車両・スタッフ ═══════════════════════════════════
def page_vehicles(館):
    with st.expander("📖 使い方ガイド（クリックで開く）", expanded=False):
        st.markdown("""
**【車両の追加・変更】**「車両」タブ → 表に入力 →「保存」ボタン
**【スタッフの追加】**「スタッフ」タブ → 氏名・運転可否・勤務時間を入力 →「保存」ボタン
⚠️「運転可」にチェックがあるスタッフのみ、当日入力画面で運転者候補として表示されます
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


# ════ ページ: カラー設定 ═══════════════════════════════
def page_colors(館):
    with st.expander("📖 使い方ガイド（クリックで開く）", expanded=False):
        st.markdown("""
**色の変更方法：** 各項目のカラーピッカーで色を選択 →「保存」ボタンを押すと反映されます
**リセット：**「デフォルトに戻す」ボタンで初期色に戻せます
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
            st.success("✅ カラー設定を保存しました。次回のExcel出力から反映されます。")
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
        st.success("☁️ Supabase連携 ON", icon="✅")
    else:
        st.warning("💾 ローカル保存（要設定）", icon="⚠️")
        with st.expander("Supabase設定方法"):
            st.markdown("""
1. [supabase.com](https://supabase.com) でプロジェクトを作成
2. SQL Editorで以下を実行:
```sql
CREATE TABLE app_data (
  key TEXT PRIMARY KEY,
  data JSONB NOT NULL,
  updated_at TIMESTAMPTZ DEFAULT now()
);
```
3. Settings → API → `Project URL` と `anon/public key` をコピー
4. Streamlit Cloud → App設定 → Secrets に追加:
```toml
[supabase]
url = "https://xxxx.supabase.co"
key = "eyJ..."
```
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
