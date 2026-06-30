import streamlit as st

st.set_page_config(page_title="送迎表ツール", page_icon="🚐", layout="wide")

st.write("## ✅ アプリが起動しています")
st.write("このメッセージが見えれば、Streamlitは正常に動作しています。")

_import_errors = []

try:
    import re, pandas as pd, json
    from pathlib import Path
    from datetime import datetime, time, date
    st.success("✅ 標準ライブラリ: OK")
except Exception as e:
    _import_errors.append(f"標準ライブラリ: {e}")

try:
    from color_config import load_colors, save_colors, reset_colors, COLOR_LABELS, COLOR_GROUPS, DEFAULT_COLORS
    st.success("✅ color_config: OK")
except Exception as e:
    _import_errors.append(f"color_config: {e}")
    st.error(f"❌ color_config: {e}")

try:
    from master import (load_master, save_master, import_from_ritalico, import_from_internal_csv,
                        is_temp_juki_no, get_facilities_needing_address, lookup_address_google,
                        import_address_from_hogosha_csv, load_history_list, load_history, MASTER_COLUMNS)
    st.success("✅ master: OK")
except Exception as e:
    _import_errors.append(f"master: {e}")
    st.error(f"❌ master: {e}")

try:
    from routing import generate_routes
    st.success("✅ routing: OK")
except Exception as e:
    _import_errors.append(f"routing: {e}")
    st.error(f"❌ routing: {e}")

try:
    from excel_export import export_schedule
    st.success("✅ excel_export: OK")
except Exception as e:
    _import_errors.append(f"excel_export: {e}")
    st.error(f"❌ excel_export: {e}")

try:
    from storage import load_json_data, save_json_data, is_gsheet_configured
    st.success("✅ storage: OK")
except Exception as e:
    _import_errors.append(f"storage: {e}")
    st.error(f"❌ storage: {e}")

if _import_errors:
    st.error("インポートエラーがあります。上記を確認してください。")
else:
    st.balloons()
    st.success("🎉 全モジュール正常！本番app.pyに戻しています...")
