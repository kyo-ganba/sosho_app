"""
データ永続化モジュール — Google Sheets + ローカルファイル二重保存

Google Sheets設定済み（st.secrets["gsheet"]）→ Sheets優先 + ローカルバックアップ
未設定 → ローカルのみ（Streamlit無料プランでは再起動時にリセットされる）
"""
import json
import pandas as pd
from pathlib import Path
from datetime import datetime

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)


# ── Streamlit / gspread 依存を安全にインポート ──────────────────
try:
    import streamlit as st

    @st.cache_resource(show_spinner=False)
    def _gsheet_resource():
        """
        gspreadクライアント と spreadsheet_id を返す（アプリ全体で1回だけ接続）。
        st.secrets["gsheet"] が未設定の場合は (None, None) を返す。
        """
        try:
            import gspread
            if "gsheet" not in st.secrets:
                return None, None
            cfg = dict(st.secrets["gsheet"])
            sid = cfg.pop("spreadsheet_id", None)
            if not sid:
                return None, None
            gc = gspread.service_account_from_dict(cfg)
            return gc, sid
        except Exception:
            return None, None

except Exception:
    # Streamlit 以外の環境（テスト等）では常にローカルモード
    def _gsheet_resource():  # type: ignore
        return None, None


def _open_spreadsheet():
    """Spreadsheet オブジェクトを返す。未設定・エラー時は None。"""
    gc, sid = _gsheet_resource()
    if gc is None:
        return None
    try:
        return gc.open_by_key(sid)
    except Exception:
        return None


def is_gsheet_configured() -> bool:
    """Google Sheets 連携が有効かどうか"""
    try:
        import streamlit as st
        return "gsheet" in st.secrets
    except Exception:
        return False


# ── DataFrame I/O ─────────────────────────────────────────────
def load_df(館: str, name: str, columns: list = None) -> pd.DataFrame:
    """
    DataFrame を読み込む。
    Sheets 設定済み → Google Sheets から読む（ローカルにフォールバック）
    未設定     → ローカル CSV から読む
    """
    sh = _open_spreadsheet()
    if sh:
        try:
            ws = sh.worksheet(f"{館}_{name}")
            records = ws.get_all_records(default_blank="")
            df = pd.DataFrame(records).fillna("").astype(str)
            if columns:
                for c in columns:
                    if c not in df.columns:
                        df[c] = ""
                return df[columns]
            return df
        except Exception:
            pass  # ローカルへフォールバック

    # ローカル CSV
    p = DATA_DIR / f"{館}_{name}.csv"
    if p.exists():
        df = pd.read_csv(p, encoding="utf-8-sig", dtype=str).fillna("")
        if columns:
            for c in columns:
                if c not in df.columns:
                    df[c] = ""
            return df[columns]
        return df
    return pd.DataFrame(columns=columns) if columns else pd.DataFrame()


def save_df(館: str, name: str, df: pd.DataFrame):
    """
    DataFrame を保存する（ローカル CSV に常に保存 + Sheets が設定済みなら Sheets にも）。
    """
    # 常にローカルに保存
    p = DATA_DIR / f"{館}_{name}.csv"
    df.to_csv(p, index=False, encoding="utf-8-sig")

    # Sheets にも保存
    sh = _open_spreadsheet()
    if sh:
        try:
            ws_name = f"{館}_{name}"
            try:
                ws = sh.worksheet(ws_name)
                ws.clear()
            except Exception:
                ws = sh.add_worksheet(
                    title=ws_name,
                    rows=max(len(df) + 10, 50),
                    cols=max(len(df.columns) + 2, 5),
                )
            data = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()
            ws.update(data)
        except Exception:
            pass  # ローカル保存は成功しているので無視


# ── JSON I/O ──────────────────────────────────────────────────
def load_json_data(館: str, name: str, default=None):
    """
    JSON データを読み込む（Sheets → ローカル JSON の順）。
    """
    sh = _open_spreadsheet()
    if sh:
        try:
            ws = sh.worksheet(f"{館}_{name}_json")
            raw = ws.acell("A1").value
            if raw:
                return json.loads(raw)
        except Exception:
            pass

    p = DATA_DIR / f"{館}_{name}.json"
    if p.exists():
        return json.loads(p.read_text(encoding="utf-8"))
    return default if default is not None else []


def save_json_data(館: str, name: str, data):
    """
    JSON データを保存する（ローカル + Sheets）。
    """
    # ローカルに常に保存
    p = DATA_DIR / f"{館}_{name}.json"
    p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    # Sheets にも保存
    sh = _open_spreadsheet()
    if sh:
        try:
            ws_name = f"{館}_{name}_json"
            try:
                ws = sh.worksheet(ws_name)
            except Exception:
                ws = sh.add_worksheet(title=ws_name, rows=5, cols=2)
            ws.update("A1", [[json.dumps(data, ensure_ascii=False)]])
        except Exception:
            pass


# ── 変更履歴（ローカルのみ） ─────────────────────────────────
def save_history(館: str, df: pd.DataFrame):
    """変更履歴をローカルに保存（最新20件を保持）"""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    hist_dir = DATA_DIR / "history" / 館
    hist_dir.mkdir(parents=True, exist_ok=True)
    df.to_csv(hist_dir / f"{ts}.csv", index=False, encoding="utf-8-sig")
    # 古いファイルを削除
    files = sorted(hist_dir.glob("*.csv"))
    for old in files[:-20]:
        old.unlink()


def load_history_list(館: str) -> list:
    """変更履歴のタイムスタンプ一覧（新しい順）"""
    hist_dir = DATA_DIR / "history" / 館
    if not hist_dir.exists():
        return []
    return sorted([f.stem for f in hist_dir.glob("*.csv")], reverse=True)


def load_history_df(館: str, timestamp: str) -> pd.DataFrame:
    """指定タイムスタンプの変更履歴を読み込む"""
    p = DATA_DIR / "history" / 館 / f"{timestamp}.csv"
    if p.exists():
        return pd.read_csv(p, encoding="utf-8-sig", dtype=str).fillna("")
    return pd.DataFrame()
