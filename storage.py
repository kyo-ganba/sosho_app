"""
データ永続化モジュール — Supabase + ローカルファイル二重保存

Supabase設定済み（st.secrets["supabase"]）→ Supabase優先 + ローカルバックアップ
未設定 → ローカルのみ（Streamlit無料プランでは再起動時にリセットされる）

Supabaseに事前作成が必要なテーブル（SQL）:
  CREATE TABLE app_data (
    key TEXT PRIMARY KEY,
    data JSONB NOT NULL,
    updated_at TIMESTAMPTZ DEFAULT now()
  );
"""
import json
import pandas as pd
from pathlib import Path
from datetime import datetime

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)


# ── Streamlit / supabase 依存を安全にインポート ──────────────────
try:
    import streamlit as st

    @st.cache_resource(show_spinner=False)
    def _supabase_client():
        """
        Supabaseクライアントを返す（アプリ全体で1回だけ接続）。
        st.secrets["supabase"] が未設定の場合は None を返す。
        """
        try:
            from supabase import create_client
            if "supabase" not in st.secrets:
                return None
            url = st.secrets["supabase"]["url"]
            key = st.secrets["supabase"]["key"]
            return create_client(url, key)
        except Exception:
            return None

except Exception:
    # Streamlit 以外の環境（テスト等）では常にローカルモード
    def _supabase_client():  # type: ignore
        return None


def is_supabase_configured() -> bool:
    """Supabase 連携が有効かどうか"""
    try:
        import streamlit as st
        return "supabase" in st.secrets
    except Exception:
        return False

# app.py の既存インポートとの後方互換性
def is_gsheet_configured() -> bool:
    return is_supabase_configured()


def _sb_get(key: str):
    """Supabase から key に対応する data を取得。なければ None。"""
    client = _supabase_client()
    if not client:
        return None
    try:
        res = client.table("app_data").select("data").eq("key", key).execute()
        if res.data:
            return res.data[0]["data"]
        return None
    except Exception:
        return None


def _sb_set(key: str, data):
    """Supabase に key と data を upsert。"""
    client = _supabase_client()
    if not client:
        return
    try:
        client.table("app_data").upsert({"key": key, "data": data}).execute()
    except Exception:
        pass


# ── DataFrame I/O ─────────────────────────────────────────────
def load_df(館: str, name: str, columns: list = None) -> pd.DataFrame:
    """
    DataFrame を読み込む。
    Supabase 設定済み → Supabase から読む（ローカルにフォールバック）
    未設定     → ローカル CSV から読む
    """
    raw = _sb_get(f"{館}_{name}")
    if raw is not None:
        try:
            df = pd.DataFrame(raw).fillna("").astype(str)
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
    DataFrame を保存する（ローカル CSV に常に保存 + Supabase が設定済みなら Supabase にも保存）。
    """
    # 常にローカルに保存
    p = DATA_DIR / f"{館}_{name}.csv"
    df.to_csv(p, index=False, encoding="utf-8-sig")

    # Supabase にも保存
    records = df.fillna("").astype(str).to_dict("records")
    _sb_set(f"{館}_{name}", records)


# ── JSON I/O ──────────────────────────────────────────────────
def load_json_data(館: str, name: str, default=None):
    """
    JSON データを読み込む（Supabase → ローカル JSON の順）。
    """
    raw = _sb_get(f"{館}_{name}_json")
    if raw is not None:
        return raw

    p = DATA_DIR / f"{館}_{name}.json"
    if p.exists():
        return json.loads(p.read_text(encoding="utf-8"))
    return default if default is not None else []


def save_json_data(館: str, name: str, data):
    """
    JSON データを保存する（ローカル + Supabase）。
    """
    # ローカルに常に保存
    p = DATA_DIR / f"{館}_{name}.json"
    p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    # Supabase にも保存
    _sb_set(f"{館}_{name}_json", data)


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
