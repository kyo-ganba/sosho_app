# 送迎表自動生成ツール — セットアップ手順

## 構成ファイル

```
sosho_app/
├── app.py            # メインアプリ（画面定義）
├── master.py         # 利用者マスタ管理
├── routing.py        # ルーティングロジック
├── excel_export.py   # Excel出力
├── requirements.txt  # 必要ライブラリ
└── data/             # マスタデータ保存先（自動生成）
    ├── Ⅴ番館_master.csv
    ├── Ⅴ番館_staff.json
    └── Ⅴ番館_vehicles.json
```

---

## クラウド無料公開（Streamlit Community Cloud）

### 手順

1. **GitHubリポジトリを作成**
   - GitHub.com で新しいリポジトリを作る（例: `sosho-app`）
   - `sosho_app/` 内のファイルをすべてアップロード

2. **Streamlit Community Cloudにデプロイ**
   - https://share.streamlit.io/ にアクセス
   - GitHubアカウントでサインイン
   - 「New app」→ リポジトリ・ブランチ・`app.py` を選択
   - 「Deploy」ボタンを押すだけ

3. **URLが発行される**
   - 例: `https://sosho-app-xxxxx.streamlit.app`
   - このURLをスタッフに共有するだけで全員が使えるようになります

### 注意点
- 無料プランでは **マスタデータはサーバー再起動のたびにリセット**されます
- 本格運用する場合は `data/` フォルダをGitリポジトリに含めるか、
  Google Sheets や Supabase（無料DB）をストレージに使う改修が必要です

---

## ローカルで試す場合

```bash
# Pythonのインストール確認
python --version   # 3.10以上推奨

# ライブラリのインストール
pip install -r requirements.txt

# アプリ起動
streamlit run app.py
```

ブラウザが自動で開き `http://localhost:8501` でアクセスできます。

---

## 初期設定の流れ

1. サイドバーで事業所（Ⅰ〜Ⅴ番館）を選択
2. **「利用者マスタ管理」** → リタリコCSVを取込 → 列マッピングを設定して保存
3. **「車両・スタッフ設定」** → 車両名・定員、スタッフ名・運転可否を登録
4. **「当日入力・送迎表生成」** → 参加者チェック → 「送迎ルートを自動生成」
5. 画面でルートを確認・修正 → 「Excelを出力」でダウンロード

---

## カスタマイズポイント

| ファイル | 変更箇所 |
|---|---|
| `routing.py` | `TIME_GAP_MINUTES` でグループ化の時間幅を調整（デフォルト20分）|
| `routing.py` | `HOME_TIME_DEFAULT` で送りのデフォルト出発時刻を変更 |
| `excel_export.py` | 列レイアウト・色・フォントを現行フォーマットに合わせる |
| `master.py` | `import_from_ritalico()` のマッピングロジックをリタリコの実際の列名に調整 |
