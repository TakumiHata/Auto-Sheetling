# Auto-Sheetling

PDF を A4 方眼紙 Excel に変換するツール。[Sheetling](../Sheetling) の自動化バージョンです。

オリジナルでは手動で行っていた LLM とのやり取り（Phase 2）を **Gemini API** で全自動化しています。

---

## オリジナル (Sheetling) との違い

| | Sheetling | Auto-Sheetling |
|---|---|---|
| Phase 1（PDF解析） | 自動 | 自動（同じ） |
| Phase 2（LLM処理） | **手動**（プロンプトをコピー&ペースト） | **自動**（Gemini API） |
| Phase 3（Excel生成） | 自動 | 自動（同じ） |
| エラー時のコード修正 | 手動 | **自動**（Gemini API でリトライ） |

---

## 仕組み

Sheetling は PDF のレイアウトを完全再現した A4 方眼紙 Excel を生成します。
方眼紙とは、全セルが同じサイズ（例: 4mm×4mm）のグリッドで、PDF のレイアウトをピクセル単位で忠実に再現できる形式です。

### パイプライン

```
data/in/*.pdf
     │
     ▼
[Phase 1] PDF解析 (pdfplumber)
  - テキスト・矩形・罫線・テーブルを抽出
  - 全要素に Excel グリッド座標 (_row, _col) を付与
  - ページ画像を PNG として出力（Vision 補正用）
     │
     ▼
[Phase 2] Gemini API による自動処理
  │
  ├─ Step 1:   TABLE_ANCHOR_PROMPT
  │            PDF解析データ → Excel レイアウト仕様 JSON
  │            （テキスト・罫線要素の座標マッピング）
  │
  ├─ Step 1.5: LAYOUT_REVIEW_PROMPT
  │            レイアウト JSON の検証・補正
  │            （欠落テキスト補完・重複除去・座標クランプ）
  │
  ├─ Step 1.6: VISUAL_BORDER_REVIEW_PROMPT  ※ --vision オプション時のみ
  │            PDF ページ画像 + JSON を Vision LLM に渡して
  │            border_rect の過剰検出を視覚的に補正
  │
  └─ Step 2:   CODE_GEN_PROMPT
               補正済み JSON → Python (openpyxl) コード生成
     │
     ▼
[Phase 3] 生成コードを実行 → output.xlsx → {pdf_name}.xlsx
  - エラー時は Gemini API で自動修正して最大 N 回リトライ
     │
     ▼
data/out/{pdf_name}/{pdf_name}.xlsx
```

### グリッドサイズ

| サイズ | セル幅 | セル高 | 最大列数 | 最大行数 |
|--------|--------|--------|----------|----------|
| small  | 4.0 mm | 4.0 mm | 62       | 76       |
| medium | 6.0 mm | 6.0 mm | 36       | 50       |
| large  | 8.0 mm | 8.0 mm | 26       | 38       |

---

## セットアップ

### 1. リポジトリのクローン

```bash
git clone <repository_url>
cd Auto-Sheetling
```

### 2. 依存関係のインストール

```bash
pip install -r requirements.txt
```

### 3. Gemini API キーの設定

#### 3-1. API キーの取得

1. [Google AI Studio](https://aistudio.google.com/app/apikey) にアクセス
2. Google アカウントでサインイン
3. **「Create API key」** をクリックしてキーを生成
4. 生成されたキー（`AIza...` で始まる文字列）をコピー

#### 3-2. .env ファイルへの記載

リポジトリルートに `.env` ファイルがあります。`your_api_key_here` の部分を取得したキーに書き換えてください:

```
GEMINI_API_KEY=AIzaXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
```

> **注意:** `.env` ファイルは `.gitignore` により git 管理対象外になっています。API キーを誤ってコミットしないよう注意してください。

---

## 使い方

### 全自動実行（推奨）

`data/in/` に PDF を配置して実行するだけです。

```bash
python -m src.main auto
```

#### オプション

```bash
# グリッドサイズを指定（デフォルト: small）
python -m src.main auto --grid-size medium

# Vision LLM による罫線補正を有効化（Step 1.6）
python -m src.main auto --vision

# 特定の PDF のみ処理
python -m src.main auto --pdf data/in/invoice.pdf

# エラー時のリトライ回数を変更（デフォルト: 3）
python -m src.main auto --max-retries 5

# Gemini モデルを指定（デフォルト: gemini-2.0-flash）
python -m src.main auto --model gemini-1.5-pro
```

### 手動実行（オリジナルと同じ操作）

Phase 1 と Phase 3 を個別に実行することもできます。

```bash
# Phase 1: PDF解析 & プロンプト生成
python -m src.main extract

# Phase 3: 生成コードを実行して Excel を出力
python -m src.main generate
```

---

## ディレクトリ構成

```
Auto-Sheetling/
├── src/
│   ├── main.py                   # CLI エントリーポイント
│   ├── parser/
│   │   └── pdf_extractor.py      # pdfplumber による PDF 解析
│   ├── templates/
│   │   └── prompts.py            # Gemini に渡す LLM プロンプト定義
│   ├── core/
│   │   ├── pipeline.py           # 基本パイプライン（Phase 1 & 3）
│   │   └── auto_pipeline.py      # 自動パイプライン（Phase 2 を Gemini API で実行）
│   └── utils/
│       └── logger.py             # ロギングユーティリティ
├── data/
│   ├── in/                       # 入力 PDF を配置するディレクトリ
│   └── out/                      # 出力ディレクトリ（自動生成）
│       └── {pdf_name}/
│           ├── {pdf_name}_extracted.json     # PDF 解析結果
│           ├── {pdf_name}_gen.py             # Gemini が生成した Excel 生成コード
│           ├── {pdf_name}.xlsx               # 最終出力 Excel
│           ├── images/
│           │   └── {pdf_name}_page1.png      # ページ画像（Vision 補正用）
│           └── prompts/
│               ├── {pdf_name}_prompt_step1.txt       # Step 1 プロンプト（デバッグ用）
│               ├── {pdf_name}_step1_output.json      # Step 1 Gemini レスポンス
│               ├── {pdf_name}_prompt_step1_5.txt     # Step 1.5 プロンプト（デバッグ用）
│               ├── {pdf_name}_step1_5_output.json    # Step 1.5 Gemini レスポンス
│               ├── {pdf_name}_prompt_step2.txt       # Step 2 プロンプト（デバッグ用）
│               └── {pdf_name}_prompt_error_fix.txt   # エラー修正プロンプト（失敗時のみ）
├── tests/
├── .env                          # Gemini API キー（git 管理外）
├── .gitignore
├── Dockerfile
├── docker-compose.yml
└── requirements.txt
```

---

## Docker での実行

```bash
# イメージをビルドして起動
docker compose up -d

# コンテナ内でコマンドを実行
docker compose exec app python -m src.main auto
```

`.env` ファイルは `docker-compose.yml` の `env_file` 設定により自動的に読み込まれます。

---

## 依存ライブラリ

| ライブラリ | 用途 |
|---|---|
| `pdfplumber` | PDF からテキスト・矩形・罫線・テーブルを抽出 |
| `openpyxl` | Excel ファイルの生成・スタイル設定 |
| `google-generativeai` | Gemini API クライアント |
| `python-dotenv` | `.env` ファイルから環境変数を読み込み |
| `Pillow` | PDF ページ画像を Vision LLM に渡す際の画像処理 |

---

## 出力 Excel の仕様

- 用紙: A4（縦向き or 横向きを PDF から自動判定）
- 全セルが等サイズのグリッド（方眼紙）
- テキスト: 左揃え・垂直中央・折り返しなし
- 縦文字: 255° 回転で再現
- 罫線: PDF の罫線情報を辺単位で正確に再現
- 塗りつぶし: PDF の fill_color を保持（白・黒は除外）
- 複数ページ: 1 シートに縦並びで配置し、ページ境界に改ページを挿入
- 印刷設定: スケール 100%、余白は自動計算
