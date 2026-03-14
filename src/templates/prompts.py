"""
Sheetling 用の LLM プロンプト定義集。
3ステップ・パイプライン方式:
  Step 1:   TABLE_ANCHOR_PROMPT   — PDF解析データ → Excelレイアウト仕様JSON（座標マッピング・テキスト結合）
  Step 1.5: LAYOUT_REVIEW_PROMPT  — レイアウト仕様JSONの検証・補正（欠落補完・重複除去・整合性チェック）
  Step 2:   CODE_GEN_PROMPT       — 補正済みレイアウト仕様JSON → Python(openpyxl)コード
"""

GRID_SIZES = {
    "small": {
        "col_width_mm": "4.0",
        "row_height_mm": "4.0",
        "max_cols": 62,
        "max_rows": 76,
        "excel_col_width": 1.45,
        "excel_row_height": 11.34,
        "margin_left": 0.43,
        "margin_right": 0.43,
        "margin_top": 0.41,
        "margin_bottom": 0.41
    },
    "medium": {
        "col_width_mm": "6.0",
        "row_height_mm": "6.0",
        "max_cols": 36,
        "max_rows": 50,
        "excel_col_width": 2.53,
        "excel_row_height": 17.01,
        "margin_left": 0.47,
        "margin_right": 0.47,
        "margin_top": 0.41,
        "margin_bottom": 0.41
    },
    "large": {
        "col_width_mm": "8.0",
        "row_height_mm": "8.0",
        "max_cols": 26,
        "max_rows": 38,
        "excel_col_width": 3.61,
        "excel_row_height": 22.68,
        "margin_left": 0.51,
        "margin_right": 0.51,
        "margin_top": 0.49,
        "margin_bottom": 0.49
    }
}

TABLE_ANCHOR_PROMPT = """あなたはPDF解析データからExcelレイアウト仕様を生成する設計者です。

## 座標について（最重要）

入力データの各要素には **事前計算済みのExcel座標** が付与されています。
独自に座標を計算せず、必ず以下のフィールドをそのまま参照してください。

| フィールド | 意味 |
|---|---|
| `words[i]._row` | テキストのExcel行番号 |
| `words[i]._col` | テキストのExcel列番号（開始） |
| `words[i].is_vertical` | 縦文字かどうか（`true` の場合のみ存在） |
| `words[i]._end_row` | 縦文字の下端行番号（`is_vertical=true` の場合のみ存在） |
| `rects[i]._row` / `_end_row` | 矩形の開始・終了行番号（テーブル外） |
| `rects[i]._col` / `_end_col` | 矩形の開始・終了列番号（テーブル外） |
| `rects[i]._borders` | 矩形枠の各辺に罫線があるか `{{top, bottom, left, right}}` |
| `table_border_rects[i]._row` / `_end_row` | テーブルセルの開始・終了行番号 |
| `table_border_rects[i]._col` / `_end_col` | テーブルセルの開始・終了列番号 |
| `table_border_rects[i]._borders` | セルの各辺に罫線があるか `{{top, bottom, left, right}}` |

## 処理手順

### Step 1: border_rect 要素の生成

**テーブルセル（`table_border_rects`）：**
- `table_border_rects` の各エントリを `border_rect` 要素として出力する
- 各エントリの `_row`, `_end_row`, `_col`, `_end_col` をそのまま使用する
- 各エントリの `_borders` を `borders` フィールドとしてそのまま出力する
- **隣接するセル同士で共有辺の `_borders` が両方 `false` の場合、それらは視覚的に1つのセルである可能性が高い。確信が持てる場合は統合して1つの `border_rect` にまとめよ**（統合後の row/col は最小値、end_row/end_col は最大値、borders は外周辺のみ true）

**テーブル外の矩形（`rects`）：**
- 各 rect の `_row`, `_col`, `_end_row`, `_end_col` をそのまま使用する
- 各 rect の `_borders` を `borders` フィールドとしてそのまま出力する

### Step 2: テキスト要素の生成

`words` を **`(_row, _col)` の組み合わせ** でグループ化し、同一 `(_row, _col)` のwordを1つの `text` 要素にまとめる。
これにより、テーブル内の各列セルが別々の `text` 要素として保持される（行単位の結合は禁止）。

- `row` = グループの `_row` 値
- `col` = グループの `_col` 値
- `content` = グループ内のwordを左から順に半角スペースで結合したテキスト
- `end_col` = `col` + `content` の文字数（概算）
- `font_color` = グループ内最初のwordの `font_color`（存在し、かつ黒 `"000000"` でない場合のみ含める）
- `font_size` = グループ内最初のwordの `font_size`（存在する場合のみ含める）

**縦文字（`is_vertical=true`）の扱い：**
- `is_vertical=true` のwordは他のwordとグループ化せず、単独で `text` 要素にする
- `is_vertical: true` を要素に含める
- `_end_row` が存在する場合は `end_row` として含める（縦方向の占有範囲）
- `end_col` は `col + 1`（縦文字は1列幅）

### Step 3: border_rect の fill_color

各 rect に `fill_color` フィールドが存在し、かつ白 `"FFFFFF"` でも黒 `"000000"` でもない場合、`border_rect` 要素に `"fill_color"` を含める。

## 出力フォーマット

[
  {{
    "page_number": 1,
    "elements": [
      {{
        "type": "text",
        "content": "請求書",
        "row": 2,
        "col": 20,
        "end_col": 28,
        "font_color": "FF0000",
        "font_size": 14
      }},
      {{
        "type": "text",
        "content": "合計金額",
        "row": 5,
        "col": 3,
        "end_col": 9
      }},
      {{
        "type": "text",
        "content": "承認",
        "row": 10,
        "col": 58,
        "end_col": 59,
        "end_row": 14,
        "is_vertical": true
      }},
      {{
        "type": "border_rect",
        "row": 8,
        "end_row": 9,
        "col": 2,
        "end_col": 17,
        "borders": {{"top": true, "bottom": true, "left": true, "right": false}},
        "fill_color": "CCDDEE"
      }},
      {{
        "type": "border_rect",
        "row": 8,
        "end_row": 9,
        "col": 18,
        "end_col": 27,
        "borders": {{"top": true, "bottom": true, "left": false, "right": true}}
      }}
    ]
  }}
]

入力データ:
{input_data}

【最重要】出力は `[` で始まり `]` で終わる純粋なJSON配列文字列のみとしてください。Markdownのコードブロック(```json等)、前後の説明文、思考プロセス、検証コメントは一切含めないでください。JSON以外の文字を1文字でも出力するとSTEP 1.5で処理できなくなります。"""


LAYOUT_REVIEW_PROMPT = """あなたはExcelレイアウト設計の検証者です。
STEP 1が生成したレイアウト仕様JSONを検証・補正し、同じフォーマットの修正済みJSONを出力してください。

## 検証・修正項目

### 1. border_rect の座標整合性
- `row > end_row` または `col > end_col` となっている場合は値を入れ替えて修正する
- `row == end_row` かつ `col == end_col`（1×1セル）のborder_rectは削除する

### 2. 重複するborder_rectの除去
- `row`, `end_row`, `col`, `end_col` がすべて同一のborder_rectが複数ある場合、1つだけ残す

### 2.5. border_rectのbordersフィールド
- `borders` フィールドが存在する場合、値を変更せずそのまま保持すること
- STEP 1 で隣接セルが統合された場合、統合後の `borders` は外周辺（視覚的に線がある辺）のみ true にすること

### 3. テキスト要素の欠落補完
以下の元PDF解析データの `words` と照合し、レイアウトJSONに欠落しているテキストを補完する：
- 元データの各ページの `words` を確認し、対応するページの `text` 要素に含まれていない `text` フィールドの内容を探す
- 欠落が確認された場合、該当wordの `_row` を `row`、`_col` を `col` として `text` 要素を追加する
- `end_col` = `col` + `text` の文字数（概算）
- 空白・記号のみのword（`text` が空文字、スペースのみ、または1文字以下の記号）は無視してよい

### 4. text 要素の重複除去
- 同一ページ内で `row` と `col` が同じ `text` 要素が複数ある場合、`content` を半角スペースで結合して1つにまとめる

### 5. 座標のクランプ
- `row`, `end_row` が {max_rows} を超えている場合は {max_rows} に切り詰める
- `col`, `end_col` が {max_cols} を超えている場合は {max_cols} に切り詰める

## 元のPDF解析データ（参照用）

{input_data}

## STEP 1の出力（検証・修正対象）

{step1_output}

【最重要】出力は `[` で始まり `]` で終わる純粋なJSON配列文字列のみとしてください。Markdownのコードブロック(```json等)、前後の説明文、思考プロセス、検証コメントは一切含めないでください。JSON以外の文字を1文字でも出力するとSTEP 2で処理できなくなります。"""


VISUAL_BORDER_REVIEW_PROMPT = """あなたはPDF帳票のレイアウト検証者です。
添付した **PDFページ画像** と、以下の **Excelレイアウト仕様JSON** を見比べて、`border_rect` 要素を視覚的に正確な状態に修正してください。

## 問題の背景

PDFの自動解析ツール（pdfplumber）は、テーブルのセル境界を過剰に検出することがあります。
たとえば、視覚的には「3行×5列=15セル」のテーブルが、内部の交点から「10行×12列=120セル」として誤認識される場合があります。
この修正ステップでは、**PDFページ画像を目視で確認しながら、実際に罫線が引かれているセルだけが残るよう** `border_rect` を修正します。

## 修正ルール

1. **画像を見て、視覚的に区別できるセルの数・形を確認する**
2. 隣接する複数の `border_rect` が視覚的に同じ1つのセルを表しているなら、それらを1つの大きな `border_rect` に統合する
   - 統合後の `row` = 統合前の最小 `row`
   - 統合後の `end_row` = 統合前の最大 `end_row`
   - 統合後の `col` = 統合前の最小 `col`
   - 統合後の `end_col` = 統合前の最大 `end_col`
3. 画像上に罫線が存在しない `border_rect` は削除する
4. `text` 要素は変更しない（座標・内容ともにそのまま保持）
5. 座標値（row/col）はすべて整数で出力する

## 修正対象のレイアウト仕様JSON（STEP 1.5の出力）

{step1_5_output}

【最重要】出力は `[` で始まり `]` で終わる純粋なJSON配列文字列のみとしてください。Markdownのコードブロック(```json等)、前後の説明文、思考プロセス、検証コメントは一切含めないでください。"""


CODE_GEN_PROMPT = """あなたはPythonプログラミングのエキスパートです。
以下のExcelレイアウト仕様（JSON）を元に、openpyxlを使ったExcel生成スクリプトを作成してください。

## グリッド設定（必須）
全列・全行に以下の値を**ハードコード**で適用してください（mmからの再計算・変数化は禁止）：
- 列幅: `ws.column_dimensions[get_column_letter(c)].width = {excel_col_width}` （これはExcel固有の単位値。{col_width_mm}mmに相当。変換禁止）
- 行高: `ws.row_dimensions[r].height = {excel_row_height}` （これはExcel固有の単位値。{row_height_mm}mmに相当。変換禁止）
- 適用範囲: 列1〜{max_cols}、行1〜(総ページ数 × {max_rows})

## 必須インポート

スクリプト冒頭に以下を含めること：
```python
from openpyxl.styles import Border, Side, Alignment, PatternFill, Font
```

## 要素の種類

入力JSONに含まれる要素は `text` と `border_rect` の2種類のみです。

## text 処理パターン（厳守・改変禁止）

以下のコードを**一字一句そのまま**使用してください。`ws.max_row` / `ws.max_column` による境界チェック、セルの存在確認、独自のガード処理は**絶対に追加しないこと**。

```python
if item["type"] == "text":
    r = item["row"] + row_offset
    try:
        cell = ws.cell(row=r, column=item["col"] + col_offset)
        cell.value = item["content"]
        if item.get("is_vertical"):
            cell.alignment = Alignment(text_rotation=255, vertical='top', wrap_text=False)
        else:
            cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)
        font_kwargs = {{}}
        if item.get("font_color"):
            font_kwargs["color"] = item["font_color"]
        if item.get("font_size"):
            font_kwargs["size"] = item["font_size"]
        if font_kwargs:
            cell.font = Font(**font_kwargs)
    except AttributeError:
        pass
```

## border_rect 処理パターン（必須）

`borders` フィールドの各辺（top/bottom/left/right）が `true` の辺のみ罫線を描画します。
隣接する `border_rect` の外枠が接することで内部の区切り線が自然に形成されます。

```python
from openpyxl.styles import Border, Side, Alignment, PatternFill, Font
thin = Side(style='thin')

def apply_border(ws, s_row, e_row, s_col, e_col, borders, fill_color=None):
    has_top    = borders.get("top",    True)
    has_bottom = borders.get("bottom", True)
    has_left   = borders.get("left",   True)
    has_right  = borders.get("right",  True)
    fill = PatternFill(patternType='solid', fgColor=fill_color) if fill_color else None
    for r in range(s_row, e_row + 1):
        for c in range(s_col, e_col + 1):
            target = ws.cell(row=r, column=c)
            try:
                target.alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)
            except AttributeError:
                pass
            top    = thin if (r == s_row and has_top)    else None
            bottom = thin if (r == e_row and has_bottom) else None
            left   = thin if (c == s_col and has_left)   else None
            right  = thin if (c == e_col and has_right)  else None
            try:
                target.border = Border(top=top, bottom=bottom, left=left, right=right)
            except AttributeError:
                pass
            if fill:
                try:
                    target.fill = fill
                except AttributeError:
                    pass

elif item["type"] == "border_rect":
    apply_border(
        ws,
        item["row"] + row_offset, item["end_row"] + row_offset,
        item["col"] + col_offset, item["end_col"] + col_offset,
        item.get("borders", {{"top": True, "bottom": True, "left": True, "right": True}}),
        fill_color=item.get("fill_color")
    )
```

## セルスタイル（必須）
- すべてのセル: `Alignment(horizontal='left', vertical='center', wrap_text=False)`
- セル結合（ws.merge_cells）は使用禁止

## 印刷設定（必須）
- 用紙サイズ: `ws.page_setup.paperSize = 9`（A4、定数は使わず直接代入）
- 向き: `ws.page_setup.orientation = '{orientation}'`（定数は使わず直接代入）
- スケーリング: fitToWidth / fitToPage は設定しない（等倍100%が大前提）
- 余白（数値は変更禁止）:
  ```python
  ws.page_margins.left = {margin_left}
  ws.page_margins.right = {margin_right}
  ws.page_margins.top = {margin_top}
  ws.page_margins.bottom = {margin_bottom}
  ```
- 印刷範囲: 入力JSONに `print_range` があればそれを使用。なければ、入力データの各要素から実際に使用されている最大行・最大列を求め、`A1:{{最大列文字}}{{最大行番号}}` を設定すること。`max_rows` や `total_pages × max_rows` を印刷範囲の計算に**使用しないこと**（コンテンツが収まらず2ページ目に溢れる原因になる）。

## 複数ページ対応
- 全ページを1シート（`wb.active`）に縦に並べる
- オフセット定数（コード先頭で定義）: `col_offset = 1`（左右1マス余白）、`row_padding = 2`（上下2マス余白）
- 2ページ目以降のrow_offset = `(page_number - 1) * {max_rows} + row_padding`
- ページ境界に改ページ設定:
  ```python
  from openpyxl.worksheet.pagebreak import Break
  ws.row_breaks.append(Break(id=page_number * ({max_rows} + row_padding)))
  ```
- 行の高さ設定範囲（`row_dimensions`）は `総ページ数 × {max_rows}` まで適用すること。ただし**印刷範囲はこの値を使わず**、上記の「実際に使用されている最大行」に基づくこと。データの最終使用行に定数を加算するバッファ処理は**禁止**。

## 技術的制約
- `True`/`False`（Python形式）を使用。JSONの`true`/`false`は禁止
- `[cite: ...]` のようなアノテーションタグは絶対に含めない（SyntaxErrorの原因）
- 出力ファイル名: `output.xlsx`
- `ws.page_margins` への dict 代入は**絶対に禁止**（`AttributeError: 'dict' object has no attribute 'to_tree'` の原因）。必ず属性代入形式を使用すること:
  ```python
  # ✗ 禁止
  ws.page_margins = {{'left': 0.47, ...}}
  # ✓ 正しい形式（上記「印刷設定」の通り）
  ws.page_margins.left = 0.47
  ```

入力データ（Step 1のレイアウト仕様）:
【重要】以下の入力データに `[` で始まる JSON 配列が含まれています。JSON配列の前後に説明文や検証コメントが含まれている場合は完全に無視し、`[` から `]` までのJSON部分のみを使用してください。
{input_data}

出力はPythonコードのみをマークダウンのコードブロック（```python ... ```）で出力してください。前後の挨拶・説明文・思考プロセスは一切不要です。"""


CODE_ERROR_FIXING_PROMPT = """あなたはPythonプログラミングのエキスパートです。
以下のExcel生成用Pythonスクリプトを実行したところ、エラーが発生しました（または期待する結果が得られませんでした）。
エラーメッセージと現在のコードを分析し、問題を修正した新しいPythonスクリプトを生成してください。

【最重要：Python文法の遵守】
修正後のコード内では、必ず Python の予約語である `True`/`False` （先頭大文字）を使用してください。小文字の `true`/`false` を含めないでください。

【エラーメッセージ・実行結果】
{error_msg}

【現在のコード】
```python
{code}
```
【既知のエラーパターンと修正方法】
- `AttributeError: 'dict' object has no attribute 'to_tree'`
  → `ws.page_margins = {{...}}` のように dict を代入している。以下の属性代入形式に修正すること:
  ```python
  ws.page_margins.left = 0.47
  ws.page_margins.right = 0.47
  ws.page_margins.top = 0.41
  ws.page_margins.bottom = 0.41
  ws.page_margins.header = 0.0
  ws.page_margins.footer = 0.0
  ```
- `AttributeError: 'MergedCell' object attribute 'value' is read-only`
  → セルへの値代入や `merge_cells` 等の処理を `try...except AttributeError: pass` で囲むこと。

【最重要】出力は、修正済みの実行可能な Python コードのみをマークダウンのコードブロック (```python ... ```) で出力してください。前後の挨拶、解説、謝罪、思考プロセスなどは一切不要です。
【最重要】出力するコード内には、`[cite: ...]` のような参照タグやアノテーション（例: ` [cite: 271]`）を絶対に含めないでください。SyntaxErrorの原因となります。純粋で実行可能なPythonコードのみを出力してください。また、コードブロックの外にいかなるテキストも記述しないでください。"""
