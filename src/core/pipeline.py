"""
Sheetling パイプライン。
3ステップ・パイプライン方式:
1. 解析 (pdfplumber) → extracted_data.json + prompt_step1.txt + prompt_step1_5.txt + prompt_step2.txt
2. 描画 (openpyxl) — LLMが生成した _gen.py を実行
"""

import json
import re
from pathlib import Path

from src.parser.pdf_extractor import extract_pdf_data
from src.templates.prompts import TABLE_ANCHOR_PROMPT, LAYOUT_REVIEW_PROMPT, VISUAL_BORDER_REVIEW_PROMPT, CODE_GEN_PROMPT, CODE_ERROR_FIXING_PROMPT, GRID_SIZES
from src.utils.logger import get_logger


def _sanitize_generated_code(code: str) -> tuple[str, list[str]]:
    """生成コードの既知の問題パターンを検出・自動修正する。"""
    fixes = []

    # 修正: ws.page_margins = {...} → 属性代入形式に変換
    margins_dict_pattern = re.compile(r"ws\.page_margins\s*=\s*\{([^}]*)\}", re.DOTALL)
    match = margins_dict_pattern.search(code)
    if match:
        kv_pattern = re.compile(r"['\"](\w+)['\"]\s*:\s*([\d.]+)")
        pairs = kv_pattern.findall(match.group(1))
        if pairs:
            replacement = "\n".join(f"ws.page_margins.{k} = {v}" for k, v in pairs)
            code = margins_dict_pattern.sub(replacement, code)
            fixes.append("ws.page_margins への dict 代入を属性代入形式に自動修正しました")

    return code, fixes


def _compute_grid_coords(page: dict, max_rows: int, max_cols: int) -> None:
    """
    PDF座標をExcel行・列番号に変換し、各要素にインプレースで付与する。
    Y・X座標ともにクラスタリングを行い、近接する座標を同一行・列に統一する。
    """
    page_height = page['height']
    page_width = page['width']
    grid_h = page_height / max_rows
    grid_w = page_width / max_cols

    def snap(v: float) -> float:
        return round(float(v), 2)

    def build_cluster_map(raw_vals: set, grid_size: float, max_idx: int, anchor_vals: set = None) -> dict:
        """
        近接する座標値をクラスタリングしてグリッドインデックスに変換する。
        anchor_vals に含まれる値は直前のクラスタと近接していても必ず独立したクラスタを開始する。
        これによりテーブル列境界が隣接列と合算されるのを防ぐ。
        """
        anchor_vals = anchor_vals or set()
        sorted_vals = sorted(raw_vals)
        clusters: list = []
        for v in sorted_vals:
            if not clusters or v - clusters[-1][0] > grid_size * 0.5 or v in anchor_vals:
                clusters.append([v])
            else:
                clusters[-1].append(v)
        val_map = {}
        for cluster in clusters:
            centroid = sum(cluster) / len(cluster)
            idx = max(1, min(max_idx, 1 + int(centroid / grid_size)))
            for v in cluster:
                val_map[v] = idx
        return val_map

    # 全Y・X座標を収集
    y_vals: set = set()
    x_vals: set = set()
    # テーブル列境界X座標（クラスタリング時に独立扱いにするため別途保持）
    table_col_x_anchors: set = set()
    # テーブル行境界Y座標（クラスタリング時に独立扱いにするため別途保持）
    table_row_y_anchors: set = set()

    for w in page['words']:
        y_vals.add(snap(w['top']))
        x_vals.add(snap(w['x0']))
        x_vals.add(snap(w['x1']))
        if w.get('is_vertical') and 'bottom' in w:
            y_vals.add(snap(w['bottom']))  # 縦文字の下端もグリッドに含める
    for r in page['rects']:
        y_vals.add(snap(r['top']))
        y_vals.add(snap(r['bottom']))
        x_vals.add(snap(r['x0']))
        x_vals.add(snap(r['x1']))
    for bbox in page['table_bboxes']:
        y_vals.add(snap(bbox[1]))  # top
        y_vals.add(snap(bbox[3]))  # bottom
    for col_xs in page['table_col_x_positions']:
        for x in col_xs:
            sx = snap(x)
            x_vals.add(sx)
            table_col_x_anchors.add(sx)
    for row_ys in page.get('table_row_y_positions', []):
        for y in row_ys:
            sy = snap(y)
            y_vals.add(sy)
            table_row_y_anchors.add(sy)
    for cells in page.get('table_cells', []):
        for c in cells:
            y_vals.add(snap(c['top']))
            y_vals.add(snap(c['bottom']))
            x_vals.add(snap(c['x0']))
            x_vals.add(snap(c['x1']))
    # エッジ座標もクラスタリングに含める（罫線位置をグリッドに正確に反映）
    for edge in page.get('h_edges', []):
        y_vals.add(snap(edge['y']))
        x_vals.add(snap(edge['x0']))
        x_vals.add(snap(edge['x1']))
    for edge in page.get('v_edges', []):
        x_vals.add(snap(edge['x']))
        y_vals.add(snap(edge['y0']))
        y_vals.add(snap(edge['y1']))

    y_map = build_cluster_map(y_vals, grid_h, max_rows, anchor_vals=table_row_y_anchors)
    x_map = build_cluster_map(x_vals, grid_w, max_cols, anchor_vals=table_col_x_anchors)

    # テーブル列境界が同一グリッド列に潰れた場合の後処理:
    # 各テーブルの列X座標を左から順に走査し、前の列と同じグリッド列になっていたら +1 する。
    for col_xs in page['table_col_x_positions']:
        snapped_xs = sorted(set(snap(x) for x in col_xs))
        prev_idx = 0
        for x in snapped_xs:
            idx = x_map[x]
            if idx <= prev_idx:
                idx = prev_idx + 1
            idx = min(idx, max_cols)
            x_map[x] = idx
            prev_idx = idx

    # テーブル行境界が同一グリッド行に潰れた場合の後処理:
    # 各テーブルの行Y座標を上から順に走査し、前の行と同じグリッド行になっていたら +1 する。
    for row_ys in page.get('table_row_y_positions', []):
        snapped_ys = sorted(set(snap(y) for y in row_ys))
        prev_idx = 0
        for y in snapped_ys:
            idx = y_map[y]
            if idx <= prev_idx:
                idx = prev_idx + 1
            idx = min(idx, max_rows)
            y_map[y] = idx
            prev_idx = idx

    # テーブル底辺直下の注釈要素がテーブル底辺と同一グリッド行に落ちる場合の補正。
    # グリッド解像度が低い（行高＞底辺〜注釈上端の間隔）ときに発生する位置重複を解消する。
    # テーブル境界 y 値（table_row_y_anchors）は除外し、注釈等のコンテンツ y 値のみ対象とする。
    for row_ys_list in page.get('table_row_y_positions', []):
        if not row_ys_list:
            continue
        table_bottom_y = snap(max(row_ys_list))
        if table_bottom_y not in y_map:
            continue
        table_end_row = y_map[table_bottom_y]
        # テーブル底辺直下（1グリッド行以内）に非テーブル境界の y 値が同一行に落ちていれば衝突
        has_collision = any(
            table_bottom_y < yv <= table_bottom_y + grid_h
            and y_map[yv] == table_end_row
            and yv not in table_row_y_anchors
            for yv in y_map
        )
        if not has_collision:
            continue
        # 衝突あり: テーブル底辺より下の全 y 値を 1 行ずらして空行を挿入
        for yv in list(y_map.keys()):
            if yv > table_bottom_y:
                y_map[yv] = min(y_map[yv] + 1, max_rows)

    # words に付与
    for w in page['words']:
        w['_row'] = y_map[snap(w['top'])]
        w['_col'] = x_map[snap(w['x0'])]
        if w.get('is_vertical') and 'bottom' in w:
            sv = snap(w['bottom'])
            w['_end_row'] = y_map.get(sv, w['_row'])

    # rects に付与
    for r in page['rects']:
        r['_row'] = y_map[snap(r['top'])]
        r['_end_row'] = y_map[snap(r['bottom'])]
        r['_col'] = x_map[snap(r['x0'])]
        r['_end_col'] = x_map[snap(r['x1'])]

    # テーブル内に含まれる rects を除外（table_border_rects で代替するため）
    tol = 1.0
    table_bboxes = page.get('table_bboxes', [])

    def is_inside_table(r: dict) -> bool:
        for bbox in table_bboxes:
            if (r['x0'] >= bbox[0] - tol and r['x1'] <= bbox[2] + tol and
                    r['top'] >= bbox[1] - tol and r['bottom'] <= bbox[3] + tol):
                return True
        return False

    page['rects'] = [r for r in page['rects'] if not is_inside_table(r)]

    # テーブルの列・行グリッドから border_rect を生成（pdfplumber が検出した列数×行数）
    table_border_rects = []
    for col_xs, row_ys in zip(page.get('table_col_x_positions', []),
                               page.get('table_row_y_positions', [])):
        col_xs_s = sorted(set(snap(x) for x in col_xs))
        row_ys_s = sorted(set(snap(y) for y in row_ys))
        for ri in range(len(row_ys_s) - 1):
            for ci in range(len(col_xs_s) - 1):
                table_border_rects.append({
                    '_row':        y_map.get(row_ys_s[ri], 1),
                    '_end_row':    y_map.get(row_ys_s[ri + 1], 1),
                    '_col':        x_map.get(col_xs_s[ci], 1),
                    '_end_col':    x_map.get(col_xs_s[ci + 1], 1),
                    '_pdf_x0':     col_xs_s[ci],
                    '_pdf_top':    row_ys_s[ri],
                    '_pdf_x1':     col_xs_s[ci + 1],
                    '_pdf_bottom': row_ys_s[ri + 1],
                })
    page['table_border_rects'] = table_border_rects

    # ---- エッジから辺ごとの罫線有無を判定 ----------------------------------------

    def _nearest_idx(val: float, coord_map: dict) -> int:
        """valに最も近いcoord_mapのキーに対応するグリッドインデックスを返す。"""
        if not coord_map:
            return 1
        sv = snap(val)
        if sv in coord_map:
            return coord_map[sv]
        return coord_map[min(coord_map.keys(), key=lambda k: abs(k - sv))]

    # エッジをグリッド座標に変換してマップ化
    # h_edge_map: row_idx -> [(col_start, col_end), ...]
    # v_edge_map: col_idx -> [(row_start, row_end), ...]
    h_edge_map: dict = {}
    for edge in page.get('h_edges', []):
        ri = _nearest_idx(edge['y'], y_map)
        cs = _nearest_idx(edge['x0'], x_map)
        ce = _nearest_idx(edge['x1'], x_map)
        h_edge_map.setdefault(ri, []).append((min(cs, ce), max(cs, ce)))

    v_edge_map: dict = {}
    for edge in page.get('v_edges', []):
        ci = _nearest_idx(edge['x'], x_map)
        rs = _nearest_idx(edge['y0'], y_map)
        re = _nearest_idx(edge['y1'], y_map)
        v_edge_map.setdefault(ci, []).append((min(rs, re), max(rs, re)))

    def _overlaps_h(edges: list, col_s: int, col_e: int) -> bool:
        """エッジリストの中に col_s〜col_e スパンと 50% 以上重複するものがあるか。"""
        span = col_e - col_s
        if span <= 0:
            return any(cs <= col_s <= ce for cs, ce in edges)
        for cs, ce in edges:
            overlap = min(ce, col_e) - max(cs, col_s)
            if overlap >= span * 0.5:
                return True
        return False

    def _overlaps_v(edges: list, row_s: int, row_e: int) -> bool:
        """エッジリストの中に row_s〜row_e スパンと 50% 以上重複するものがあるか。"""
        span = row_e - row_s
        if span <= 0:
            return any(rs <= row_s <= re for rs, re in edges)
        for rs, re in edges:
            overlap = min(re, row_e) - max(rs, row_s)
            if overlap >= span * 0.5:
                return True
        return False

    def _has_h(row: int, col_s: int, col_e: int) -> bool:
        """指定行に col_s〜col_e をカバーする水平エッジがあるか。"""
        return _overlaps_h(h_edge_map.get(row, []), col_s, col_e)

    def _has_v(col: int, row_s: int, row_e: int) -> bool:
        """指定列に row_s〜row_e をカバーする垂直エッジがあるか。"""
        return _overlaps_v(v_edge_map.get(col, []), row_s, row_e)

    # table_border_rects に _borders を付与
    for tbr in page['table_border_rects']:
        r, er, c, ec = tbr['_row'], tbr['_end_row'], tbr['_col'], tbr['_end_col']
        tbr['_borders'] = {
            'top':    _has_h(r,  c, ec),
            'bottom': _has_h(er, c, ec),
            'left':   _has_v(c,  r, er),
            'right':  _has_v(ec, r, er),
        }

    # rects にも _borders を付与（矩形枠の各辺）
    for rect in page['rects']:
        r, er = rect['_row'], rect['_end_row']
        c, ec = rect['_col'], rect['_end_col']
        rect['_borders'] = {
            'top':    _has_h(r,  c, ec),
            'bottom': _has_h(er, c, ec),
            'left':   _has_v(c,  r, er),
            'right':  _has_v(ec, r, er),
        }

    # ---- 隣接セル間の _borders 整合性パス ---------------------------------------
    # 量子化誤差で A.right と B.left（共有辺）が不一致になる場合を OR で統一する。
    tbrs = page['table_border_rects']
    # 水平方向: (row, end_row) -> {col -> tbr}
    h_band: dict = {}
    # 垂直方向: (col, end_col) -> {row -> tbr}
    v_band: dict = {}
    for tbr in tbrs:
        h_band.setdefault((tbr['_row'], tbr['_end_row']), {})[tbr['_col']] = tbr
        v_band.setdefault((tbr['_col'], tbr['_end_col']), {})[tbr['_row']] = tbr

    for tbr in tbrs:
        # 右隣: 同じ row/end_row 帯で _col == self._end_col
        right_neighbor = h_band.get((tbr['_row'], tbr['_end_row']), {}).get(tbr['_end_col'])
        if right_neighbor:
            merged = tbr['_borders']['right'] or right_neighbor['_borders']['left']
            tbr['_borders']['right'] = merged
            right_neighbor['_borders']['left'] = merged
        # 下隣: 同じ col/end_col 帯で _row == self._end_row
        bottom_neighbor = v_band.get((tbr['_col'], tbr['_end_col']), {}).get(tbr['_end_row'])
        if bottom_neighbor:
            merged = tbr['_borders']['bottom'] or bottom_neighbor['_borders']['top']
            tbr['_borders']['bottom'] = merged
            bottom_neighbor['_borders']['top'] = merged

    # ---------------------------------------------------------------------------------

    # ---- PDF 余白分の空き行を除去するため _row を正規化 --------------------------------
    # PDF の上部余白（margin_top）がグリッド行として現れ、コンテンツ前に数行の空白が生じる。
    # 全要素の最小 _row を求め、1 になるようにシフトして空き行を除去する。
    all_rows = (
        [w['_row'] for w in page['words'] if '_row' in w]
        + [r['_row'] for r in page['rects'] if '_row' in r]
        + [tbr['_row'] for tbr in page['table_border_rects']]
    )
    if all_rows:
        row_shift = min(all_rows) - 1
        if row_shift > 0:
            for w in page['words']:
                if '_row' in w:
                    w['_row'] -= row_shift
                    if '_end_row' in w:
                        w['_end_row'] -= row_shift
            for r in page['rects']:
                if '_row' in r:
                    r['_row'] -= row_shift
                    r['_end_row'] -= row_shift
            for tbr in page['table_border_rects']:
                tbr['_row'] -= row_shift
                tbr['_end_row'] -= row_shift

    # ---- 以下はグリッド座標計算には使用したが LLM には渡さない -------------------------
    page.pop('table_cells', None)
    page.pop('table_data', None)
    page.pop('table_row_y_positions', None)
    page.pop('h_edges', None)
    page.pop('v_edges', None)

logger = get_logger(__name__)


def _apply_borders_to_xlsx(xlsx_path: str, extracted_data: dict, max_rows: int) -> None:
    """
    extracted_data の table_border_rects / rects の _borders を
    openpyxl で直接 XLSX ファイルに適用する（LLM 非依存の確定的罫線描画）。

    CODE_GEN_PROMPT と同じオフセット定数を使用:
      COL_OFFSET  = 1  （左1マス余白）
      ROW_PADDING = 2  （ページ上部2マス余白）
      row_offset for page N = (N-1) * max_rows + ROW_PADDING
    """
    from openpyxl import load_workbook
    from openpyxl.styles import Border, Side

    COL_OFFSET = 1
    ROW_PADDING = 2

    wb = load_workbook(xlsx_path)
    ws = wb.active
    thin = Side(style='thin')

    def _draw(s_row: int, e_row: int, s_col: int, e_col: int, borders: dict) -> None:
        has_top    = borders.get('top',    True)
        has_bottom = borders.get('bottom', True)
        has_left   = borders.get('left',   True)
        has_right  = borders.get('right',  True)
        for r in range(s_row, e_row + 1):
            for c in range(s_col, e_col + 1):
                try:
                    cell = ws.cell(row=r, column=c)
                    # 既存の Border オブジェクトを読み取り、各辺を個別にマージ書き込み
                    # （完全上書きすると他のセルが書いた辺を消してしまうため）
                    existing = cell.border
                    new_top    = thin if (r == s_row and has_top)    else existing.top
                    new_bottom = thin if (r == e_row and has_bottom) else existing.bottom
                    new_left   = thin if (c == s_col and has_left)   else existing.left
                    new_right  = thin if (c == e_col and has_right)  else existing.right
                    cell.border = Border(top=new_top, bottom=new_bottom,
                                         left=new_left, right=new_right)
                except AttributeError:
                    pass

    total = 0
    for page in extracted_data.get('pages', []):
        page_number = page.get('page_number', 1)
        row_offset = (page_number - 1) * max_rows + ROW_PADDING

        for tbr in page.get('table_border_rects', []):
            _draw(
                tbr['_row'] + row_offset, tbr['_end_row'] + row_offset,
                tbr['_col'] + COL_OFFSET, tbr['_end_col'] + COL_OFFSET,
                tbr.get('_borders', {'top': True, 'bottom': True, 'left': True, 'right': True}),
            )
            total += 1

        for rect in page.get('rects', []):
            if '_row' not in rect:
                continue
            _draw(
                rect['_row'] + row_offset, rect['_end_row'] + row_offset,
                rect['_col'] + COL_OFFSET, rect['_end_col'] + COL_OFFSET,
                rect.get('_borders', {'top': True, 'bottom': True, 'left': True, 'right': True}),
            )
            total += 1

    wb.save(xlsx_path)
    logger.info(f"[apply_borders] {total} 個の罫線要素を適用しました")


def _has_japanese(text: str) -> bool:
    """文字列に日本語文字（漢字・ひらがな・カタカナ・全角記号）が含まれるか判定する。"""
    return any(
        '\u3040' <= c <= '\u30ff'  # ひらがな・カタカナ
        or '\u4e00' <= c <= '\u9fff'  # CJK 統合漢字
        or '\uff00' <= c <= '\uffef'  # 全角英数・記号
        for c in text
    )


def _join_word_texts(texts: list) -> str:
    """
    word テキストのリストを結合する。
    TABLE_ANCHOR_PROMPT と同じルール:
      - 日本語文字を含む場合はスペースなし
      - 英数字のみの場合は半角スペースで結合
    """
    combined = ''.join(texts)
    if _has_japanese(combined):
        return combined
    return ' '.join(t for t in texts if t.strip())


def _fill_missing_text(layout_json_str: str, extracted_data: dict) -> str:
    """
    LLMが生成したレイアウトJSONに対し、extracted_dataのwordsと照合して
    欠落しているテキスト要素をプログラム的に補完する。

    Step 1 / Step 1.5 の LLM が見落とした word を確実に補う。
    既に text 要素が存在する (row, col) には追加しない（上書き禁止）。
    """
    try:
        layout = json.loads(layout_json_str)
    except (json.JSONDecodeError, ValueError):
        return layout_json_str  # パース失敗時はそのまま返す

    total_added = 0
    for page_layout in layout:
        page_num = page_layout.get('page_number', 1)
        page_data = next(
            (p for p in extracted_data['pages'] if p['page_number'] == page_num),
            None,
        )
        if not page_data:
            continue

        # 既存 text 要素の (row, col) を収集
        existing: set = set()
        for elem in page_layout.get('elements', []):
            if elem.get('type') == 'text' and 'row' in elem and 'col' in elem:
                existing.add((elem['row'], elem['col']))

        # words を (_row, _col) でグループ化
        groups: dict = {}
        for w in page_data.get('words', []):
            if '_row' not in w or '_col' not in w:
                continue
            key = (w['_row'], w['_col'])
            groups.setdefault(key, []).append(w)

        added = []
        for (row, col), words in sorted(groups.items()):
            if (row, col) in existing:
                continue
            content = _join_word_texts([w.get('text', '') for w in words])
            stripped = content.strip()
            # 空白・純粋な区切り記号（ASCII句読点の1文字）はスキップ
            # ただし △▼○● 等の図形記号・日本語1文字は意味があるため残す
            if not stripped or (len(stripped) == 1 and stripped in '!"#$%&\'()*+,-./:;<=>?@[\\]^_`{|}~'):
                continue
            elem: dict = {
                'type': 'text',
                'content': content,
                'row': row,
                'col': col,
                'end_col': col + len(content),
            }
            first = words[0]
            if first.get('font_color') and first['font_color'] != '000000':
                elem['font_color'] = first['font_color']
            if first.get('font_size'):
                elem['font_size'] = first['font_size']
            added.append(elem)

        if added:
            page_layout['elements'].extend(added)
            total_added += len(added)

    if total_added:
        logger.info(f"[fill_missing_text] {total_added} 個の欠落テキスト要素を補完しました")

    return json.dumps(layout, ensure_ascii=False)


# A4 縦の基準サイズ (pt) — GRID_SIZES のセル密度はこのサイズを基準に調整されている
_A4_W_PT: float = 595.28
_A4_H_PT: float = 841.89


def _setup_grid_params(first_page: dict, grid_size: str) -> dict:
    """
    ページ寸法に基づいてグリッドパラメータを設定する。

    GRID_SIZES の A4 基準値から 1グリッドセルあたりのポイント数を算出し、
    実際のページ寸法に比例して max_cols / max_rows を動的計算する。
    これにより A4 以外の用紙サイズ（A3 など）にも正しく対応できる。
    """
    ref = GRID_SIZES.get(grid_size, GRID_SIZES["small"])
    grid_params = dict(ref)

    # 実ページ寸法から max_cols / max_rows を動的計算
    pt_per_col = _A4_W_PT / ref['max_cols']
    pt_per_row = _A4_H_PT / ref['max_rows']
    grid_params['max_cols'] = max(1, round(first_page['width'] / pt_per_col))
    grid_params['max_rows'] = max(1, round(first_page['height'] / pt_per_row))

    # 用紙サイズ検出（long side > 1000pt → A3）
    max_dim_pt = max(first_page['width'], first_page['height'])
    grid_params['paper_size'] = 8 if max_dim_pt > 1000 else 9  # 8=A3, 9=A4

    # 向き
    is_landscape = first_page['width'] > first_page['height']
    grid_params['orientation'] = 'landscape' if is_landscape else 'portrait'

    return grid_params


class SheetlingPipeline:
    """
    1. PDF を解析してプロンプトを出力する (Phase 1)。
    2. ユーザーがLLMから得たコードを実行し、Excel方眼紙を生成する (Phase 3)。
    """

    def __init__(self, output_base_dir: str):
        self.output_base_dir = Path(output_base_dir)

    def generate_prompts(self, pdf_path: str, in_base_dir: str = "data/in", grid_size: str = "small") -> dict:
        """
        Phase 1: PDFを解析し、LLMに渡すためのプロンプトを data/out/ に出力する。
        """
        logger.info(f"--- [Phase 1] PDF解析 & プロンプト生成: {Path(pdf_path).name} ---")
        path_obj = Path(pdf_path)
        pdf_name = path_obj.stem

        # 出力先のディレクトリを作成
        try:
            rel_path = path_obj.parent.relative_to(Path(in_base_dir))
            out_dir = self.output_base_dir / rel_path
        except ValueError:
            out_dir = self.output_base_dir / pdf_name

        out_dir.mkdir(parents=True, exist_ok=True)

        # PDFから情報を抽出
        extracted_data = extract_pdf_data(pdf_path)

        # ページ画像を PNG として書き出し（Step 1.6 視覚検証用）
        page_image_paths = []
        try:
            import pdfplumber
            with pdfplumber.open(pdf_path) as pdf_img:
                images_dir = out_dir / "images"
                images_dir.mkdir(parents=True, exist_ok=True)
                for i, pg in enumerate(pdf_img.pages, start=1):
                    img_path = images_dir / f"{pdf_name}_page{i}.png"
                    pg.to_image(resolution=150).save(str(img_path))
                    page_image_paths.append(img_path)
        except Exception as e:
            logger.warning(f"ページ画像の出力に失敗しました（Step 1.6 はスキップ可能）: {e}")

        first_page = extracted_data['pages'][0]
        grid_params = _setup_grid_params(first_page, grid_size)

        # Phase 3（render_excel）の罫線後処理で参照するためメタデータとして保存
        with open(out_dir / f"{pdf_name}_grid_params.json", "w", encoding="utf-8") as f:
            json.dump(grid_params, f, ensure_ascii=False)

        # Y・X座標のクラスタリングを行い、各要素に事前計算済みExcel座標を付与
        for page in extracted_data['pages']:
            _compute_grid_coords(page, grid_params['max_rows'], grid_params['max_cols'])

        # グリッド座標付与済みの状態で保存（_row/_col をデバッグ用ファイルに含める）
        extracted_json_path = out_dir / f"{pdf_name}_extracted.json"
        with open(extracted_json_path, "w", encoding="utf-8") as f:
            json.dump(extracted_data, f, indent=2, ensure_ascii=False)

        input_data_str = json.dumps(extracted_data, indent=2, ensure_ascii=False)

        # Step 1.5 用スリム版: words の text/_row/_col のみ（トークン削減）
        slim_data = {"pages": [
            {
                "page_number": p["page_number"],
                "words": [
                    {"text": w.get("text", ""), "_row": w["_row"], "_col": w["_col"]}
                    for w in p.get("words", [])
                ]
            }
            for p in extracted_data["pages"]
        ]}
        slim_input_data_str = json.dumps(slim_data, indent=2, ensure_ascii=False)

        # Step 1: 列アンカー確定プロンプト（PDF解析データを直接埋め込む）
        prompt_1 = TABLE_ANCHOR_PROMPT.format(
            input_data=input_data_str,
            **grid_params
        )

        # Step 1.5: レイアウトJSON検証・補正プロンプト（Step 1の出力を貼り付けるプレースホルダー）
        prompt_1_5 = LAYOUT_REVIEW_PROMPT.format(
            input_data=slim_input_data_str,
            step1_output="[ここにSTEP 1の出力（JSON部分のみ）を貼り付けてください]",
            **grid_params
        )

        # Step 1.6: 視覚検証プロンプト（Step 1.5の出力 + ページ画像を使って border_rect を修正）
        image_note = ""
        if page_image_paths:
            image_list = "\n".join(f"  - {p.name}" for p in page_image_paths)
            image_note = f"\n\n【画像ファイル】このプロンプトと一緒に以下の画像を LLM に添付してください:\n{image_list}"
        prompt_1_6 = VISUAL_BORDER_REVIEW_PROMPT.format(
            step1_5_output="[ここにSTEP 1.5の出力（JSON部分のみ）を貼り付けてください]"
        ) + image_note

        # Step 2: コード生成プロンプト（Step 1.5 or 1.6 の出力を貼り付けるプレースホルダー）
        prompt_2 = CODE_GEN_PROMPT.format(
            input_data="[ここにSTEP 1.5（または1.6）の出力（JSON部分のみ）を貼り付けてください]",
            **grid_params
        )

        # プロンプト保存
        prompts_dir = out_dir / "prompts"
        prompts_dir.mkdir(parents=True, exist_ok=True)

        prompt_1_path = prompts_dir / f"{pdf_name}_prompt_step1.txt"
        prompt_1_5_path = prompts_dir / f"{pdf_name}_prompt_step1_5.txt"
        prompt_1_6_path = prompts_dir / f"{pdf_name}_prompt_step1_6.txt"
        prompt_2_path = prompts_dir / f"{pdf_name}_prompt_step2.txt"

        with open(prompt_1_path, "w", encoding="utf-8") as f:
            f.write(prompt_1)
        with open(prompt_1_5_path, "w", encoding="utf-8") as f:
            f.write(prompt_1_5)
        with open(prompt_1_6_path, "w", encoding="utf-8") as f:
            f.write(prompt_1_6)
        with open(prompt_2_path, "w", encoding="utf-8") as f:
            f.write(prompt_2)

        # 生成コード保存用の空ファイルを作成
        generated_code_path = out_dir / f"{pdf_name}_gen.py"
        if not generated_code_path.exists():
            with open(generated_code_path, "w", encoding="utf-8") as f:
                f.write("# Please paste final AI Python code (from STEP 2) here.\n")

        logger.info(f"✅ Phase 1 完了: {pdf_name}")
        logger.info(f"  抽出データ: {extracted_json_path}")
        logger.info(f"  STEP 1   プロンプト: {prompt_1_path}")
        logger.info(f"  STEP 1.5 プロンプト: {prompt_1_5_path}")
        logger.info(f"  STEP 1.6 プロンプト: {prompt_1_6_path}（Vision LLM でページ画像と照合・罫線修正）")
        logger.info(f"  STEP 2   プロンプト: {prompt_2_path}")
        if page_image_paths:
            logger.info(f"  ページ画像: {', '.join(p.name for p in page_image_paths)}")
        logger.info(f"  ※ STEP1 → STEP1.5 → STEP1.6（任意・Vision LLM） → STEP2 → コードを {generated_code_path.name} に保存")

        return {
            "json_path": str(extracted_json_path),
            "prompt_step1_path": str(prompt_1_path),
            "prompt_step1_5_path": str(prompt_1_5_path),
            "prompt_step1_6_path": str(prompt_1_6_path),
            "prompt_step2_path": str(prompt_2_path),
            "generated_code_base_path": str(generated_code_path),
            "page_image_paths": [str(p) for p in page_image_paths],
        }

    def fill_layout(self, pdf_name: str, step1_5_json: str, specific_out_dir: str = None) -> str:
        """
        手動パイプライン用: STEP 1.5 の LLM 出力に対してプログラム的テキスト補完を適用する。

        STEP 1.5 の LLM が見落とした word を extracted_data と照合して補完し、
        補完済み JSON を prompts/{pdf_name}_step1_5_output.json として保存する。
        また、STEP 2 プロンプトを補完済み JSON で更新して保存する。

        Args:
            pdf_name:        PDF ファイル名（拡張子なし）
            step1_5_json:    STEP 1.5 の LLM 出力 JSON 文字列
            specific_out_dir: 出力ディレクトリ（省略時は data/out/{pdf_name}）

        Returns:
            補完済みレイアウト JSON 文字列
        """
        if specific_out_dir:
            out_dir = Path(specific_out_dir)
        else:
            out_dir = self.output_base_dir / pdf_name

        # extracted.json を読み込む（グリッド座標付与済み）
        extracted_json_path = out_dir / f"{pdf_name}_extracted.json"
        if not extracted_json_path.exists():
            raise FileNotFoundError(
                f"extracted.json が見つかりません: {extracted_json_path}. "
                "generate_prompts() を先に実行してください。"
            )

        with open(extracted_json_path, "r", encoding="utf-8") as f:
            extracted_data = json.load(f)

        # テキスト補完を適用
        filled_json = _fill_missing_text(step1_5_json, extracted_data)

        # 補完済み JSON を保存
        prompts_dir = out_dir / "prompts"
        prompts_dir.mkdir(parents=True, exist_ok=True)
        filled_json_path = prompts_dir / f"{pdf_name}_step1_5_output.json"
        with open(filled_json_path, "w", encoding="utf-8") as f:
            f.write(filled_json)
        logger.info(f"[fill_layout] 補完済みレイアウト JSON を保存しました: {filled_json_path}")

        # STEP 2 プロンプトのプレースホルダーを補完済み JSON で置換して保存
        prompt_2_path = prompts_dir / f"{pdf_name}_prompt_step2.txt"
        if prompt_2_path.exists():
            with open(prompt_2_path, "r", encoding="utf-8") as f:
                prompt_2 = f.read()
            placeholder = "[ここにSTEP 1.5（または1.6）の出力（JSON部分のみ）を貼り付けてください]"
            if placeholder in prompt_2:
                with open(prompt_2_path, "w", encoding="utf-8") as f:
                    f.write(prompt_2.replace(placeholder, filled_json))
                logger.info(f"[fill_layout] STEP 2 プロンプトを補完済み JSON で更新しました: {prompt_2_path}")

        return filled_json

    def render_excel(self, pdf_name: str, specific_out_dir: str = None) -> str:
        """
        Phase 3: AI出力の生成コードを読み込み、Excel方眼紙を描画する。
        """
        logger.info(f"--- [Phase 3] Excel生成: {pdf_name} ---")
        if specific_out_dir:
            out_dir = Path(specific_out_dir)
        else:
            out_dir = self.output_base_dir / pdf_name

        output_xlsx_path = out_dir / f"{pdf_name}.xlsx"
        generated_code_path = out_dir / f"{pdf_name}_gen.py"

        if generated_code_path.exists():
            with open(generated_code_path, "r", encoding="utf-8") as f:
                content = f.read().strip()

            code_lines = [line for line in content.splitlines() if not line.strip().startswith("#")]
            actual_code = "\n".join(code_lines).strip()
            is_placeholder = len(actual_code) < 50

            if content and not is_placeholder:
                # 既知の問題パターンを静的チェック・自動修正
                sanitized_content, fixes = _sanitize_generated_code(content)
                if fixes:
                    for fix in fixes:
                        logger.warning(f"🔧 静的修正: {fix}")
                    with open(generated_code_path, "w", encoding="utf-8") as f:
                        f.write(sanitized_content)
                    content = sanitized_content

                logger.info(f"✨ 生成されたコードを実行します: {generated_code_path.name}")
                import subprocess
                import os
                import sys

                try:
                    env = os.environ.copy()
                    env["PYTHONPATH"] = os.getcwd()

                    result = subprocess.run(
                        [sys.executable, generated_code_path.name],
                        cwd=str(out_dir),
                        env=env,
                        capture_output=True,
                        text=True
                    )

                    if result.returncode == 0:
                        temp_xlsx = out_dir / "output.xlsx"
                        if temp_xlsx.exists():
                            temp_xlsx.replace(output_xlsx_path)
                            # 罫線を Python で直接適用（LLM 生成コードの罫線を上書き補正）
                            extracted_json_path = out_dir / f"{pdf_name}_extracted.json"
                            grid_params_path = out_dir / f"{pdf_name}_grid_params.json"
                            if extracted_json_path.exists() and grid_params_path.exists():
                                try:
                                    with open(extracted_json_path, "r", encoding="utf-8") as f:
                                        extracted_data = json.load(f)
                                    with open(grid_params_path, "r", encoding="utf-8") as f:
                                        grid_params = json.load(f)
                                    _apply_borders_to_xlsx(
                                        str(output_xlsx_path), extracted_data, grid_params['max_rows']
                                    )
                                except Exception as e:
                                    logger.warning(f"罫線後処理に失敗しました（Excel は生成済み）: {e}")
                            logger.info(f"✅ Phase 3 完了: {output_xlsx_path}")
                            return str(output_xlsx_path)
                        else:
                            error_msg = "生成コードは正常終了しましたが、output.xlsx が生成されませんでした。"
                            logger.error(f"❌ {error_msg}")
                            self._generate_error_prompt(out_dir, pdf_name, error_msg, content)
                    else:
                        error_msg = f"生成コードの実行に失敗しました:\n{result.stderr}"
                        logger.error(f"❌ {error_msg}")
                        self._generate_error_prompt(out_dir, pdf_name, error_msg, content)
                except Exception as e:
                    error_msg = f"生成コード実行中に例外が発生しました: {e}"
                    logger.error(f"❌ {error_msg}")
                    self._generate_error_prompt(out_dir, pdf_name, error_msg, content)
            else:
                logger.warning(f"⚠️ 生成コードファイル {generated_code_path.name} が空、または未編集です。")
        else:
            logger.error(f"❌ 生成コードファイル {generated_code_path.name} が見つかりません。STEP 2 の結果を保存してください。")

        raise RuntimeError(f"Excelの生成に失敗しました ({pdf_name})")

    def _generate_error_prompt(self, out_dir: Path, pdf_name: str, error_msg: str, current_code: str):
        prompt_text = CODE_ERROR_FIXING_PROMPT.format(error_msg=error_msg, code=current_code)
        prompts_dir = out_dir / "prompts"
        prompts_dir.mkdir(parents=True, exist_ok=True)
        error_prompt_path = prompts_dir / f"{pdf_name}_prompt_error_fix.txt"
        with open(error_prompt_path, "w", encoding="utf-8") as f:
            f.write(prompt_text)
        logger.info(f"💡 エラー修正用プロンプトを出力しました: {error_prompt_path}")
