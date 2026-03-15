"""
Auto-Sheetling 自動パイプライン。
Gemini API を使って Phase 2（LLMとのやり取り）を自動化する。

フロー:
  Phase 1: PDF解析 → extracted_data.json
  Phase 2 (自動): Gemini API
    Step 1:   TABLE_ANCHOR_PROMPT   → レイアウトJSON
    Step 1.5: LAYOUT_REVIEW_PROMPT         → 補正済みレイアウトJSON
    Step 1.6b: BORDER_VISION_REVIEW_PROMPT（オプション・Vision）→ table_border_rects._borders 補正
    Step 1.6: VISUAL_BORDER_REVIEW_PROMPT （オプション・Vision）→ 視覚補正JSON
    Step 2:   CODE_GEN_PROMPT             → Python(openpyxl)コード
  Phase 3: 生成コードを実行 → Excel方眼紙（エラー時は自動リトライ）
"""

import json
import re
import time
import random
from pathlib import Path

import google.generativeai as genai

from src.core.pipeline import SheetlingPipeline, _compute_grid_coords, _sanitize_generated_code, _setup_grid_params, _fill_missing_text
from src.parser.pdf_extractor import extract_pdf_data
from src.templates.prompts import (
    TABLE_ANCHOR_PROMPT,
    LAYOUT_REVIEW_PROMPT,
    VISUAL_BORDER_REVIEW_PROMPT,
    BORDER_VISION_REVIEW_PROMPT,
    CODE_GEN_PROMPT,
    CODE_ERROR_FIXING_PROMPT,
    GRID_SIZES,
)
from src.utils.logger import get_logger

logger = get_logger(__name__)


def _extract_json(text: str) -> str:
    """LLMのレスポンスからJSON配列部分を抽出する。"""
    # コードブロックのマーカーを除去
    text = re.sub(r'```(?:json)?\s*', '', text)
    text = re.sub(r'```\s*', '', text)
    # [ ... ] の範囲を抽出
    start = text.find('[')
    end = text.rfind(']')
    if start != -1 and end != -1 and end > start:
        return text[start:end + 1]
    return text.strip()


def _extract_code(text: str) -> str:
    """LLMのレスポンスから ```python ... ``` ブロックのコードを抽出する。"""
    match = re.search(r'```python\s*(.*?)\s*```', text, re.DOTALL)
    if match:
        return match.group(1)
    # フォールバック: コードブロックなしの場合はそのまま返す
    return text.strip()


class AutoSheetlingPipeline(SheetlingPipeline):
    """
    Gemini API を使って Phase 2（LLMとのやり取り）を全自動化したパイプライン。
    Phase 1 と Phase 3 は親クラス SheetlingPipeline をそのまま利用する。
    """

    def __init__(self, output_base_dir: str, api_key: str, model_name: str = "gemini-3.1-flash-lite-preview"):
        super().__init__(output_base_dir)
        genai.configure(api_key=api_key)
        self.model_name = model_name
        # max_output_tokens を明示設定: デフォルト(8192)だと大規模PDFの
        # コード生成レスポンスが切れて 500 Internal Error になることがある
        self.model = genai.GenerativeModel(
            model_name,
            generation_config={"max_output_tokens": 65536},
        )
        logger.info(f"Gemini モデル '{model_name}' を使用します。")

    # API 呼び出しタイムアウト（秒）
    API_TIMEOUT = 180

    def _call_gemini(self, prompt: str, images: list = None) -> str:
        """Gemini API を呼び出してテキストを生成する（クォータエラー時はリトライする）。

        grpc のブロッキング呼び出しにタイムアウトが効かないケースがあるため、
        concurrent.futures による Python レベルのタイムアウトも併用する。
        """
        import concurrent.futures

        if images:
            contents = [prompt] + images
        else:
            contents = prompt

        max_retries = 5
        base_delay = 5
        request_options = {"timeout": self.API_TIMEOUT}

        for attempt in range(max_retries):
            try:
                with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
                    future = executor.submit(
                        self.model.generate_content, contents,
                        request_options=request_options
                    )
                    try:
                        response = future.result(timeout=self.API_TIMEOUT + 10)
                    except concurrent.futures.TimeoutError:
                        raise TimeoutError(
                            f"Gemini API が {self.API_TIMEOUT} 秒以内に応答しませんでした"
                        )
                return response.text
            except Exception as e:
                err_str = str(e)
                # 429 Resource Exhausted / タイムアウト / 500 Internal Error はリトライ対象
                # 500 は一時的なサーバー側エラーで再試行により回復することが多い
                is_retryable = (
                    "429" in err_str
                    or "500" in err_str
                    or "ResourceExhausted" in type(e).__name__
                    or "InternalServerError" in type(e).__name__
                    or "DeadlineExceeded" in type(e).__name__
                    or "TimeoutError" in type(e).__name__
                    or "timeout" in err_str.lower()
                )
                if is_retryable and attempt < max_retries - 1:
                    delay = base_delay * (2 ** attempt) + random.uniform(0, 1)
                    logger.warning(
                        f"API リトライ ({attempt + 1}/{max_retries}): "
                        f"{delay:.2f}秒後に再試行します... ({type(e).__name__})"
                    )
                    time.sleep(delay)
                    continue
                logger.error(f"Gemini API 呼び出し中にエラーが発生しました: {e}")
                raise
        raise RuntimeError("Gemini API のリトライ回数が上限に達しました。")

    # PDF 画像レンダリング解像度（auto_pipeline 全体で統一）
    RENDER_DPI = 200

    def _annotate_page_image(self, image, page_data: dict) -> "PIL.Image.Image":
        """
        PDF ページ画像に table_border_rects の境界を赤枠で描画して返す。
        _pdf_x0/top/x1/bottom フィールドをピクセル座標に変換して描画する。
        """
        try:
            from PIL import ImageDraw
        except ImportError:
            return image

        pt_to_px = self.RENDER_DPI / 72.0
        annotated = image.copy()
        draw = ImageDraw.Draw(annotated)

        for tbr in page_data.get("table_border_rects", []):
            if not all(k in tbr for k in ('_pdf_x0', '_pdf_top', '_pdf_x1', '_pdf_bottom')):
                continue
            x0 = tbr['_pdf_x0']  * pt_to_px
            y0 = tbr['_pdf_top'] * pt_to_px
            x1 = tbr['_pdf_x1']  * pt_to_px
            y1 = tbr['_pdf_bottom'] * pt_to_px
            draw.rectangle([x0, y0, x1, y1], outline='red', width=1)

        return annotated

    def _correct_borders_with_vision(
        self,
        extracted_data: dict,
        page_image_paths: list,
        grid_params: dict,
        out_dir: Path,
        prompts_dir: Path,
        pdf_name: str,
    ) -> dict:
        """
        アノテーション済み PDF ページ画像と table_border_rects を Vision LLM に渡し、
        差分パッチ形式で _borders を補正して extracted_data を更新する。

        補正に失敗した場合は元の extracted_data をそのまま返す。
        """
        try:
            import PIL.Image
            raw_images = [PIL.Image.open(str(p)) for p in page_image_paths]
        except Exception as e:
            logger.warning(f"[Border Vision] 画像読み込み失敗: {e}")
            return extracted_data

        # 各ページの画像にセル境界アノテーションを描画
        annotated_images = []
        for raw_img, page in zip(raw_images, extracted_data["pages"]):
            ann = self._annotate_page_image(raw_img, page)
            annotated_images.append(ann)
            # デバッグ用にアノテーション済み画像を保存
            try:
                ann_path = out_dir / "images" / f"{pdf_name}_page{page['page_number']}_annotated.png"
                ann.save(str(ann_path))
            except Exception:
                pass

        # LLM に渡す JSON: _pdf_* 座標フィールドは除外（画像アノテーションで代替）
        def _strip_pdf_coords(cell: dict) -> dict:
            return {k: v for k, v in cell.items() if not k.startswith('_pdf_')}

        border_data = json.dumps([
            {
                "page_number": page["page_number"],
                "table_border_rects": [_strip_pdf_coords(c) for c in page.get("table_border_rects", [])],
            }
            for page in extracted_data["pages"]
        ], ensure_ascii=False)

        prompt = BORDER_VISION_REVIEW_PROMPT.format(
            border_data=border_data,
            max_rows=grid_params["max_rows"],
            max_cols=grid_params["max_cols"],
        )
        with open(prompts_dir / f"{pdf_name}_prompt_border_vision.txt", "w", encoding="utf-8") as f:
            f.write(prompt)

        try:
            raw = self._call_gemini(prompt, images=annotated_images)
            patch_pages = json.loads(_extract_json(raw))

            # パッチ適用: page_number + (_row, _end_row, _col, _end_col) でセルを特定して _borders を上書き
            page_map = {p["page_number"]: p for p in extracted_data["pages"]}
            corrected_count = 0
            for patch_page in patch_pages:
                page_num = patch_page.get("page_number", 1)
                corrections = patch_page.get("corrections", [])
                if page_num not in page_map or not corrections:
                    continue
                tbrs = page_map[page_num].get("table_border_rects", [])
                # (row, end_row, col, end_col) をキーにしたセル辞書を構築
                cell_index = {
                    (c['_row'], c['_end_row'], c['_col'], c['_end_col']): c
                    for c in tbrs
                }
                for patch in corrections:
                    key = (patch.get('_row'), patch.get('_end_row'), patch.get('_col'), patch.get('_end_col'))
                    if key in cell_index and '_borders' in patch:
                        cell_index[key]['_borders'] = patch['_borders']
                        corrected_count += 1

            logger.info(f"[Border Vision] {corrected_count} 個のセル罫線データを補正しました")
        except Exception as e:
            logger.warning(f"[Border Vision] 補正に失敗しました。元データを使用します: {e}")

        return extracted_data

    def run(
        self,
        pdf_path: str,
        in_base_dir: str = "data/in",
        grid_size: str = "small",
        use_vision_step: bool = False,
        max_retries: int = 3,
    ) -> str:
        """
        Phase 1〜3 を全自動で実行する。

        Args:
            pdf_path:        処理対象のPDFファイルパス
            in_base_dir:     PDFの入力ベースディレクトリ（出力先の相対パス計算に使用）
            grid_size:       グリッドサイズ ("small" / "medium" / "large")
            use_vision_step: Step 1.6 の Vision LLM補正を使用するか
            max_retries:     Phase 3 失敗時の最大リトライ回数

        Returns:
            生成された Excel ファイルのパス
        """
        logger.info(f"=== 自動パイプライン開始: {Path(pdf_path).name} ===")
        path_obj = Path(pdf_path)
        pdf_name = path_obj.stem

        # --- 出力ディレクトリ設定 ---
        try:
            rel_path = path_obj.parent.relative_to(Path(in_base_dir))
            out_dir = self.output_base_dir / rel_path
        except ValueError:
            out_dir = self.output_base_dir / pdf_name
        out_dir.mkdir(parents=True, exist_ok=True)
        prompts_dir = out_dir / "prompts"
        prompts_dir.mkdir(parents=True, exist_ok=True)

        # -----------------------------------------------------------------------
        # Phase 1: PDF解析
        # -----------------------------------------------------------------------
        logger.info("[Phase 1] PDF解析中...")
        extracted_data = extract_pdf_data(pdf_path)

        # ページ画像を生成（Step 1.6 / Vision用）
        page_image_paths = []
        try:
            import pdfplumber
            with pdfplumber.open(pdf_path) as pdf_img:
                images_dir = out_dir / "images"
                images_dir.mkdir(parents=True, exist_ok=True)
                for i, pg in enumerate(pdf_img.pages, start=1):
                    img_path = images_dir / f"{pdf_name}_page{i}.png"
                    pg.to_image(resolution=self.RENDER_DPI).save(str(img_path))
                    page_image_paths.append(img_path)
        except Exception as e:
            logger.warning(f"ページ画像の出力に失敗しました: {e}")

        # グリッドパラメータ設定・座標付与（extracted.json 保存前に実施して _row/_col を含める）
        first_page = extracted_data['pages'][0]
        grid_params = _setup_grid_params(first_page, grid_size)

        for page in extracted_data['pages']:
            _compute_grid_coords(page, grid_params['max_rows'], grid_params['max_cols'])

        # Phase 3（render_excel）の罫線後処理で参照するためメタデータとして保存
        with open(out_dir / f"{pdf_name}_grid_params.json", "w", encoding="utf-8") as f:
            json.dump(grid_params, f, ensure_ascii=False)

        # 抽出データを保存（グリッド座標付与済み）
        extracted_json_path = out_dir / f"{pdf_name}_extracted.json"
        with open(extracted_json_path, "w", encoding="utf-8") as f:
            json.dump(extracted_data, f, indent=2, ensure_ascii=False)
        logger.info(f"[Phase 1] 完了: {extracted_json_path}")

        # Step 1.6b (オプション): Vision で table_border_rects の _borders を補正
        if use_vision_step and page_image_paths:
            logger.info("[Phase 2 / Border Vision] PDF画像で罫線データを補正中...")
            extracted_data = self._correct_borders_with_vision(
                extracted_data, page_image_paths, grid_params, out_dir, prompts_dir, pdf_name
            )
            # 補正済みデータで extracted.json を上書き保存
            with open(extracted_json_path, "w", encoding="utf-8") as f:
                json.dump(extracted_data, f, indent=2, ensure_ascii=False)
            logger.info("[Phase 2 / Border Vision] 完了")

        input_data_str = json.dumps(extracted_data, indent=2, ensure_ascii=False)

        # Step 1.5 用スリム版（トークン削減）
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

        # -----------------------------------------------------------------------
        # Phase 2: Gemini API による自動処理
        # -----------------------------------------------------------------------

        # Step 1: TABLE_ANCHOR_PROMPT → レイアウトJSON
        logger.info("[Phase 2 / Step 1] Gemini API でレイアウトJSON生成中...")
        prompt_1 = TABLE_ANCHOR_PROMPT.format(input_data=input_data_str, **grid_params)
        with open(prompts_dir / f"{pdf_name}_prompt_step1.txt", "w", encoding="utf-8") as f:
            f.write(prompt_1)

        step1_raw = self._call_gemini(prompt_1)
        step1_output = _extract_json(step1_raw)
        with open(prompts_dir / f"{pdf_name}_step1_output.json", "w", encoding="utf-8") as f:
            f.write(step1_output)
        logger.info("[Phase 2 / Step 1] 完了")

        # Step 1.5: LAYOUT_REVIEW_PROMPT → 補正済みレイアウトJSON
        logger.info("[Phase 2 / Step 1.5] Gemini API でレイアウトJSON検証・補正中...")
        prompt_1_5 = LAYOUT_REVIEW_PROMPT.format(
            input_data=slim_input_data_str,
            step1_output=step1_output,
            **grid_params
        )
        with open(prompts_dir / f"{pdf_name}_prompt_step1_5.txt", "w", encoding="utf-8") as f:
            f.write(prompt_1_5)

        step1_5_raw = self._call_gemini(prompt_1_5)
        step1_5_output = _extract_json(step1_5_raw)
        # LLM の見落とし補完: extracted_data の全 words と照合して欠落テキストを補充
        step1_5_output = _fill_missing_text(step1_5_output, extracted_data)
        with open(prompts_dir / f"{pdf_name}_step1_5_output.json", "w", encoding="utf-8") as f:
            f.write(step1_5_output)
        logger.info("[Phase 2 / Step 1.5] 完了")

        final_json = step1_5_output

        # Step 1.6 (オプション): VISUAL_BORDER_REVIEW_PROMPT → Vision補正JSON
        if use_vision_step and page_image_paths:
            logger.info("[Phase 2 / Step 1.6] Gemini Vision API で border_rect 視覚補正中...")
            prompt_1_6 = VISUAL_BORDER_REVIEW_PROMPT.format(step1_5_output=step1_5_output)
            with open(prompts_dir / f"{pdf_name}_prompt_step1_6.txt", "w", encoding="utf-8") as f:
                f.write(prompt_1_6)
            try:
                import PIL.Image
                images = [PIL.Image.open(str(p)) for p in page_image_paths]
                step1_6_raw = self._call_gemini(prompt_1_6, images=images)
                step1_6_output = _extract_json(step1_6_raw)
                with open(prompts_dir / f"{pdf_name}_step1_6_output.json", "w", encoding="utf-8") as f:
                    f.write(step1_6_output)
                final_json = step1_6_output
                logger.info("[Phase 2 / Step 1.6] 完了")
            except Exception as e:
                logger.warning(f"Step 1.6 (Vision) に失敗しました。Step 1.5 の結果を使用します: {e}")

        # Step 2: CODE_GEN_PROMPT → Pythonコード
        logger.info("[Phase 2 / Step 2] Gemini API で Excel 生成コード生成中...")
        prompt_2 = CODE_GEN_PROMPT.format(input_data=final_json, **grid_params)
        with open(prompts_dir / f"{pdf_name}_prompt_step2.txt", "w", encoding="utf-8") as f:
            f.write(prompt_2)

        step2_raw = self._call_gemini(prompt_2)
        generated_code = _extract_code(step2_raw)
        generated_code_path = out_dir / f"{pdf_name}_gen.py"
        with open(generated_code_path, "w", encoding="utf-8") as f:
            f.write(generated_code)
        logger.info(f"[Phase 2 / Step 2] 完了: {generated_code_path.name}")

        # -----------------------------------------------------------------------
        # Phase 3: Excel生成（エラー時は自動リトライ）
        # -----------------------------------------------------------------------
        for attempt in range(1, max_retries + 1):
            logger.info(f"[Phase 3] Excel生成 (試行 {attempt}/{max_retries})...")
            try:
                return self.render_excel(pdf_name, specific_out_dir=str(out_dir))
            except RuntimeError:
                if attempt >= max_retries:
                    logger.error(f"❌ Excel生成が {max_retries} 回失敗しました: {pdf_name}")
                    raise

                # エラー修正プロンプトを読んで Gemini に自動修正させる
                error_prompt_path = prompts_dir / f"{pdf_name}_prompt_error_fix.txt"
                if not error_prompt_path.exists():
                    logger.error("エラー修正プロンプトが見つかりません。リトライを中止します。")
                    raise

                logger.info(f"[Phase 3] Gemini API でコードを自動修正中 (試行 {attempt})...")
                with open(error_prompt_path, "r", encoding="utf-8") as f:
                    error_prompt = f.read()

                fixed_raw = self._call_gemini(error_prompt)
                fixed_code = _extract_code(fixed_raw)
                with open(generated_code_path, "w", encoding="utf-8") as f:
                    f.write(fixed_code)
                logger.info("コードを修正しました。再実行します...")

        raise RuntimeError(f"Excelの生成に失敗しました ({pdf_name})")
