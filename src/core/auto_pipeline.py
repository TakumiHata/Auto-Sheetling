"""
Auto-Sheetling 自動パイプライン。
Gemini API を使って Phase 2（LLMとのやり取り）を自動化する。

フロー:
  Phase 1: PDF解析 → extracted_data.json
  Phase 2 (自動): Gemini API
    Step 1:   TABLE_ANCHOR_PROMPT   → レイアウトJSON
    Step 1.5: LAYOUT_REVIEW_PROMPT  → 補正済みレイアウトJSON
    Step 1.6: VISUAL_BORDER_REVIEW_PROMPT（オプション・Vision）→ 視覚補正JSON
    Step 2:   CODE_GEN_PROMPT       → Python(openpyxl)コード
  Phase 3: 生成コードを実行 → Excel方眼紙（エラー時は自動リトライ）
"""

import json
import re
from pathlib import Path

import google.generativeai as genai

from src.core.pipeline import SheetlingPipeline, _compute_grid_coords, _sanitize_generated_code
from src.parser.pdf_extractor import extract_pdf_data
from src.templates.prompts import (
    TABLE_ANCHOR_PROMPT,
    LAYOUT_REVIEW_PROMPT,
    VISUAL_BORDER_REVIEW_PROMPT,
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

    def __init__(self, output_base_dir: str, api_key: str, model_name: str = "gemini-2.0-flash"):
        super().__init__(output_base_dir)
        genai.configure(api_key=api_key)
        self.model_name = model_name
        self.model = genai.GenerativeModel(model_name)
        logger.info(f"Gemini モデル '{model_name}' を使用します。")

    def _call_gemini(self, prompt: str, images: list = None) -> str:
        """Gemini API を呼び出してテキストを生成する。"""
        if images:
            contents = [prompt] + images
        else:
            contents = prompt
        response = self.model.generate_content(contents)
        return response.text

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
                    pg.to_image(resolution=150).save(str(img_path))
                    page_image_paths.append(img_path)
        except Exception as e:
            logger.warning(f"ページ画像の出力に失敗しました: {e}")

        # 抽出データを保存
        extracted_json_path = out_dir / f"{pdf_name}_extracted.json"
        with open(extracted_json_path, "w", encoding="utf-8") as f:
            json.dump(extracted_data, f, indent=2, ensure_ascii=False)
        logger.info(f"[Phase 1] 完了: {extracted_json_path}")

        # グリッドパラメータ設定
        grid_params = dict(GRID_SIZES.get(grid_size, GRID_SIZES["small"]))
        first_page = extracted_data['pages'][0]
        is_landscape = first_page['width'] > first_page['height']
        if is_landscape:
            grid_params['max_rows'], grid_params['max_cols'] = grid_params['max_cols'], grid_params['max_rows']
            _row_padding = 2
            grid_params['excel_row_height'] = round(595.0 / (grid_params['max_rows'] + _row_padding), 2)
        grid_params['orientation'] = 'landscape' if is_landscape else 'portrait'

        # グリッド座標付与
        for page in extracted_data['pages']:
            _compute_grid_coords(page, grid_params['max_rows'], grid_params['max_cols'])

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
