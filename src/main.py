import argparse
import os
from pathlib import Path

from dotenv import load_dotenv

from src.core.pipeline import SheetlingPipeline
from src.utils.logger import get_logger

load_dotenv()
logger = get_logger(__name__)


def main():
    parser = argparse.ArgumentParser(description="Auto-Sheetling: PDF to Excel conversion (with Gemini API automation)")
    parser.add_argument(
        "phase",
        choices=["extract", "generate", "auto"],
        help=(
            "Phase to run: "
            "extract (Phase 1: PDF解析 & プロンプト生成), "
            "generate (Phase 3: 生成コードを実行してExcel出力), "
            "auto (Phase 1〜3 を Gemini API で全自動実行)"
        )
    )
    parser.add_argument("--pdf", type=str, help="PDF file path. If not provided, processes all PDFs in data/in/")
    parser.add_argument(
        "--grid-size",
        type=str,
        choices=["small", "medium", "large"],
        default="small",
        help="Grid size for Excel layout (small=4mm, medium=6mm, large=8mm)"
    )
    parser.add_argument(
        "--vision",
        action="store_true",
        help="[auto mode only] Enable Step 1.6 Vision LLM for border_rect visual correction"
    )
    parser.add_argument(
        "--max-retries",
        type=int,
        default=3,
        help="[auto mode only] Maximum retries for Phase 3 on error (default: 3)"
    )
    parser.add_argument(
        "--model",
        type=str,
        default="gemini-2.0-flash",
        help="[auto mode only] Gemini model name (default: gemini-2.0-flash)"
    )
    args = parser.parse_args()

    if args.phase == "auto":
        # Gemini API キーを環境変数から取得
        api_key = os.environ.get("GEMINI_API_KEY")
        if not api_key:
            logger.error("❌ 環境変数 GEMINI_API_KEY が設定されていません。.env ファイルを確認してください。")
            return

        from src.core.auto_pipeline import AutoSheetlingPipeline
        pipeline = AutoSheetlingPipeline("data/out", api_key=api_key, model_name=args.model)

        if args.pdf:
            pdf_files = [Path(args.pdf)]
        else:
            pdf_files = list(Path("data/in").rglob("*.pdf"))

        if not pdf_files:
            logger.warning("No PDF files found in data/in. Please place PDF files to process.")
            return

        for pdf_path in pdf_files:
            try:
                output_path = pipeline.run(
                    str(pdf_path),
                    grid_size=args.grid_size,
                    use_vision_step=args.vision,
                    max_retries=args.max_retries,
                )
                logger.info(f"✅ 完了: {output_path}")
            except Exception as e:
                logger.error(f"❌ 自動パイプライン失敗 ({pdf_path.name}): {e}", exc_info=True)

    elif args.phase == "extract":
        pipeline = SheetlingPipeline("data/out")

        if args.pdf:
            pdf_files = [Path(args.pdf)]
        else:
            pdf_files = list(Path("data/in").rglob("*.pdf"))

        if not pdf_files:
            logger.warning("No PDF files found in data/in. Please place PDF files to process.")
            return

        for pdf_path in pdf_files:
            try:
                pipeline.generate_prompts(str(pdf_path), grid_size=args.grid_size)
            except Exception as e:
                logger.error(f"❌ Phase 1 failed for {pdf_path.name}: {e}", exc_info=True)

    elif args.phase == "generate":
        pipeline = SheetlingPipeline("data/out")
        output_base_dir = Path("data/out")

        gen_files = list(output_base_dir.rglob("*_gen.py"))

        generated_count = 0
        for gen_file in gen_files:
            out_dir = gen_file.parent

            if gen_file.name.endswith("_gen.py"):
                pdf_name = gen_file.name[:-7]
                generated_count += 1
                try:
                    pipeline.render_excel(pdf_name, specific_out_dir=str(out_dir))
                except Exception as e:
                    logger.error(f"❌ Phase 3 failed for {pdf_name}: {e}", exc_info=True)

        if generated_count == 0:
            logger.warning(f"No *_gen.py files found in subdirectories of {output_base_dir}. Please paste AI generated code first.")


if __name__ == "__main__":
    main()
