"""
GPT Judge API Caller (Root-dir Version, Per-model Excel)

- If input is a JSONL file: process that file and write to --output (single Excel).
- If input is a root directory: for each model subdir (optionally filtered by --models),
  read <jsonl-name> under that model dir, and write results to:
      <model_dir>/<output_excel_name>
  (i.e., each model gets its own Excel)

Examples:
  # Single JSONL (original behavior)
  python call_gpt_judge.py path/to/content_parts.jsonl -o results.xlsx --api-key ...

  # Root dir: process all models, each model writes its own results.xlsx
  python call_gpt_judge.py data/eval_dataset_opus -o results.xlsx --api-key ...

  # Root dir: process only selected models
  python call_gpt_judge.py data/eval_dataset_opus --models opus_4.5_output,model2 -o results.xlsx --api-key ...
"""

import argparse
import json
import logging
import time
from pathlib import Path
from typing import Dict, Any, List, Optional
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook
from openai import AzureOpenAI



# ==================== Configuration ====================

class APIConfig:
    # Azure OpenAI settings (code1 style)
    AZURE_ENDPOINT = ""          # e.g. "https://xxx.openai.azure.com/"
    API_KEY = ""                 # your azure api key
    API_VERSION = ""

    # In Azure OpenAI, this is typically your *deployment name*
    MODEL = ""

    MAX_TOKENS = 2000
    MAX_COMPLETION_TOKENS=128000
    TEMPERATURE = 1.0
    TIMEOUT = 120

    MAX_RETRIES = 3
    RETRY_DELAY = 5

    RATE_LIMIT_DELAY = 1.0


# ==================== Excel Manager ====================

class ExcelResultManager:
    """Manages incremental writing of results to Excel file."""

    def __init__(self, output_path: Path):
        self.output_path = Path(output_path)
        self.columns = [
            "task_id",
            "score",
            "detailed_analysis",
            "timestamp",
            "model",
            "api_call_duration",
            "error",
            "source_jsonl",
            "line_num",
        ]

        if not self.output_path.exists():
            self._create_new_file()

        logging.info(f"Excel results will be saved to: {self.output_path}")

    def _create_new_file(self) -> None:
        df = pd.DataFrame(columns=self.columns)
        df.to_excel(self.output_path, index=False, sheet_name="Results")
        logging.info(f"Created new Excel file: {self.output_path}")

    def append_result(self, result: Dict[str, Any]) -> None:
        book = load_workbook(self.output_path)
        sheet = book["Results"]

        row_data = [
            result.get("task_id", ""),
            result.get("score", ""),
            result.get("detailed_analysis", ""),
            result.get("timestamp", ""),
            result.get("model", ""),
            result.get("api_call_duration", ""),
            result.get("error", ""),
            result.get("source_jsonl", ""),
            result.get("line_num", ""),
        ]

        sheet.append(row_data)
        book.save(self.output_path)
        book.close()

        logging.info(f"✓ Appended result for task: {result.get('task_id')}")

    def get_processed_task_ids(self) -> set:
        try:
            df = pd.read_excel(self.output_path, sheet_name="Results")
            return set(df["task_id"].dropna().astype(str).tolist())
        except Exception as e:
            logging.warning(f"Could not read processed tasks from {self.output_path}: {e}")
            return set()


# ==================== GPT Judge Caller ====================

class GPTJudgeCaller:
    def __init__(self, config: APIConfig):
        self.config = config
        self.client = AzureOpenAI(
            azure_endpoint=config.AZURE_ENDPOINT,
            api_key=config.API_KEY,
            api_version=config.API_VERSION,
            timeout=config.TIMEOUT,
        )
        logging.info(f"Initialized GPT Judge with Azure deployment/model: {config.MODEL}")


    def call_api(self, content_parts: List[Dict[str, Any]]) -> Dict[str, Any]:
        start_time = time.time()

        for attempt in range(self.config.MAX_RETRIES):
            try:
                messages = [{"role": "user", "content": content_parts}]

                response = self.client.chat.completions.create(
                    model=self.config.MODEL,
                    messages=messages,
                    max_completion_tokens=self.config.MAX_COMPLETION_TOKENS,
                    temperature=self.config.TEMPERATURE
                )

                response_text = response.choices[0].message.content
                result = self._parse_response(response_text)

                result["api_call_duration"] = time.time() - start_time
                result["model"] = self.config.MODEL
                return result

            except Exception as e:
                logging.warning(f"API call attempt {attempt + 1} failed: {e}")
                if attempt < self.config.MAX_RETRIES - 1:
                    time.sleep(self.config.RETRY_DELAY)
                else:
                    return {
                        "score": None,
                        "detailed_analysis": None,
                        "error": str(e),
                        "api_call_duration": time.time() - start_time,
                        "model": self.config.MODEL
                    }

    def _parse_response(self, response_text: str) -> Dict[str, Any]:
        """
        Parse model response.

        If parsing fails, treat as score=0 (fail) and preserve raw output in detailed_analysis.
        """
        def _sanitize_json_text(text: str) -> str:
            # Escape control characters and invalid escapes inside JSON strings.
            out = []
            in_string = False
            i = 0
            valid_escapes = set('"\\/bfnrtu')
            while i < len(text):
                ch = text[i]
                if ch == "\"":
                    in_string = not in_string
                    out.append(ch)
                    i += 1
                    continue
                if in_string and ch == "\\":
                    nxt = text[i + 1] if i + 1 < len(text) else ""
                    if nxt and nxt in valid_escapes:
                        out.append(ch)
                        out.append(nxt)
                        i += 2
                        continue
                    # Invalid escape, treat backslash as literal
                    out.append("\\\\")
                    i += 1
                    continue
                if in_string:
                    code = ord(ch)
                    if ch == "\n":
                        out.append("\\n")
                        i += 1
                        continue
                    if ch == "\r":
                        out.append("\\r")
                        i += 1
                        continue
                    if ch == "\t":
                        out.append("\\t")
                        i += 1
                        continue
                    if code < 32:
                        out.append(f"\\u{code:04x}")
                        i += 1
                        continue
                out.append(ch)
                i += 1
            return "".join(out)

        try:
            result = json.loads(response_text)
            if "score" in result and "detailed_analysis" in result:
                return {
                    "score": result["score"],
                    "detailed_analysis": result["detailed_analysis"],
                    "error": None
                }
            # Missing required fields => count as fail
            return {
                "score": 0,
                "detailed_analysis": response_text,
                "error": "Response missing required fields"
            }

        except json.JSONDecodeError:
            import re
            m = re.search(r"```json\s*(\{.*?\})\s*```", response_text, re.DOTALL)
            if m:
                try:
                    fixed = _sanitize_json_text(m.group(1))
                    result = json.loads(fixed)
                    if "score" in result and "detailed_analysis" in result:
                        return {
                            "score": result.get("score"),
                            "detailed_analysis": result.get("detailed_analysis"),
                            "error": None
                        }
                    # JSON extracted but missing fields => fail
                    return {
                        "score": 0,
                        "detailed_analysis": response_text,
                        "error": "Response missing required fields"
                    }
                except Exception:
                    pass

            # Last attempt: sanitize whole response and try JSON parse
            try:
                fixed = _sanitize_json_text(response_text)
                result = json.loads(fixed)
                if "score" in result and "detailed_analysis" in result:
                    return {
                        "score": result.get("score"),
                        "detailed_analysis": result.get("detailed_analysis"),
                        "error": None
                    }
            except Exception:
                pass

            # Unparseable => fail with score=0
            return {
                "score": 0,
                "detailed_analysis": response_text,
                "error": "Failed to parse JSON response"
            }



# ==================== Processor ====================

class JudgeProcessor:
    """
    Processes:
    - single JSONL file (writes to a single Excel), OR
    - one model directory containing content_parts.jsonl (writes to that model's Excel)
    """

    def __init__(
        self,
        jsonl_path: Path,
        output_excel: Path,
        config: APIConfig,
        skip_processed: bool = True,
        task_id_prefix_mode: str = "none",  # "none" or "model"
    ):
        self.jsonl_path = Path(jsonl_path)
        self.output_excel = Path(output_excel)
        self.config = config
        self.skip_processed = skip_processed
        self.task_id_prefix_mode = task_id_prefix_mode

        self.caller = GPTJudgeCaller(config)
        self.excel_manager = ExcelResultManager(self.output_excel)

        self.processed_tasks = self.excel_manager.get_processed_task_ids() if skip_processed else set()
        logging.info(f"Found {len(self.processed_tasks)} already processed tasks in: {self.output_excel}")

    def _compose_task_id(self, raw_task_id: str) -> str:
        if self.task_id_prefix_mode == "model":
            model_name = self.jsonl_path.parent.name
            return f"{model_name}::{raw_task_id}"
        return raw_task_id

    def process_all(self) -> None:
        if not self.jsonl_path.exists():
            logging.error(f"Input JSONL not found: {self.jsonl_path}")
            return

        total_lines = 0
        with self.jsonl_path.open("r", encoding="utf-8") as f:
            for _ in f:
                total_lines += 1

        logging.info(f"Processing {total_lines} lines from: {self.jsonl_path}")

        processed_count = 0
        skipped_count = 0
        error_count = 0

        with self.jsonl_path.open("r", encoding="utf-8") as f:
            for line_num, line in enumerate(f, 1):
                data = None
                try:
                    data = json.loads(line.strip())

                    raw_task_id = data.get("task_id", f"task_{line_num}")
                    task_id = self._compose_task_id(raw_task_id)

                    if self.skip_processed and task_id in self.processed_tasks:
                        logging.info(f"[{line_num}/{total_lines}] Skipping processed: {task_id}")
                        skipped_count += 1
                        continue

                    content_parts = data.get("content_parts", [])
                    logging.info(f"[{line_num}/{total_lines}] Judging: {task_id}")

                    result = self.caller.call_api(content_parts)

                    excel_result = {
                        "task_id": task_id,
                        "score": result.get("score"),
                        "detailed_analysis": result.get("detailed_analysis"),
                        "timestamp": datetime.now().isoformat(),
                        "model": result.get("model", self.config.MODEL),
                        "api_call_duration": result.get("api_call_duration", 0),
                        "error": result.get("error"),
                        "source_jsonl": str(self.jsonl_path),
                        "line_num": line_num,
                    }

                    self.excel_manager.append_result(excel_result)
                    processed_count += 1

                    if result.get("error"):
                        error_count += 1
                        logging.warning(f"Task {task_id} completed with error: {result.get('error')}")

                    time.sleep(self.config.RATE_LIMIT_DELAY)

                except Exception as e:
                    error_count += 1
                    logging.error(f"Failed at {self.jsonl_path} line {line_num}: {e}", exc_info=True)

                    # Best-effort: append error row
                    try:
                        raw_task_id = (data or {}).get("task_id", f"task_{line_num}")
                        task_id = self._compose_task_id(raw_task_id)
                        self.excel_manager.append_result({
                            "task_id": task_id,
                            "score": None,
                            "detailed_analysis": None,
                            "timestamp": datetime.now().isoformat(),
                            "model": self.config.MODEL,
                            "api_call_duration": 0,
                            "error": str(e),
                            "source_jsonl": str(self.jsonl_path),
                            "line_num": line_num,
                        })
                    except Exception:
                        pass

        logging.info(f"\nSummary for {self.jsonl_path}:")
        logging.info(f"  Total: {total_lines}")
        logging.info(f"  Processed: {processed_count}")
        logging.info(f"  Skipped: {skipped_count}")
        logging.info(f"  Errors: {error_count}")
        logging.info(f"  Output: {self.output_excel}")


# ==================== Root-dir Helpers ====================

def iter_model_dirs(root_dir: Path, models: Optional[List[str]]) -> List[Path]:
    model_dirs = [d for d in root_dir.iterdir() if d.is_dir()]
    if models:
        allow = set(models)
        model_dirs = [d for d in model_dirs if d.name in allow]
    return sorted(model_dirs)


def resolve_model_output_excel(model_dir: Path, output_arg: str) -> Path:
    """
    output_arg can be:
      - "results.xlsx" (recommended) => model_dir/results.xlsx
      - "sub/results.xlsx" (relative) => model_dir/sub/results.xlsx
      - "/abs/path/results.xlsx" (absolute) => still write per model by using filename under model_dir
        (to avoid mixing models into one shared absolute file)
    """
    out_path = Path(output_arg)

    # Always write per model directory; absolute output is treated as a filename template only.
    if out_path.is_absolute():
        return model_dir / out_path.name

    # Relative path: place under model_dir
    return model_dir / out_path


# ==================== CLI Entry Point ====================

def main():
    parser = argparse.ArgumentParser(description="Call GPT Judge API for content evaluation (root-dir + per-model Excel).")
    parser.add_argument(
        "input",
        type=str,
        help="Path to input JSONL file OR root dir containing model subdirs with content_parts.jsonl"
    )
    parser.add_argument(
        "-o", "--output",
        type=str,
        default="results.xlsx",
        help="Output Excel filename/path (default: results.xlsx). "
             "In root-dir mode, this is used as a per-model path under each model dir."
    )
    parser.add_argument("--api-key", type=str, help="OpenAI API key (overrides config)")
    parser.add_argument("--azure-endpoint", type=str, help="Azure OpenAI endpoint (overrides config)")
    parser.add_argument("--api-version", type=str, help="Azure OpenAI api_version (overrides config)")
    parser.add_argument("--model", type=str, help="Azure deployment name (overrides config)")


    parser.add_argument(
        "--models",
        type=str,
        default=None,
        help="Comma-separated model subdir names to process (root-dir mode only). Example: opus_4.5_output,model2"
    )
    parser.add_argument(
        "--jsonl-name",
        type=str,
        default="content_parts.jsonl",
        help="JSONL filename under each model directory (default: content_parts.jsonl)"
    )
    parser.add_argument(
        "--task-id-prefix",
        type=str,
        choices=["none", "model"],
        default="none",
        help="Prefix task_id with model dir (default: none). Useful only if you want uniqueness."
    )

    parser.add_argument("--no-skip-processed", action="store_true", help="Reprocess all tasks")
    parser.add_argument("--max-retries", type=int, help="Maximum number of retries on API failure")
    parser.add_argument("--rate-limit-delay", type=float, help="Delay between API calls in seconds")
    parser.add_argument("--log-level", type=str, choices=["DEBUG", "INFO", "WARNING", "ERROR"], default="INFO")

    args = parser.parse_args()

    logging.basicConfig(
        level=getattr(logging, args.log_level),
        format="%(asctime)s - %(levelname)s - %(message)s"
    )

    config = APIConfig()
    if args.api_key:
        config.API_KEY = args.api_key
    if args.azure_endpoint:
        config.AZURE_ENDPOINT = args.azure_endpoint
    if args.api_version:
        config.API_VERSION = args.api_version

    if args.model:
        config.MODEL = args.model
    if args.max_retries is not None:
        config.MAX_RETRIES = args.max_retries
    if args.rate_limit_delay is not None:
        config.RATE_LIMIT_DELAY = args.rate_limit_delay

    if not config.API_KEY:
        logging.error("Please provide --api-key or set API_KEY in the script")
        return 1

    if not config.AZURE_ENDPOINT:
        logging.error("Please provide --azure-endpoint or set AZURE_ENDPOINT in the script")
        return 1

    input_path = Path(args.input)

    # Parse models filter
    models_list = None
    if args.models:
        models_list = [m.strip() for m in args.models.split(",") if m.strip()]

    # -------- Mode A: single JSONL file (original behavior preserved) --------
    if input_path.is_file():
        output_excel = Path(args.output)
        output_excel.parent.mkdir(parents=True, exist_ok=True)

        processor = JudgeProcessor(
            jsonl_path=input_path,
            output_excel=output_excel,
            config=config,
            skip_processed=not args.no_skip_processed,
            task_id_prefix_mode=args.task_id_prefix
        )
        processor.process_all()
        return 0

    # -------- Mode B: root dir => per-model Excel --------
    if not input_path.is_dir():
        logging.error(f"Input path not found: {input_path}")
        return 1

    model_dirs = iter_model_dirs(input_path, models_list)
    if not model_dirs:
        logging.error("No model directories found to process.")
        return 1

    logging.info(f"Root-dir mode: found {len(model_dirs)} model dir(s).")

    for model_dir in model_dirs:
        jsonl_path = model_dir / args.jsonl_name
        if not jsonl_path.exists():
            logging.warning(f"Skip model '{model_dir.name}': missing {args.jsonl_name}")
            continue

        output_excel = resolve_model_output_excel(model_dir, args.output)
        output_excel.parent.mkdir(parents=True, exist_ok=True)

        logging.info(f"\n{'='*70}")
        logging.info(f"MODEL: {model_dir.name}")
        logging.info(f"JSONL:  {jsonl_path}")
        logging.info(f"EXCEL: {output_excel}")
        logging.info(f"{'='*70}")

        processor = JudgeProcessor(
            jsonl_path=jsonl_path,
            output_excel=output_excel,
            config=config,
            skip_processed=not args.no_skip_processed,
            task_id_prefix_mode=args.task_id_prefix
        )
        processor.process_all()

    logging.info("All models finished.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
