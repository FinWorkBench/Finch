"""
Main Content Builder

Coordinates the content building process for evaluation tasks.
Scans model directories, identifies task types (Excel or general),
and delegates to appropriate specialized builders.

python -m src2.build_prompt.content_builder.content_builder data/eval_dataset_new_test
python -m src2.build_prompt.content_builder.content_builder data/eval_dataset_new_test --models opus_4.5_output
python -m src2.build_prompt.content_builder.content_builder data/eval_dataset_new_test --models opus_4.5_output gpt4o_output
python -m ... data/eval_dataset_new_test --models all
"""



import json
import logging
from pathlib import Path
from typing import List, Dict, Any, Optional, Set

from .config import OUTPUT_JSONL_NAME, MAX_IMAGES, MAX_TEXT_CHARS
from .cache_manager import CacheManager

from .utils import get_first_excel_file

from .excel_content_builder import ExcelContentBuilder, ExcelSideResult
from .general_content_builder import GeneralContentBuilder
from .prompts import JUDGE_PROMPT_GENERAL, JUDGE_PROMPT_EXCEL
from .token_counter import truncate_content_parts


logger = logging.getLogger(__name__)




class ContentBuilder:
    """
    Main content builder that coordinates the entire evaluation process.
    
    Responsibilities:
    - Scan root directory for model subdirectories
    - For each model, scan for task subdirectories
    - For each task, determine evaluation type (Excel vs general)
    - Delegate to ExcelContentBuilder or GeneralContentBuilder
    - Write content parts to JSONL file incrementally
    - Apply token limit truncation when necessary
    
    Directory structure:
        root_dir/
            model1/
                task1/
                    preprocessed/
                    metadata.json
                    ...
                task2/
                    ...
            model2/
                ...
    """
    
    def __init__(self, root_dir: str, models: Optional[List[str]] = None):
        """
        Args:
            root_dir: Root directory containing model subdirectories
            models: Optional list of model directory names to process.
                    If None or empty -> process all model dirs.
        """
        self.root_dir = Path(root_dir)
        self.models = [m.strip() for m in (models or []) if m and m.strip()]
    
    def build_all(self) -> None:
        """
        Build content parts for all models and tasks in the root directory.
        
        Iterates through each model directory and processes all tasks within.
        Results are written to individual JSONL files per model.
        """
        if not self.root_dir.exists():
            logger.error(f"Root directory does not exist: {self.root_dir}")
            return
        
        logger.info(f"Starting content building for: {self.root_dir}")
        
        # Iterate through model directories
        all_model_dirs = sorted([d for d in self.root_dir.iterdir() if d.is_dir()])

        # Filter model dirs if user specified
        model_dirs = all_model_dirs
        if self.models and "all" not in {m.lower() for m in self.models}:
            name_set = {m for m in self.models}
            model_dirs = [d for d in all_model_dirs if d.name in name_set]

            missing = sorted(name_set - {d.name for d in all_model_dirs})
            if missing:
                logger.warning(f"Requested model directories not found under {self.root_dir}: {missing}")

        if not model_dirs:
            logger.warning("No model directories to process after filtering.")
            return

        for model_dir in model_dirs:
            logger.info(f"\n{'='*60}")
            logger.info(f"Processing model: {model_dir.name}")
            logger.info(f"{'='*60}")
            self.build_model(model_dir)
        
        logger.info(f"\n{'='*60}")
        logger.info("Content building completed for all models")
        logger.info(f"{'='*60}")
    
    def build_model(self, model_dir: Path) -> None:
        """
        Build content parts for all tasks within a single model directory.

        Always writes one JSONL line per task.
        If a task fails, writes an empty content_parts list instead of skipping.
        """
        output_file = model_dir / OUTPUT_JSONL_NAME
        task_dirs = [d for d in model_dir.iterdir() if d.is_dir()]

        logger.info(f"Found {len(task_dirs)} task directories")

        with output_file.open("w", encoding="utf-8") as f:
            for task_dir in sorted(task_dirs):
                content_parts: List[Dict[str, Any]] = []
                err_msg: Optional[str] = None

                try:
                    logger.info(f"\n--- Processing task: {task_dir.name} ---")
                    content_parts = self.build_task(task_dir)
                    logger.info(f"✓ Task {task_dir.name} completed")
                except Exception as e:
                    # Do not skip writing output. Emit empty content_parts for this task.
                    err_msg = f"{type(e).__name__}: {e}"
                    logger.error(f"✗ Failed to process task {task_dir.name}: {err_msg}", exc_info=True)
                    content_parts = []

                # Always write to JSONL immediately (success or failure)
                record: Dict[str, Any] = {
                    "task_id": task_dir.name,
                    "content_parts": content_parts,
                }

                # Optional: keep error for debugging; remove if you must keep schema strict
                if err_msg is not None:
                    record["error"] = err_msg

                json_line = json.dumps(record, ensure_ascii=False)
                f.write(json_line + "\n")
                f.flush()

                if err_msg is None:
                    logger.info(f"✓ Task {task_dir.name} written to JSONL")
                else:
                    logger.info(f"✓ Task {task_dir.name} failed but empty content_parts written to JSONL")

        logger.info(f"\nModel output written to: {output_file}")

    
    def build_task(self, task_dir: Path) -> List[Dict[str, Any]]:
        # Load metadata.json (if missing, proceed with empty metadata)
        metadata_file = task_dir / "metadata.json"
        if not metadata_file.exists():
            logger.warning(f"No metadata.json found in {task_dir.name}")
            metadata: Dict[str, Any] = {}
        else:
            with metadata_file.open("r", encoding="utf-8") as f:
                metadata = json.load(f)

        # Resolve file lists from metadata
        reference_outputs = self._get_files_from_metadata(task_dir, metadata, "reference_outputs")
        outputs = self._get_files_from_metadata(task_dir, metadata, "outputs")
        source_files = self._get_files_from_metadata(task_dir, metadata, "source_files")

        # Detect Excel files per side
        ref_excel = get_first_excel_file(reference_outputs)   # Path | None
        out_excel = get_first_excel_file(outputs)             # Path | None
        source_excel = get_first_excel_file(source_files)     # Path | None

        ref_is_excel = bool(ref_excel and Path(ref_excel).exists())
        out_is_excel = bool(out_excel and Path(out_excel).exists())
        src_is_excel = bool(source_excel and Path(source_excel).exists())

        # Base prompt strategy:
        # - If BOTH sides are Excel, use the Excel judge prompt.
        # - Otherwise, use the general judge prompt (supports mixed-format tasks).
        base_prompt = JUDGE_PROMPT_EXCEL if (ref_is_excel and out_is_excel) else JUDGE_PROMPT_GENERAL
        content_parts: List[Dict[str, Any]] = [{"type": "text", "text": base_prompt}]

        # Add common metadata content once
        self._add_instruction(content_parts, metadata)
        self._add_metadata_outputs_text(content_parts, metadata)
        self._add_preprocess_info(content_parts, metadata, task_dir)


        # Builders
        cache = CacheManager(task_dir)
        excel_builder = ExcelContentBuilder(task_dir, cache)
        general_builder = GeneralContentBuilder(task_dir)

        ref_excel_result = ExcelSideResult.empty()
        out_excel_result = ExcelSideResult.empty()

        # Build reference side
        if ref_is_excel:
            ref_excel_result = excel_builder.build_side(
                content_parts=content_parts,
                side="reference",
                target_file=Path(ref_excel),
                source_file=Path(source_excel) if src_is_excel else None,
            )
        else:
            general_builder.build_side_files(
                content_parts=content_parts,
                side="reference",
                files=reference_outputs,
            )

        # Build output side
        if out_is_excel:
            out_excel_result = excel_builder.build_side(
                content_parts=content_parts,
                side="output",
                target_file=Path(out_excel),
                source_file=Path(source_excel) if src_is_excel else None,
            )
        else:
            general_builder.build_side_files(
                content_parts=content_parts,
                side="output",
                files=outputs,
            )

        # Append source content ONCE (snapshot + screenshots) if applicable
        if src_is_excel and (ref_is_excel or out_is_excel):
            focus_sheets: Set[str] = set()
            if ref_excel_result.used_diff:
                focus_sheets |= ref_excel_result.changed_sheets
            if out_excel_result.used_diff:
                focus_sheets |= out_excel_result.changed_sheets

            def _merge_dict_sets(
                a: Optional[Dict[str, Set[int]]],
                b: Optional[Dict[str, Set[int]]],
            ) -> Dict[str, Set[int]]:
                merged: Dict[str, Set[int]] = {k: set(v) for k, v in (a or {}).items()}
                for k, v in (b or {}).items():
                    merged.setdefault(k, set()).update(v)
                return merged

            changed_rows = _merge_dict_sets(ref_excel_result.changed_rows, out_excel_result.changed_rows)
            changed_cols = _merge_dict_sets(ref_excel_result.changed_cols, out_excel_result.changed_cols)

            excel_builder.add_source_once(
                content_parts=content_parts,
                source_file=Path(source_excel),
                focus_sheets=focus_sheets,  # empty => all sheets
                changed_rows=changed_rows,
                changed_cols=changed_cols,
            )

        # Apply global truncation (images/text limits)
        content_parts, truncation_info = truncate_content_parts(
            content_parts,
            max_images=MAX_IMAGES,
            max_text_chars=MAX_TEXT_CHARS,
        )

        if truncation_info.get("was_truncated"):
            logger.warning(
                f"Reduced content for task {task_dir.name}: "
                f"images {truncation_info.get('original_images')} -> {truncation_info.get('final_images')} "
                f"(removed={truncation_info.get('images_removed')}), "
                f"text_chars {truncation_info.get('original_text_chars')} -> {truncation_info.get('final_text_chars')} "
                f"(removed={truncation_info.get('text_chars_removed')}), "
                f"text_parts_deleted={truncation_info.get('text_parts_deleted', 0)}, "
                f"note_added={truncation_info.get('note_added')}"
            )

        return content_parts

    # Move common metadata helpers here so they are applied once for mixed tasks

    def _add_instruction(self, content_parts: List[Dict[str, Any]], metadata: Dict[str, Any]) -> None:
        instruction_en = (metadata.get("instruction_en") or "").strip()
        task_constraints = (metadata.get("task_constraints") or "").strip()
        if not instruction_en and not task_constraints:
            return

        merged: List[str] = []
        if instruction_en:
            merged.append("## Instruction\n" + instruction_en)
        if task_constraints:
            merged.append("## Task Constraints\n" + task_constraints)

        content_parts.append({"type": "text", "text": "\n\n".join(merged)})

    def _add_metadata_outputs_text(self, content_parts: List[Dict[str, Any]], metadata: Dict[str, Any]) -> None:
        from .config import Captions

        def _normalize_text(val: Any) -> str:
            if val is None:
                return ""
            if isinstance(val, str):
                return val.strip()
            if isinstance(val, list):
                return "\n".join(str(x) for x in val).strip()
            return str(val).strip()

        def _append_if_present(obj: Any, caption_tpl: str) -> None:
            if not obj:
                return

            items: List[Dict[str, Any]]
            if isinstance(obj, dict):
                items = [obj]
            elif isinstance(obj, list):
                items = [x for x in obj if isinstance(x, dict)]
            else:
                return

            for item in items:
                text = _normalize_text(item.get("text"))
                if text:
                    content_parts.append({"type": "text", "text": caption_tpl.format(text=text)})

        _append_if_present(metadata.get("reference_outputs"), Captions.METADATA_REFERENCE_OUTPUT_TEXT)
        _append_if_present(metadata.get("outputs"), Captions.METADATA_OUTPUT_TEXT)

    def _add_preprocess_info(
        self,
        content_parts: List[Dict[str, Any]],
        metadata: Dict[str, Any],
        task_dir: Path,
    ) -> None:
        from .config import Captions
        from .utils import image_to_data_url

        def _resolve_path(p: str) -> Path:
            # Resolve relative paths against task_dir
            pp = Path(p)
            if pp.is_absolute():
                return pp
            return (task_dir / pp)

        preprocess_info = metadata.get("preprocess_info", [])
        for item in preprocess_info:
            item_type = item.get("type")
            value = item.get("value")
            description = item.get("description", "")

            if item_type == "text":
                content_parts.append({
                    "type": "text",
                    "text": Captions.PREPROCESS_TEXT.format(description=description, text=value)
                })
                continue

            if item_type == "img":
                # Add description once
                if description:
                    content_parts.append({
                        "type": "text",
                        "text": Captions.PREPROCESS_IMAGE.format(description=description)
                    })

                # value is expected to be a list of image paths
                if not isinstance(value, list):
                    logger.warning(f"preprocess_info.img value is not a list: {type(value)}")
                    continue

                for img_path in value:
                    try:
                        p = _resolve_path(str(img_path))
                        if p.exists():
                            data_url = image_to_data_url(p)
                            content_parts.append({"type": "image_url", "image_url": {"url": data_url}})
                        else:
                            logger.warning(f"Preprocess image not found: {p}")
                    except Exception as e:
                        logger.warning(f"Failed to add preprocess image {img_path}: {e}")

    
    
    def _get_files_from_metadata(
        self,
        task_dir: Path,
        metadata: Dict[str, Any],
        field_name: str
    ) -> List[str]:
        """
        Extract file paths from metadata and convert to absolute paths.
        
        Supports multiple metadata structures:
        - {"files": ["file1.xlsx", "file2.pdf"], ...}
        - ["file1.xlsx", "file2.pdf"]
        
        Validates that files exist and logs warnings for missing files.
        
        Args:
            task_dir: Task directory path (base for relative paths)
            metadata: Task metadata dictionary
            field_name: Metadata field name (e.g., "reference_outputs", "outputs", "source_files")
        
        Returns:
            List of absolute file path strings for existing files
        """
        file_list = []
        
        # Get the field from metadata
        field_data = metadata.get(field_name, {})
        
        if not field_data:
            logger.debug(f"No '{field_name}' field in metadata")
            return []
        
        # Handle different metadata structures
        if isinstance(field_data, dict):
            # Structure: {"files": ["file1.xlsx", "file2.pdf"], ...}
            files = field_data.get("files", [])
        elif isinstance(field_data, list):
            # Structure: ["file1.xlsx", "file2.pdf"]
            files = field_data
        else:
            logger.warning(f"Unexpected format for '{field_name}' in metadata: {type(field_data)}")
            return []
        
        # Convert relative filenames to absolute paths
        for filename in files:
            # Try to find file in task directory
            file_path = task_dir / filename
            
            if file_path.exists():
                file_list.append(str(file_path.absolute()))
            else:
                logger.warning(f"File not found: {file_path}")
        
        logger.debug(f"Found {len(file_list)} files for '{field_name}'")
        return file_list
    
    def _get_file_list(self, directory: Path) -> List[str]:
        """
        Get list of all file paths in a directory (fallback method).
        
        Used when metadata is unavailable or incomplete. Returns all files
        in the specified directory (non-recursive).
        
        Args:
            directory: Directory path to scan
        
        Returns:
            List of absolute file path strings, empty list if directory doesn't exist
        """
        if not directory.exists():
            return []
        
        return [
            str(f.absolute())
            for f in directory.iterdir()
            if f.is_file()
        ]


# ==================== Command Line Interface ====================

def main():
    """
    Command line interface for the content builder.
    
    Usage:
        python content_builder.py <root_dir> [--log-level LEVEL]
    
    Arguments:
        root_dir: Root directory containing model subdirectories
        --log-level: Logging verbosity (DEBUG, INFO, WARNING, ERROR)
    
    Example:
        python content_builder.py ./evaluation_data --log-level INFO
    """
    import argparse
    
    parser = argparse.ArgumentParser(
        description="Build evaluation content parts for model outputs"
    )
    parser.add_argument(
        "--models",
        nargs="*",
        default=None,
        help=(
            "Model subdirectory names to process. "
            "Examples: --models opus_4.5_output  |  --models opus_4.5_output gpt4o_output  |  --models all. "
            "If omitted, all model subdirectories are processed."
        )
    )

    parser.add_argument(
        "root_dir",
        type=str,
        help="Root directory containing model subdirectories"
    )
    parser.add_argument(
        "--log-level",
        type=str,
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        default="INFO",
        help="Logging level (default: INFO)"
    )
    
    args = parser.parse_args()
    
    # Configure logging
    logging.basicConfig(
        level=getattr(logging, args.log_level),
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )
    
    # Build content
    builder = ContentBuilder(args.root_dir, models=args.models)
    builder.build_all()


if __name__ == "__main__":
    main()
