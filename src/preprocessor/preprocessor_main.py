#!/usr/bin/env python3
"""
Main preprocessor entry point.

Uses chain of responsibility pattern to process files based on their type.
Scans each task directory and applies appropriate preprocessors.
"""

import json
import logging
import argparse
import sys
from pathlib import Path
from typing import Dict, List, Optional, Any

# Ensure the preprocessor package directory is on sys.path so that sibling
# module imports work regardless of how this script is invoked (direct script
# execution, subprocess from pipeline, etc.).
_THIS_DIR = Path(__file__).resolve().parent
if str(_THIS_DIR) not in sys.path:
    sys.path.insert(0, str(_THIS_DIR))

from preprocessor_base import PreprocessorConfig
from preprocessor_pdf import PDFPreprocessor
from preprocessor_markdown import MarkdownPreprocessor
from preprocessor_word import WordPreprocessor
from preprocessor_excel import ExcelPreprocessor
from preprocessor_image import ImagePreprocessor

try:
    import xlwings as xw
except ImportError:
    xw = None


# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class PreprocessorManager:
    """Main manager for preprocessing files using chain of responsibility."""

    SOURCE_NON_EXCEL_EXTENSIONS = (
        {".pdf", ".md", ".markdown", ".docx"}
        | ImagePreprocessor.IMAGE_EXTENSIONS
    )
    
    def __init__(self, root_dir: str, config: Optional[PreprocessorConfig] = None):
        """
        Initialize the preprocessor manager.
        
        Args:
            root_dir: Root directory containing model subdirectories
            config: Configuration object (uses default if not provided)
        """
        self.root_dir = Path(root_dir)
        self.config = config or PreprocessorConfig()
        
        # Build library index for resolving file references
        self.library_index: Dict[str, Path] = {}
        self.build_library_index()
        
        # Track special cases from all preprocessors
        self.all_special_cases: List[str] = []
        
        # Initialize Excel app if available
        self.excel_app = None
        if xw is not None:
            try:
                self.excel_app = xw.App(visible=False, add_book=False)
                self.excel_app.display_alerts = False
                self.excel_app.screen_updating = False
                logger.info("Excel application initialized successfully")
            except Exception as e:
                logger.warning(f"Failed to initialize Excel application: {e}")
                self.excel_app = None
        
        # Build the chain of responsibility
        self.chain_head = self.build_preprocessor_chain()
    
    def __del__(self):
        """Clean up Excel application."""
        if self.excel_app is not None:
            try:
                self.excel_app.quit()
                logger.info("Excel application closed")
            except Exception as e:
                logger.warning(f"Error closing Excel application: {e}")
    
    def build_library_index(self):
        """Build an index of all files for reference resolution."""
        logger.info("Building library index...")
        
        for file_path in self.root_dir.rglob("*"):
            if file_path.is_file():
                filename = file_path.name
                if filename not in self.library_index:
                    self.library_index[filename] = file_path
        
        logger.info(f"Indexed {len(self.library_index)} files")
    
    def build_preprocessor_chain(self):
        """
        Build the chain of responsibility for preprocessors.
        
        Returns:
            Head of the preprocessor chain
        """
        # Create preprocessors
        pdf_proc = PDFPreprocessor(self.config)
        markdown_proc = MarkdownPreprocessor(self.config, self.library_index)
        word_proc = WordPreprocessor(self.config)
        excel_proc = ExcelPreprocessor(self.config, self.excel_app)
        image_proc = ImagePreprocessor(self.config)
        
        # Build chain: PDF -> Markdown -> Word -> Excel -> Image
        pdf_proc.set_next(markdown_proc)
        markdown_proc.set_next(word_proc)
        word_proc.set_next(excel_proc)
        excel_proc.set_next(image_proc)
        
        return pdf_proc

    @staticmethod
    def _extract_basename_list(field_data: Any) -> List[str]:
        """
        Normalize metadata file field to a list of basenames.

        Supports:
        - source_files: ["a.pdf", "dir/b.docx"]
        - reference_outputs / outputs: {"files": [...]}
        """
        files: List[str] = []
        if isinstance(field_data, dict):
            raw = field_data.get("files", [])
            files = raw if isinstance(raw, list) else []
        elif isinstance(field_data, list):
            files = field_data
        else:
            return []

        result: List[str] = []
        for f in files:
            if not f:
                continue
            result.append(Path(str(f)).name)
        return result

    def _build_file_kind_map(self, metadata: Dict[str, Any]) -> Dict[str, str]:
        """
        Build filename -> kind map from metadata.
        Priority: source > reference > output
        """
        kind_map: Dict[str, str] = {}

        for name in self._extract_basename_list(metadata.get("outputs", {})):
            kind_map[name] = "output"
        for name in self._extract_basename_list(metadata.get("reference_outputs", {})):
            kind_map[name] = "reference"
        for name in self._extract_basename_list(metadata.get("source_files", [])):
            kind_map[name] = "source"

        return kind_map

    @staticmethod
    def _is_source_file_in_metadata(file_path: Path, metadata: Dict[str, Any]) -> bool:
        source_files = metadata.get("source_files", [])
        if not isinstance(source_files, list):
            return False
        source_names = {Path(str(p)).name for p in source_files if p}
        return file_path.name in source_names

    def _should_skip_source_non_excel(
        self,
        file_path: Path,
        metadata: Dict[str, Any],
        kind: str,
    ) -> bool:
        if getattr(self.config, "enable_source_non_excel_preprocess", False):
            return False

        ext = file_path.suffix.lower()
        if ext not in self.SOURCE_NON_EXCEL_EXTENSIONS:
            return False

        if kind == "source":
            return True

        return self._is_source_file_in_metadata(file_path, metadata)
    
    def process_task_directory(self, task_dir: Path, model_name: str, id_str: str):
        """
        Process all files in a task directory using the preprocessor chain.
        
        Args:
            task_dir: Path to the task directory
            model_name: Name of the model
            id_str: Task ID
        """
        logger.info(f"Processing task: {model_name} / {id_str}")
        
        # Find metadata.json
        metadata_path = task_dir / "metadata.json"
        if not metadata_path.exists():
            logger.warning(f"No metadata.json found in {task_dir}")
            return
        
        # Load metadata
        try:
            with open(metadata_path, 'r', encoding='utf-8') as f:
                metadata = json.load(f)
        except Exception as e:
            logger.error(f"Failed to load metadata.json: {e}")
            return
        
        # Initialize preprocess_info if not exists
        if "preprocess_info" not in metadata:
            metadata["preprocess_info"] = []
        
        # Create preprocessed output directory
        preprocessed_dir = task_dir / "preprocessed"
        preprocessed_dir.mkdir(exist_ok=True)
        
        # First, scan for image files in metadata (special handling)
        image_entries = ImagePreprocessor.scan_and_process_images(
            metadata,
            task_dir,
            self.config
        )
        if image_entries:
            metadata["preprocess_info"].extend(image_entries)

        file_kind_map = self._build_file_kind_map(metadata)
        
        # Process all files in the directory using the chain
        for file_path in task_dir.iterdir():
            if file_path.is_file() and file_path.name != "metadata.json":
                self.process_file_with_chain(
                    file_path,
                    preprocessed_dir,
                    metadata,
                    model_name,
                    id_str,
                    kind_hint=file_kind_map.get(file_path.name)
                )
        
        # Save updated metadata
        try:
            with open(metadata_path, 'w', encoding='utf-8') as f:
                json.dump(metadata, f, indent=2, ensure_ascii=False)
            logger.info(f"Updated metadata.json for {model_name} / {id_str}")
        except Exception as e:
            logger.error(f"Failed to save metadata.json: {e}")
    
    def process_file_with_chain(
        self,
        file_path: Path,
        output_dir: Path,
        metadata: Dict,
        model_name: str,
        id_str: str,
        kind_hint: Optional[str] = None,
    ):
        """
        Process a single file using the preprocessor chain.
        
        Args:
            file_path: Path to the file
            output_dir: Output directory for preprocessed files
            metadata: Metadata dictionary
            model_name: Model name
            id_str: Task ID
        """
        # Determine file kind: metadata map first, then filename fallback.
        if kind_hint in {"source", "reference", "output"}:
            kind = kind_hint
        else:
            filename_lower = file_path.name.lower()
            if '_src_' in filename_lower or 'source' in filename_lower:
                kind = "source"
            elif '_ref_' in filename_lower or 'reference' in filename_lower:
                kind = "reference"
            else:
                kind = "output"

        if self._should_skip_source_non_excel(file_path, metadata, kind):
            logger.info(f"Skipping source non-Excel preprocessing: {file_path.name}")
            return
        
        # Try to process with the chain
        logger.info(f"Attempting to process {file_path.name} ({kind})")
        
        preprocess_entries = self.chain_head.handle(
            file_path,
            output_dir,
            metadata,
            model_name,
            id_str,
            kind
        )
        
        if preprocess_entries is not None and len(preprocess_entries) > 0:
            # Add file information to each entry
            for entry in preprocess_entries:
                if "source_file" not in entry:
                    entry["source_file"] = file_path.name
                if "kind" not in entry:
                    entry["kind"] = kind
            
            metadata["preprocess_info"].extend(preprocess_entries)
            
            # Collect special cases from chain
            self.collect_special_cases_from_chain(self.chain_head)
        else:
            logger.info(f"No preprocessor handled {file_path.name}, skipping")
    
    def collect_special_cases_from_chain(self, handler):
        """Recursively collect special cases from all handlers in chain."""
        if hasattr(handler, 'special_cases') and handler.special_cases:
            self.all_special_cases.extend(handler.special_cases)
            handler.special_cases = []  # Clear after collecting
        
        if hasattr(handler, 'next_handler') and handler.next_handler is not None:
            self.collect_special_cases_from_chain(handler.next_handler)
    
    def preprocess_all(self, specific_models: Optional[List[str]] = None):
        """
        Preprocess all tasks in the root directory.
        
        Args:
            specific_models: Optional list of specific model directory names to process.
                           If None, all model directories will be processed.
        """
        logger.info("Starting preprocessing...")
        logger.info(f"Root directory: {self.root_dir}")
        
        # Find all model directories
        all_model_dirs = [d for d in self.root_dir.iterdir() if d.is_dir()]
        
        if not all_model_dirs:
            logger.warning(f"No model directories found in {self.root_dir}")
            return
        
        # Filter model directories if specific models are specified
        if specific_models:
            model_dirs = [d for d in all_model_dirs if d.name in specific_models]
            
            # Check for invalid model names
            found_names = {d.name for d in model_dirs}
            invalid_models = set(specific_models) - found_names
            if invalid_models:
                logger.warning(f"Specified models not found: {', '.join(invalid_models)}")
            
            if not model_dirs:
                logger.error(f"None of the specified models found in {self.root_dir}")
                logger.info(f"Available models: {', '.join(d.name for d in all_model_dirs)}")
                return
            
            logger.info(f"Processing {len(model_dirs)} specified model(s): {', '.join(d.name for d in model_dirs)}")
        else:
            model_dirs = all_model_dirs
            logger.info(f"Found {len(model_dirs)} model(s) (processing all)")
        
        total_tasks = 0
        processed_tasks = 0
        
        # Process each model
        for model_dir in model_dirs:
            model_name = model_dir.name
            logger.info(f"\nProcessing model: {model_name}")
            
            # Find all task directories
            task_dirs = [d for d in model_dir.iterdir() if d.is_dir()]
            
            for task_dir in task_dirs:
                id_str = task_dir.name
                total_tasks += 1
                
                try:
                    self.process_task_directory(task_dir, model_name, id_str)
                    processed_tasks += 1
                except Exception as e:
                    logger.error(
                        f"Failed to process task {model_name}/{id_str}: {e}",
                        exc_info=True
                    )
        
        logger.info(f"\nPreprocessing complete!")
        logger.info(f"Processed {processed_tasks}/{total_tasks} tasks")
        
        # Save special cases log
        if self.all_special_cases:
            special_cases_path = self.root_dir / "preprocessing_special_cases.log"
            with open(special_cases_path, 'w', encoding='utf-8') as f:
                f.write("\n".join(self.all_special_cases))
            logger.info(f"Special cases logged to: {special_cases_path}")
            logger.info(f"Total special cases: {len(self.all_special_cases)}")


def main():
    """Main entry point for the preprocessor."""
    parser = argparse.ArgumentParser(
        description="Preprocess files and add preprocess_info to metadata.json"
    )
    parser.add_argument(
        "--root-dir",
        required=True,
        help="Root directory containing model subdirectories"
    )
    parser.add_argument(
        "--models",
        nargs="+",
        default=None,
        help="Specific model subdirectories to process (space-separated). If not specified, all models will be processed. Example: --models model1 model2"
    )
    parser.add_argument(
        "--enable-source-non-excel-preprocess",
        action="store_true",
        help=(
            "Allow source_files img/md/docx/pdf to be preprocessed into preprocess_info. "
            "Default is disabled."
        ),
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="Logging level (default: INFO)"
    )
    
    args = parser.parse_args()
    
    # Set log level
    logging.getLogger().setLevel(getattr(logging, args.log_level))
    
    config = PreprocessorConfig()
    config.enable_source_non_excel_preprocess = bool(args.enable_source_non_excel_preprocess)

    # Create and run preprocessor
    manager = PreprocessorManager(root_dir=args.root_dir, config=config)
    manager.preprocess_all(specific_models=args.models)


if __name__ == "__main__":
    main()
