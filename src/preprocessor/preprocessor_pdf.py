#!/usr/bin/env python3
"""
PDF file preprocessor.

Converts each PDF page to PNG screenshots.
"""

import logging
from pathlib import Path
from typing import Dict, List, Any

from preprocessor_base import BasePreprocessor

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None


logger = logging.getLogger(__name__)


class PDFPreprocessor(BasePreprocessor):
    """Preprocessor for PDF files."""

    PDF_EXTENSIONS = {".pdf"}

    @staticmethod
    def _is_source_file(file_path: Path, metadata: Dict[str, Any], kind: str) -> bool:
        """
        Determine whether this file belongs to metadata.source_files.
        Uses both explicit kind and metadata lookup by basename for robustness.
        """
        if kind == "source":
            return True

        source_files = metadata.get("source_files", [])
        if not isinstance(source_files, list):
            return False

        file_name = file_path.name
        source_names = {Path(str(p)).name for p in source_files if p}
        return file_name in source_names

    def can_process(self, file_path: Path) -> bool:
        """Check if file is a PDF."""
        return file_path.suffix.lower() in self.PDF_EXTENSIONS

    def process(
        self,
        file_path: Path,
        output_dir: Path,
        metadata: Dict[str, Any],
        model_name: str,
        id_str: str,
        kind: str,
    ) -> List[Dict[str, Any]]:
        """Convert each PDF page to PNG screenshot."""

        # Source-side PDF handling is controlled by config toggle.
        if (
            self._is_source_file(file_path, metadata, kind)
            and not getattr(self.config, "enable_source_non_excel_preprocess", False)
        ):
            logger.info(f"Skipping source PDF preprocessing: {file_path.name}")
            return []

        if fitz is None:
            logger.warning(f"PyMuPDF not installed, cannot process PDF: {file_path}")
            self.special_cases.append(
                f"[NO_PYMUPDF]\tmodel={model_name}\tid={id_str}\tkind={kind}\tfile={file_path.name}"
            )
            return []

        # In most datasets, output_dir is task_dir / "preprocessed"
        # So task_dir is the parent of output_dir.
        task_dir = output_dir.parent

        try:
            doc = fitz.open(file_path)
            page_count = doc.page_count

            if page_count == 0:
                logger.warning(f"PDF file has no pages: {file_path}")
                self.special_cases.append(
                    f"[EMPTY_PDF]\tmodel={model_name}\tid={id_str}\tkind={kind}\tfile={file_path.name}"
                )
                return []

            # Create directory with the same stem as the input PDF under output_dir
            stem = file_path.stem
            file_output_dir = output_dir / stem
            file_output_dir.mkdir(parents=True, exist_ok=True)

            img_paths: List[str] = []
            for i in range(page_count):
                page = doc.load_page(i)
                pix = page.get_pixmap()
                out_name = f"page{i+1}.png"
                out_path = file_output_dir / out_name
                pix.save(str(out_path))

                # Store relative path FROM task_dir (not from output_dir)
                # Example: "preprocessed/<pdf_stem>/page1.png"
                try:
                    rel_path = out_path.relative_to(task_dir)
                except ValueError:
                    # Fallback if task_dir is not an ancestor (should not happen in normal layout)
                    rel_path = out_path.relative_to(output_dir)

                # Normalize to forward slashes for JSON portability across OSes
                img_paths.append(rel_path.as_posix())

                logger.info(
                    f"[OK] {model_name} / {id_str}: {kind} PDF {file_path.name} "
                    f"page {i+1} -> PNG {out_path}"
                )

            doc.close()

            return [
                {
                    "type": "img",
                    "value": img_paths,
                    "description": self.config.pdf_config["description"],
                    "source_file": file_path.name,
                    "kind": kind,
                }
            ]

        except Exception as e:
            logger.warning(f"PDF conversion failed: {file_path}, error: {e}")
            self.special_cases.append(
                f"[PDF_CONVERT_FAIL]\tmodel={model_name}\tid={id_str}\tkind={kind}\t"
                f"file={file_path.name}\terror={e}"
            )
            return []
