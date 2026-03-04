#!/usr/bin/env python3
"""
Word document preprocessor.

Extracts text and page screenshots from Word documents.
"""

import logging
from pathlib import Path
from typing import Dict, List, Any

from preprocessor_base import BasePreprocessor

try:
    from docx import Document
except ImportError:
    Document = None

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    import win32com.client as win32
except ImportError:
    win32 = None


logger = logging.getLogger(__name__)


class WordPreprocessor(BasePreprocessor):
    """Preprocessor for Word documents."""

    WORD_EXTENSIONS = {".docx"}
    WD_EXPORT_FORMAT_PDF = 17

    def can_process(self, file_path: Path) -> bool:
        """Check if file is a Word document."""
        return file_path.suffix.lower() in self.WORD_EXTENSIONS

    def process(
        self,
        file_path: Path,
        output_dir: Path,
        metadata: Dict[str, Any],
        model_name: str,
        id_str: str,
        kind: str,
    ) -> List[Dict[str, Any]]:
        """Split Word document into text and page screenshot sequence."""

        if Document is None:
            logger.warning(f"python-docx not installed, cannot process docx: {file_path}")
            self.special_cases.append(
                f"[NO_PYTHON_DOCX]\tmodel={model_name}\tid={id_str}\tkind={kind}\tfile={file_path.name}"
            )
            return []

        if win32 is None:
            logger.warning(f"pywin32 not installed, cannot export docx to PDF: {file_path}")
            self.special_cases.append(
                f"[NO_WIN32COM]\tmodel={model_name}\tid={id_str}\tkind={kind}\tfile={file_path.name}"
            )
            return []

        if fitz is None:
            logger.warning(f"PyMuPDF not installed, cannot render docx pages: {file_path}")
            self.special_cases.append(
                f"[NO_PYMUPDF]\tmodel={model_name}\tid={id_str}\tkind={kind}\tfile={file_path.name}"
            )
            return []

        # In most datasets, output_dir is task_dir / "preprocessed"
        # So task_dir is the parent of output_dir.
        task_dir = output_dir.parent

        # Create directory with same name as file under output_dir
        stem = file_path.stem
        file_output_dir = output_dir / stem
        file_output_dir.mkdir(parents=True, exist_ok=True)

        try:
            text_doc = Document(str(file_path))
        except Exception as e:
            logger.warning(f"Failed to open Word document: {file_path}, error: {e}")
            self.special_cases.append(
                f"[DOCX_OPEN_FAIL]\tmodel={model_name}\tid={id_str}\tkind={kind}\tfile={file_path.name}\terror={e}"
            )
            return []

        preprocess_info: List[Dict[str, Any]] = []
        text_idx = 0

        def _rel_to_task(p: Path) -> str:
            # Store relative path FROM task_dir (not from output_dir)
            # Example: "preprocessed/<doc_stem>/page1.png"
            try:
                rel = p.relative_to(task_dir)
            except ValueError:
                rel = p.relative_to(output_dir)
            return rel.as_posix()

        # 1. Process text: combine all paragraphs
        texts: List[str] = []
        for para in text_doc.paragraphs:
            if para.text:
                texts.append(para.text)

        if texts:
            text_idx += 1
            text_content = "\n".join(texts)
            text_path = file_output_dir / f"text{text_idx}.txt"
            with text_path.open("w", encoding="utf-8") as f:
                f.write(text_content)

            preprocess_info.append(
                {
                    "type": "text",
                    "value": text_content,
                    "description": self.config.word_config["text_description"],
                    "source_file": file_path.name,
                    "kind": kind,
                }
            )

            logger.info(f"[OK] {model_name} / {id_str}: {kind} docx text -> {text_path}")

        # 2. Export DOCX to PDF with Word, then render each PDF page as PNG screenshot.
        pdf_path = file_output_dir / f"{stem}.pdf"
        pdf_path_abs = str(pdf_path.resolve())
        word_app = None
        com_doc = None

        try:
            word_app = win32.Dispatch("Word.Application")
            word_app.Visible = False
            if hasattr(word_app, "DisplayAlerts"):
                word_app.DisplayAlerts = 0

            com_doc = word_app.Documents.Open(str(file_path.resolve()))
            com_doc.ExportAsFixedFormat(pdf_path_abs, self.WD_EXPORT_FORMAT_PDF)
            logger.info(f"[OK] {model_name} / {id_str}: {kind} docx exported -> {pdf_path_abs}")
        except Exception as e:
            logger.warning(f"Failed to export Word document to PDF: {file_path}, error: {e}")
            self.special_cases.append(
                f"[DOCX_EXPORT_PDF_FAIL]\tmodel={model_name}\tid={id_str}\tkind={kind}\tfile={file_path.name}\terror={e}"
            )
            return preprocess_info
        finally:
            if com_doc is not None:
                try:
                    com_doc.Close(False)
                except Exception:
                    pass
            if word_app is not None:
                try:
                    word_app.Quit()
                except Exception:
                    pass

        img_paths: List[str] = []
        pdf_doc = None
        try:
            pdf_doc = fitz.open(pdf_path_abs)
            if pdf_doc.page_count == 0:
                logger.warning(f"Exported PDF has no pages: {pdf_path}")
                self.special_cases.append(
                    f"[EMPTY_DOCX_PDF]\tmodel={model_name}\tid={id_str}\tkind={kind}\tfile={file_path.name}"
                )
                return preprocess_info

            for i in range(pdf_doc.page_count):
                try:
                    page = pdf_doc.load_page(i)
                    pix = page.get_pixmap()
                    page_path = file_output_dir / f"page{i+1}.png"
                    pix.save(str(page_path))

                    rel_path = _rel_to_task(page_path)
                    img_paths.append(rel_path)

                    logger.info(
                        f"[OK] {model_name} / {id_str}: {kind} docx page {i+1} -> {page_path}"
                    )
                except Exception as e:
                    logger.warning(f"Failed to render docx page {i+1}: {e}")
                    self.special_cases.append(
                        f"[DOCX_PAGE_RENDER_FAIL]\tmodel={model_name}\tid={id_str}\tkind={kind}\tfile={file_path.name}\tpage={i+1}\terror={e}"
                    )
        except Exception as e:
            logger.warning(f"Failed to render exported PDF for {file_path}: {e}")
            self.special_cases.append(
                f"[DOCX_PDF_RENDER_FAIL]\tmodel={model_name}\tid={id_str}\tkind={kind}\tfile={file_path.name}\terror={e}"
            )
            return preprocess_info
        finally:
            if pdf_doc is not None:
                try:
                    pdf_doc.close()
                except Exception:
                    pass
            if pdf_path.exists():
                try:
                    pdf_path.unlink()
                except Exception as cleanup_error:
                    logger.debug(f"Failed to delete temporary PDF {pdf_path}: {cleanup_error}")

        if img_paths:
            preprocess_info.append(
                {
                    "type": "img",
                    "value": img_paths,
                    "description": self.config.word_config["img_description"],
                    "source_file": file_path.name,
                    "kind": kind,
                }
            )

        return preprocess_info
