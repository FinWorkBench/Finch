#!/usr/bin/env python3
"""
Markdown file preprocessor.

Splits markdown content into text and image sequences.
"""

import re
import shutil
import logging
from pathlib import Path
from typing import Dict, List, Any, Optional

from preprocessor_base import BasePreprocessor


logger = logging.getLogger(__name__)


class MarkdownPreprocessor(BasePreprocessor):
    """Preprocessor for Markdown files."""

    MARKDOWN_EXTENSIONS = {".md", ".markdown"}

    def __init__(self, config, library_index: Dict[str, Path]):
        """
        Initialize Markdown preprocessor.

        Args:
            config: Configuration object
            library_index: Dictionary mapping filenames to their paths
        """
        super().__init__(config)
        self.library_index = library_index

    def can_process(self, file_path: Path) -> bool:
        """Check if file is a Markdown file."""
        return file_path.suffix.lower() in self.MARKDOWN_EXTENSIONS

    def resolve_library_file(self, img_path_raw: str) -> Optional[Path]:
        """Resolve image path using library index."""
        filename = Path(img_path_raw).name
        return self.library_index.get(filename)

    def process(
        self,
        file_path: Path,
        output_dir: Path,
        metadata: Dict[str, Any],
        model_name: str,
        id_str: str,
        kind: str,
    ) -> List[Dict[str, Any]]:
        """Split markdown into text and image sequence."""

        # In most datasets, output_dir is task_dir / "preprocessed"
        # So task_dir is the parent of output_dir.
        task_dir = output_dir.parent

        # Create directory with the same stem as the input file under output_dir
        stem = file_path.stem
        file_output_dir = output_dir / stem
        file_output_dir.mkdir(parents=True, exist_ok=True)

        try:
            content = file_path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            content = file_path.read_text(encoding="utf-8", errors="ignore")

        img_pattern = re.compile(r"!\[[^\]]*\]\(([^)]+)\)")
        pos = 0
        text_idx = 0
        pic_idx = 0

        preprocess_info: List[Dict[str, Any]] = []

        def _rel_to_task(p: Path) -> str:
            # Store relative path FROM task_dir (not from output_dir)
            # Example: "preprocessed/<stem>/pic1.png"
            try:
                rel = p.relative_to(task_dir)
            except ValueError:
                rel = p.relative_to(output_dir)
            return rel.as_posix()

        for m in img_pattern.finditer(content):
            # 1. Process text before image
            before = content[pos : m.start()]
            if before:
                text_idx += 1
                text_path = file_output_dir / f"text{text_idx}.txt"
                with text_path.open("w", encoding="utf-8") as f:
                    f.write(before)

                preprocess_info.append(
                    {
                        "type": "text",
                        "value": before,
                        "description": self.config.markdown_config["text_description"],
                        "source_file": file_path.name,
                        "kind": kind,
                    }
                )

                logger.info(f"[OK] {model_name} / {id_str}: {kind} markdown text -> {text_path}")

            # 2. Process image
            img_path_raw = m.group(1).strip()
            img_fp = self.resolve_library_file(img_path_raw)

            if img_fp:
                pic_idx += 1
                ext = img_fp.suffix or ".png"
                pic_path = file_output_dir / f"pic{pic_idx}{ext}"
                shutil.copy2(img_fp, pic_path)

                rel_path = _rel_to_task(pic_path)
                preprocess_info.append(
                    {
                        "type": "img",
                        "value": [rel_path],
                        "description": self.config.markdown_config["img_description"],
                        "source_file": file_path.name,
                        "kind": kind,
                    }
                )

                logger.info(
                    f"[OK] {model_name} / {id_str}: {kind} markdown image {img_fp} -> {pic_path}"
                )
            else:
                # Image not found, treat the markdown image snippet as text
                text_idx += 1
                text_path = file_output_dir / f"text{text_idx}.txt"
                snippet = content[m.start() : m.end()]
                with text_path.open("w", encoding="utf-8") as f:
                    f.write(snippet)

                preprocess_info.append(
                    {
                        "type": "text",
                        "value": snippet,
                        "description": self.config.markdown_config["text_description"],
                        "source_file": file_path.name,
                        "kind": kind,
                    }
                )

                self.special_cases.append(
                    f"[MISSING_MD_IMAGE]\tmodel={model_name}\tid={id_str}\tkind={kind}\tref={img_path_raw}"
                )

            pos = m.end()

        # 3. Process remaining tail text
        tail = content[pos:]
        if tail:
            text_idx += 1
            text_path = file_output_dir / f"text{text_idx}.txt"
            with text_path.open("w", encoding="utf-8") as f:
                f.write(tail)

            preprocess_info.append(
                {
                    "type": "text",
                    "value": tail,
                    "description": self.config.markdown_config["text_description"],
                    "source_file": file_path.name,
                    "kind": kind,
                }
            )

            logger.info(f"[OK] {model_name} / {id_str}: {kind} markdown tail text -> {text_path}")

        return preprocess_info
