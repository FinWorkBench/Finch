"""
General Content Builder

Builds content parts for general (non-Excel) evaluation tasks.
"""

import logging
from pathlib import Path
from typing import List, Dict, Any


from .utils import image_to_data_url


logger = logging.getLogger(__name__)



class GeneralContentBuilder:
    """
    Builds content parts for general evaluation tasks.
    
    Handles non-Excel files like PDF, Word, Markdown, images, text, etc.
    """
    
    def __init__(self, task_dir: Path):
        """
        Initialize general content builder.
        
        Args:
            task_dir: Path to task directory
        """
        self.task_dir = Path(task_dir)
    
    def build_side_files(
        self,
        content_parts: List[Dict[str, Any]],
        side: str,                 # "reference" | "output"
        files: List[str],          # absolute paths
        max_chars_per_file: int = 50_000,
    ) -> None:
        """
        Attach side-specific files (non-Excel) into content_parts.

        - Images: embed as image_url
        - .docx: extract text via python-docx
        - Text-like: read as utf-8 with errors ignored (truncate)
        - Others: add a placeholder line (extend later for PDF, etc.)
        """
        side = (side or "").lower().strip()
        assert side in {"reference", "output"}

        header = "## Reference Files" if side == "reference" else "## Output Files"
        content_parts.append({"type": "text", "text": header})

        for fp in files:
            p = Path(fp)
            if not p.exists() or not p.is_file():
                continue

            ext = p.suffix.lower()
            content_parts.append({"type": "text", "text": f"### {p.name} ({ext})"})

            # Images
            if ext in {".png", ".jpg", ".jpeg", ".webp", ".bmp"}:
                try:
                    content_parts.append({"type": "image_url", "image_url": {"url": image_to_data_url(p)}})
                except Exception as e:
                    content_parts.append({"type": "text", "text": f"[image load failed] {p.name}: {e}"})
                continue

            # DOCX
            if ext == ".docx":
                try:
                    from docx import Document
                    doc = Document(str(p))
                    text = "\n".join([para.text for para in doc.paragraphs if para.text]).strip()
                    # No truncation here; let outer token_counter handle limits.
                    content_parts.append({"type": "text", "text": text or "[empty docx]"})
                except Exception as e:
                    content_parts.append({"type": "text", "text": f"[docx parse failed] {p.name}: {e}"})
                continue

            # Text-like files. Since JSX is relatively uncommon and requires additional rendering, we will temporarily extract its content and process it as plain text.
            if ext in {".txt", ".md", ".py", ".json", ".yaml", ".yml", ".csv", ".log", ".jsx"}:
                try:
                    text = p.read_text(encoding="utf-8", errors="ignore").strip()
                    # No truncation here; let outer token_counter handle limits.
                    content_parts.append({"type": "text", "text": text or "[empty text file]"})
                except Exception as e:
                    content_parts.append({"type": "text", "text": f"[text read failed] {p.name}: {e}"})
                continue

            # Unhandled formats (extend later: PDF parsing, etc.)
            content_parts.append({"type": "text", "text": f"[file attached but not parsed] path={str(p)}"})
    
