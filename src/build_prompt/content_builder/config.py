"""
Content Builder Configuration

All configurable settings for the content building process.
"""

from pathlib import Path


# ==================== Token Limits ====================

# Maximum total tokens for content parts

MAX_IMAGES = 40
MAX_TEXT_CHARS = 200000




# ==================== File Extensions ====================

EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".gif"}
TEXT_EXTENSIONS = {".txt", ".md", ".csv", ".json"}
PDF_EXTENSIONS = {".pdf"}
WORD_EXTENSIONS = {".docx", ".doc"}


# ==================== Cache Settings ====================

# Cache directory name (created in each task subdirectory)
CACHE_DIR_NAME = "_cache"

# Cache file names
CACHE_DIFF_JSON = "diff_cache.json"
CACHE_FULL_SNAPSHOT_PREFIX = "full_snapshot"
CACHE_RICH_SNAPSHOT_PREFIX = "rich_snapshot"
CACHE_SIMPLE_SNAPSHOT_PREFIX = "simple_snapshot"
CACHE_SCREENSHOTS_DIR = "screenshots"


# ==================== Excel Settings ====================

# Excel snapshot builder settings
EXCEL_SNAPSHOT_MAX_CHARS = 500000
EXCEL_SNAPSHOT_MAX_ROWS_FRONT = 10
EXCEL_SNAPSHOT_MAX_ROWS_BACK = 10
EXCEL_SNAPSHOT_MAX_COLS_SIMPLE = 5

# Excel screenshot settings
EXCEL_SCREENSHOT_MAX_ROWS = 200
EXCEL_SCREENSHOT_MAX_COLS = 50


# ==================== Output Settings ====================

# JSONL output file name in model subdirectory
OUTPUT_JSONL_NAME = "content_parts.jsonl"


# ==================== Caption Templates ====================

class Captions:
    """
    Configurable caption templates for different content types.
    All captions can be customized here.
    """
    
    METADATA_REFERENCE_OUTPUT_TEXT = "## Reference Outputs Text\n{text}"
    METADATA_OUTPUT_TEXT = "## Output Text\n{text}"

    # Excel evaluation captions
    EXCEL_INSTRUCTION = "## Task Instruction\n\n{instruction}"
    EXCEL_REFERENCE_FULL = "## Reference xlsx file - Full Snapshot\n\n{snapshot}"
    EXCEL_OUTPUT_FULL = "## Model Output xlsx file - Full Snapshot\n\n{snapshot}"
    EXCEL_REFERENCE_DIFF = "## Reference Output xlsx file - Changes from Source\n\n{diff}"
    EXCEL_OUTPUT_DIFF = "## Model Output xlsx file - Changes from Source\n\n{diff}"
    EXCEL_SOURCE_RICH = "## Source xlsx file - Rich Snapshot\n\n{snapshot}"
    EXCEL_SOURCE_SIMPLE = "## Source xlsx file - Simple Snapshot\n\n{snapshot}"
    EXCEL_REFERENCE_SHEET = "Reference Output - Sheet: {sheet_name}"
    EXCEL_OUTPUT_SHEET = "Model Output - Sheet: {sheet_name}"
    EXCEL_SOURCE_SHEET = "Source File - Sheet: {sheet_name}"
    
    EXCEL_REFERENCE_MISSING = (
    "## Reference file missing\n"
    "Reference Excel file was not found: {path}\n"
    "Only available files will be evaluated."
    )

    EXCEL_OUTPUT_MISSING = (
    "## Output file missing\n"
    "Output Excel file was not found: {path}\n"
    "Only available files will be evaluated."
    )
    
    # General evaluation captions
    GENERAL_INSTRUCTION = "## Task Instruction\n\n{instruction}"
    GENERAL_REFERENCE = "## Reference Answer\n\n{content}"
    GENERAL_OUTPUT = "## Model Output\n\n{content}"
    
    # Preprocess info captions
    PREPROCESS_TEXT = "{description}\n\n{text}"
    PREPROCESS_IMAGE = "{description}"
    
    # Metadata captions
    METADATA_SEPARATOR = "\n\n---\n\n"


# ==================== Logging Settings ====================

LOG_LEVEL = "INFO"
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
