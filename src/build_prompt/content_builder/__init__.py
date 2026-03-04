"""
Content Builder Package

A modular system for building evaluation content parts from model outputs.

Main components:
- ContentBuilder: Main entry point for building content
- ExcelContentBuilder: Specialized builder for Excel evaluations
- GeneralContentBuilder: Builder for general (non-Excel) evaluations
- CacheManager: Handles caching of expensive computations
- Token counter utilities for managing token limits
"""

__version__ = "1.0.0"

from .content_builder import ContentBuilder
from .excel_content_builder import ExcelContentBuilder
from .general_content_builder import GeneralContentBuilder
from .cache_manager import CacheManager
from .config import Captions
from .token_counter import truncate_content_parts

__all__ = [
    "ContentBuilder",
    "ExcelContentBuilder",
    "GeneralContentBuilder",
    "CacheManager",
    "MAX_TOKENS",
    "Captions",
    "truncate_content_parts"
]

