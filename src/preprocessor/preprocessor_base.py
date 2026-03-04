#!/usr/bin/env python3
"""
Base classes and configuration for file preprocessing.

This module contains the abstract base classes and configuration
used by all preprocessor implementations.
"""

import logging
from pathlib import Path
from typing import Dict, List, Any, Optional
from abc import ABC, abstractmethod


logger = logging.getLogger(__name__)


class PreprocessorConfig:
    """Configuration for different file type preprocessors."""
    
    def __init__(self):
        # Feature toggle:
        # - False (default): skip source_files non-Excel preprocess for img/md/docx/pdf
        # - True: allow source_files img/md/docx/pdf into preprocess_info
        self.enable_source_non_excel_preprocess = False

        # PDF configuration
        self.pdf_config = {
            "description": "PDF page screenshot"
        }
        
        # Markdown configuration
        self.markdown_config = {
            "text_description": "Markdown text content",
            "img_description": "Markdown embedded image"
        }
        
        # Word configuration
        self.word_config = {
            "text_description": "Word document text content",
            "img_description": "Word document page screenshot"
        }
        
        # Excel configuration
        self.excel_config = {
            "description": "Excel spreadsheet recalculated"
        }
        
        # Image configuration
        self.image_config = {
            "description": "Image file"
        }


class BasePreprocessor(ABC):
    """Abstract base class for file preprocessing."""
    
    def __init__(self, config: PreprocessorConfig):
        """
        Initialize the preprocessor.
        
        Args:
            config: Configuration object
        """
        self.config = config
        self.special_cases: List[str] = []
    
    @abstractmethod
    def can_process(self, file_path: Path) -> bool:
        """
        Check if this preprocessor can handle the given file.
        
        Args:
            file_path: Path to the file
            
        Returns:
            True if this preprocessor can handle the file
        """
        pass
    
    @abstractmethod
    def process(
        self,
        file_path: Path,
        output_dir: Path,
        metadata: Dict[str, Any],
        model_name: str,
        id_str: str,
        kind: str
    ) -> List[Dict[str, Any]]:
        """
        Process a file and return preprocess_info entries.
        
        Args:
            file_path: Path to the file to process
            output_dir: Directory to save processed outputs
            metadata: Metadata dictionary for the task
            model_name: Name of the model
            id_str: ID string
            kind: Kind of file (source, reference, output)
            
        Returns:
            List of preprocess_info dictionaries
        """
        pass
    
    def set_next(self, handler: 'BasePreprocessor') -> 'BasePreprocessor':
        """
        Set the next handler in the chain.
        
        Args:
            handler: Next preprocessor in the chain
            
        Returns:
            The handler that was set
        """
        self.next_handler = handler
        return handler
    
    def handle(
        self,
        file_path: Path,
        output_dir: Path,
        metadata: Dict[str, Any],
        model_name: str,
        id_str: str,
        kind: str
    ) -> Optional[List[Dict[str, Any]]]:
        """
        Handle the file using chain of responsibility pattern.
        
        Args:
            file_path: Path to the file to process
            output_dir: Directory to save processed outputs
            metadata: Metadata dictionary
            model_name: Name of the model
            id_str: ID string
            kind: Kind of file (source, reference, output)
            
        Returns:
            List of preprocess_info dictionaries or None
        """
        if self.can_process(file_path):
            return self.process(file_path, output_dir, metadata, model_name, id_str, kind)
        
        # Try next handler in chain
        if hasattr(self, 'next_handler') and self.next_handler is not None:
            return self.next_handler.handle(
                file_path, output_dir, metadata, model_name, id_str, kind
            )
        
        return None
