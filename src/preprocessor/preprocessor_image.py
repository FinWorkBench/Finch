#!/usr/bin/env python3
"""
Image file preprocessor.

Scans metadata for image files in reference_outputs and outputs,
removes them from those fields, and adds to preprocess_info.
"""

import logging
from pathlib import Path
from typing import Dict, List, Any

from preprocessor_base import BasePreprocessor


logger = logging.getLogger(__name__)


class ImagePreprocessor(BasePreprocessor):
    """Preprocessor for image files."""
    
    IMAGE_EXTENSIONS = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.svg', '.tif', '.tiff'}
    
    def can_process(self, file_path: Path) -> bool:
        """Check if file is an image."""
        return file_path.suffix.lower() in self.IMAGE_EXTENSIONS
    
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
        Process image file by scanning metadata.
        
        This preprocessor looks for the file in metadata's reference_outputs
        and outputs fields, removes it from there, and creates preprocess_info entry.
        
        Note: This doesn't process individual files directly but scans metadata.
        The actual processing happens in the scan_and_process_images method.
        """
        # This is handled by scan_and_process_images instead
        return []
    
    @staticmethod
    def scan_and_process_images(
        metadata: Dict[str, Any],
        task_dir: Path,
        config
    ) -> List[Dict[str, Any]]:
        """
        Scan metadata for image files and process them.
        
        Args:
            metadata: Metadata dictionary
            task_dir: Path to the task directory
            config: Configuration object
            
        Returns:
            List of preprocess_info entries for images
        """
        preprocess_info = []
        image_extensions = ImagePreprocessor.IMAGE_EXTENSIONS
        
        # Scan reference_outputs
        reference_outputs = metadata.get("reference_outputs", {})
        if isinstance(reference_outputs, dict):
            files = reference_outputs.get("files", [])
            remaining_files = []
            image_files = []
            
            for filename in files:
                file_ext = Path(filename).suffix.lower()
                if file_ext in image_extensions:
                    image_files.append(filename)
                    logger.info(f"Found image in reference_outputs: {filename}")
                else:
                    remaining_files.append(filename)
            
            # Update reference_outputs to remove image files
            if image_files:
                reference_outputs["files"] = remaining_files
                
                # Add to preprocess_info
                preprocess_info.append({
                    "type": "img",
                    "value": image_files,
                    "description": config.image_config["description"]
                })
        
        # Scan outputs
        outputs = metadata.get("outputs", {})
        if isinstance(outputs, dict):
            files = outputs.get("files", [])
            remaining_files = []
            image_files = []
            
            for filename in files:
                file_ext = Path(filename).suffix.lower()
                if file_ext in image_extensions:
                    image_files.append(filename)
                    logger.info(f"Found image in outputs: {filename}")
                else:
                    remaining_files.append(filename)
            
            # Update outputs to remove image files
            if image_files:
                outputs["files"] = remaining_files
                
                # Add to preprocess_info
                preprocess_info.append({
                    "type": "img",
                    "value": image_files,
                    "description": config.image_config["description"]
                })

        # Optional: include source_files images into preprocess_info.
        # Do not mutate source_files list here.
        if getattr(config, "enable_source_non_excel_preprocess", False):
            source_files = metadata.get("source_files", [])
            if isinstance(source_files, list):
                source_image_files = [
                    filename
                    for filename in source_files
                    if Path(filename).suffix.lower() in image_extensions
                ]
                if source_image_files:
                    preprocess_info.append({
                        "type": "img",
                        "value": source_image_files,
                        "description": config.image_config["description"]
                    })
        
        return preprocess_info
