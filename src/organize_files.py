#!/usr/bin/env python3
"""
Module to organize dataset files by combining output, source, and reference files.

This module reads a JSONL dataset file, locates corresponding source and reference files,
and organizes them together with model outputs into a structured directory hierarchy.
"""

import os
import json
import argparse
import shutil
import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from collections import defaultdict


# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class FileOrganizer:
    """Organizes dataset files by combining outputs with source and reference files."""
    
    def __init__(self, dataset_dir: str, output_dir: str, target_dir: str):
        """
        Initialize the FileOrganizer.
        
        Args:
            dataset_dir: Root directory of the dataset containing the JSONL file
            output_dir: Directory containing model output subdirectories
            target_dir: Target directory for organized files
        """
        self.dataset_dir = Path(dataset_dir)
        self.output_dir = Path(output_dir)
        self.target_dir = Path(target_dir)
        
        # Cache for file locations to avoid repeated searches
        self.file_cache: Dict[str, List[Path]] = {}
        
    def find_jsonl_file(self) -> Optional[Path]:
        """
        Find the JSONL file in the dataset directory.
        
        Returns:
            Path to the JSONL file, or None if not found
        """
        jsonl_files = list(self.dataset_dir.glob("*.jsonl"))
        
        if len(jsonl_files) == 0:
            logger.error(f"No JSONL file found in {self.dataset_dir}")
            return None
        elif len(jsonl_files) > 1:
            logger.warning(f"Multiple JSONL files found in {self.dataset_dir}, using first one: {jsonl_files[0]}")
        
        return jsonl_files[0]
    
    def build_file_index(self):
        """Build an index of all files in the dataset directory for quick lookup."""
        logger.info("Building file index...")
        
        for file_path in self.dataset_dir.rglob("*"):
            if file_path.is_file():
                filename = file_path.name
                if filename not in self.file_cache:
                    self.file_cache[filename] = []
                self.file_cache[filename].append(file_path)
        
        logger.info(f"Indexed {len(self.file_cache)} unique filenames")
    
    def find_file(self, filename: str, base_path: Optional[Path] = None) -> Optional[Path]:
        """
        Find a file in the dataset directory.
        
        Args:
            filename: Name of the file or relative path
            base_path: Base path for relative path resolution
            
        Returns:
            Path to the file, or None if not found or multiple matches found
        """
        # Check if it's a relative path
        if '/' in filename or '\\' in filename:
            # It's a relative path
            if base_path:
                full_path = base_path / filename
            else:
                full_path = self.dataset_dir / filename
            
            if full_path.exists() and full_path.is_file():
                return full_path
            else:
                logger.error(f"File not found at relative path: {filename}")
                return None
        
        # Search in the file cache
        if filename in self.file_cache:
            matches = self.file_cache[filename]
            
            if len(matches) == 1:
                return matches[0]
            elif len(matches) > 1:
                logger.error(f"Filename points to multiple files: {filename}")
                logger.error(f"  Found at: {[str(m) for m in matches]}")
                return None
        
        logger.error(f"File not found: {filename}")
        return None
    
    def find_output_files(self, model_name: str, file_id: str) -> List[Tuple[Path, str]]:
        """
        Find all output files for a given ID in a model's output directory.
        
        Args:
            model_name: Name of the model (subdirectory in output_dir)
            file_id: ID to search for
            
        Returns:
            List of tuples (file_path, suffix) where suffix is empty for single files
            or "_1", "_2", etc. for multiple files
        """
        model_output_dir = self.output_dir / model_name
        
        if not model_output_dir.exists():
            logger.warning(f"Model output directory not found: {model_output_dir}")
            return []
        
        output_files = []
        seen_paths = set()
        
        # Search for files matching the pattern
        # Pattern 1: exact match (e.g., "0.xlsx")
        for file_path in model_output_dir.rglob(f"{file_id}.*"):
            if file_path.is_file() and file_path not in seen_paths:
                output_files.append((file_path, ""))
                seen_paths.add(file_path)
        
        # Pattern 2: numbered matches (e.g., "0_1.xlsx", "0_2.xlsx")
        for file_path in model_output_dir.rglob(f"{file_id}_*.*"):
            if file_path.is_file():
                stem = file_path.stem
                # Extract the number part
                if stem.startswith(f"{file_id}_"):
                    suffix_part = stem[len(file_id):]
                    if file_path not in seen_paths:
                        output_files.append((file_path, suffix_part))
                        seen_paths.add(file_path)

        # Pattern 3: dashed matches (e.g., "0-1.png", "0-2.png")
        for file_path in model_output_dir.rglob(f"{file_id}-*.*"):
            if file_path.is_file():
                stem = file_path.stem
                if stem.startswith(f"{file_id}-"):
                    suffix_part = stem[len(file_id):]
                    if file_path not in seen_paths:
                        output_files.append((file_path, suffix_part))
                        seen_paths.add(file_path)
        
        return output_files
    
    def read_txt_files(self, file_paths: List[Path]) -> str:
        """
        Read and concatenate content from text files.
        
        Args:
            file_paths: List of text file paths
            
        Returns:
            Concatenated content
        """
        content_parts = []
        
        for file_path in file_paths:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content_parts.append(f.read())
            except Exception as e:
                logger.error(f"Error reading text file {file_path}: {e}")
        
        return "\n\n".join(content_parts)
    
    def process_entry(self, entry: Dict, model_name: str) -> bool:
        """
        Process a single JSONL entry for a specific model.
        
        Args:
            entry: Dictionary containing the JSONL entry data
            model_name: Name of the model being processed
            
        Returns:
            True if processing was successful, False otherwise
        """
        file_id = entry.get("id")
        if file_id is None:
            logger.error("Entry missing 'id' field")
            return False
        
        logger.info(f"Processing ID: {file_id} for model: {model_name}")

        # Only create task directory when there is at least one model output.
        output_files = self.find_output_files(model_name, file_id)
        if not output_files:
            logger.info(f"  Skipping ID {file_id}: no output files found")
            return False

        # Create target directory: target_dir/model_name/id/
        target_path = self.target_dir / model_name / file_id
        target_path.mkdir(parents=True, exist_ok=True)
        
        # Find and copy source files
        source_files = entry.get("source_files", [])
        copied_sources = []
        
        for source_file in source_files:
            source_path = self.find_file(source_file)
            if source_path:
                dest_path = target_path / source_path.name
                shutil.copy2(source_path, dest_path)
                copied_sources.append(source_path.name)
                logger.info(f"  Copied source file: {source_path.name}")
            else:
                logger.warning(f"  Source file not found: {source_file}")
        
        # Find and copy reference files
        reference_outputs = entry.get("reference_outputs", {})
        reference_files = reference_outputs.get("files", [])
        copied_references = []
        
        for ref_file in reference_files:
            ref_path = self.find_file(ref_file)
            if ref_path:
                dest_path = target_path / ref_path.name
                shutil.copy2(ref_path, dest_path)
                copied_references.append(ref_path.name)
                logger.info(f"  Copied reference file: {ref_path.name}")
            else:
                logger.warning(f"  Reference file not found: {ref_file}")
        
        # Copy output files
        copied_outputs = []
        txt_content_parts = []
        
        for output_path, suffix in output_files:
            if output_path.suffix.lower() == '.txt':
                # Handle text files separately
                try:
                    with open(output_path, 'r', encoding='utf-8') as f:
                        txt_content_parts.append(f.read())
                    logger.info(f"  Read text content from: {output_path.name}")
                except Exception as e:
                    logger.error(f"  Error reading text file {output_path.name}: {e}")
            else:
                # Copy non-text files
                dest_path = target_path / output_path.name
                shutil.copy2(output_path, dest_path)
                copied_outputs.append(output_path.name)
                logger.info(f"  Copied output file: {output_path.name}")
        
        # Create metadata.json
        metadata = entry.copy()
        
        # Add outputs field
        metadata["outputs"] = {
            "files": copied_outputs,
            "text": "\n\n".join(txt_content_parts) if txt_content_parts else ""
        }
        
        # Save metadata.json
        metadata_path = target_path / "metadata.json"
        try:
            with open(metadata_path, 'w', encoding='utf-8') as f:
                json.dump(metadata, f, indent=2, ensure_ascii=False)
            logger.info(f"  Created metadata.json")
        except Exception as e:
            logger.error(f"  Error writing metadata.json: {e}")
            return False
        
        return True
    
    def organize(self) -> bool:
        """Main method to organize all files.

        Returns:
            True if organization completed successfully, False on critical errors.
        """
        logger.info("Starting file organization...")

        # Find JSONL file
        jsonl_path = self.find_jsonl_file()
        if not jsonl_path:
            return False

        logger.info(f"Using JSONL file: {jsonl_path}")

        # Build file index
        self.build_file_index()

        # Get list of models from output directory
        if not self.output_dir.exists():
            logger.error(f"Output directory does not exist: {self.output_dir}")
            return False

        model_dirs = [d for d in self.output_dir.iterdir() if d.is_dir()]

        if not model_dirs:
            logger.error(f"No model subdirectories found in: {self.output_dir}")
            return False
        
        logger.info(f"Found {len(model_dirs)} model(s): {[d.name for d in model_dirs]}")
        
        # Process each model
        for model_dir in model_dirs:
            model_name = model_dir.name
            logger.info(f"\nProcessing model: {model_name}")
            
            # Read and process JSONL entries
            try:
                with open(jsonl_path, 'r', encoding='utf-8') as f:
                    entry_count = 0
                    success_count = 0
                    
                    for line_num, line in enumerate(f, 1):
                        line = line.strip()
                        if not line:
                            continue
                        
                        try:
                            entry = json.loads(line)
                            entry_count += 1
                            
                            if self.process_entry(entry, model_name):
                                success_count += 1
                        except json.JSONDecodeError as e:
                            logger.error(f"Error parsing JSON on line {line_num}: {e}")
                    
                    logger.info(f"Model {model_name}: Processed {success_count}/{entry_count} entries successfully")
            
            except Exception as e:
                logger.error(f"Error reading JSONL file: {e}")
        
        logger.info("\nFile organization complete!")
        return True


def main():
    """Main entry point for the script."""
    parser = argparse.ArgumentParser(
        description="Organize dataset files by combining outputs with source and reference files."
    )
    parser.add_argument(
        "--dataset-dir",
        required=True,
        help="Path to the dataset root directory containing the JSONL file"
    )
    parser.add_argument(
        "--output-dir",
        required=True,
        help="Path to the directory containing model output subdirectories"
    )
    parser.add_argument(
        "--target-dir",
        required=True,
        help="Path to the target directory for organized files"
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
    
    # Create organizer and run
    organizer = FileOrganizer(
        dataset_dir=args.dataset_dir,
        output_dir=args.output_dir,
        target_dir=args.target_dir
    )

    success = organizer.organize()
    if not success:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
