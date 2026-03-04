"""
Cache Manager for Content Builder

Manages caching of expensive computations like diffs, snapshots, and screenshots.
"""

import json
import logging
from pathlib import Path
from typing import Optional, Dict, Any, List

from .config import (
    CACHE_DIR_NAME,
    CACHE_DIFF_JSON,
    CACHE_FULL_SNAPSHOT_PREFIX,
    CACHE_RICH_SNAPSHOT_PREFIX,
    CACHE_SIMPLE_SNAPSHOT_PREFIX,
    CACHE_SCREENSHOTS_DIR,
)


logger = logging.getLogger(__name__)


class CacheManager:
    """
    Manages caching of expensive computations.
    
    Cache directory structure:
    task_dir/
        _cache/
            diff_cache.json                    # Diff results
            full_snapshot_reference.txt        # Full snapshots
            full_snapshot_output.txt
            rich_snapshot_source.txt           # Rich/simple snapshots
            simple_snapshot_source.txt
            screenshots/                       # Screenshots
                reference_Sheet1.png
                output_Sheet1.png
                source_Sheet1.png
    """
    
    def __init__(self, task_dir: Path):
        """
        Initialize cache manager for a task directory.
        
        Args:
            task_dir: Path to the task directory
        """
        self.task_dir = Path(task_dir)
        self.cache_dir = self.task_dir / CACHE_DIR_NAME
        self.screenshots_dir = self.cache_dir / CACHE_SCREENSHOTS_DIR
        
        # Create cache directories if they don't exist
        self.cache_dir.mkdir(exist_ok=True)
        self.screenshots_dir.mkdir(exist_ok=True)
    
    # ==================== Diff Cache ====================
    
    def get_diff_cache(self, source_name: str, target_name: str) -> Optional[Dict[str, Any]]:
        """
        Get cached diff result.
        
        Args:
            source_name: Source file name (e.g., "input.xlsx")
            target_name: Target file name (e.g., "output.xlsx")
        
        Returns:
            Cached diff dict or None if not found
        """
        cache_file = self.cache_dir / f"diff_{source_name}_{target_name}.json"
        
        if not cache_file.exists():
            logger.debug(f"Diff cache miss: {cache_file.name}")
            return None
        
        try:
            with cache_file.open("r", encoding="utf-8") as f:
                diff = json.load(f)
            logger.debug(f"Diff cache hit: {cache_file.name}")
            return diff
        except Exception as e:
            logger.warning(f"Failed to load diff cache: {e}")
            return None
    
    def save_diff_cache(
        self, 
        source_name: str, 
        target_name: str, 
        diff: Dict[str, Any]
    ) -> None:
        """
        Save diff result to cache.
        
        Args:
            source_name: Source file name
            target_name: Target file name
            diff: Diff dict to cache
        """
        cache_file = self.cache_dir / f"diff_{source_name}_{target_name}.json"
        
        try:
            with cache_file.open("w", encoding="utf-8") as f:
                json.dump(diff, f, ensure_ascii=False, indent=2)
            logger.debug(f"Diff cache saved: {cache_file.name}")
        except Exception as e:
            logger.warning(f"Failed to save diff cache: {e}")
    
    # ==================== Snapshot Cache ====================
    
    def get_snapshot_cache(
        self, 
        file_name: str, 
        snapshot_type: str
    ) -> Optional[str]:
        """
        Get cached snapshot text.
        
        Args:
            file_name: Excel file name (e.g., "output.xlsx")
            snapshot_type: One of "full", "rich", "simple"
        
        Returns:
            Cached snapshot text or None if not found
        """
        prefix_map = {
            "full": CACHE_FULL_SNAPSHOT_PREFIX,
            "rich": CACHE_RICH_SNAPSHOT_PREFIX,
            "simple": CACHE_SIMPLE_SNAPSHOT_PREFIX,
        }
        
        if snapshot_type not in prefix_map:
            logger.warning(f"Invalid snapshot type: {snapshot_type}")
            return None
        
        prefix = prefix_map[snapshot_type]
        cache_file = self.cache_dir / f"{prefix}_{file_name}.txt"
        
        if not cache_file.exists():
            logger.debug(f"Snapshot cache miss: {cache_file.name}")
            return None
        
        try:
            with cache_file.open("r", encoding="utf-8") as f:
                snapshot = f.read()
            logger.debug(f"Snapshot cache hit: {cache_file.name}")
            return snapshot
        except Exception as e:
            logger.warning(f"Failed to load snapshot cache: {e}")
            return None
    
    def save_snapshot_cache(
        self,
        file_name: str,
        snapshot_type: str,
        snapshot: str
    ) -> None:
        """
        Save snapshot text to cache.
        
        Args:
            file_name: Excel file name
            snapshot_type: One of "full", "rich", "simple"
            snapshot: Snapshot text to cache
        """
        prefix_map = {
            "full": CACHE_FULL_SNAPSHOT_PREFIX,
            "rich": CACHE_RICH_SNAPSHOT_PREFIX,
            "simple": CACHE_SIMPLE_SNAPSHOT_PREFIX,
        }
        
        if snapshot_type not in prefix_map:
            logger.warning(f"Invalid snapshot type: {snapshot_type}")
            return
        
        prefix = prefix_map[snapshot_type]
        cache_file = self.cache_dir / f"{prefix}_{file_name}.txt"
        
        try:
            with cache_file.open("w", encoding="utf-8") as f:
                f.write(snapshot)
            logger.debug(f"Snapshot cache saved: {cache_file.name}")
        except Exception as e:
            logger.warning(f"Failed to save snapshot cache: {e}")
    
    # ==================== Screenshot Cache ====================
    
    def get_screenshots_cache(
        self,
        file_name: str,
        sheet_names: List[str]
    ) -> Optional[List[Dict[str, str]]]:
        """
        Get cached screenshots.
        
        Args:
            file_name: Excel file name (e.g., "output.xlsx")
            sheet_names: List of sheet names
        
        Returns:
            List of dicts with "path" and "sheet_name" or None if any missing
        """
        results = []
        base_name = Path(file_name).stem  # e.g., "output"
        
        for sheet_name in sheet_names:
            # Sanitize sheet name for filename
            safe_sheet_name = "".join(
                c if c.isalnum() or c in (' ', '-', '_') else '_'
                for c in sheet_name
            )
            screenshot_path = self.screenshots_dir / f"{base_name}_{safe_sheet_name}.png"
            
            if not screenshot_path.exists():
                logger.debug(f"Screenshot cache miss: {screenshot_path.name}")
                return None
            
            results.append({
                "path": str(screenshot_path.absolute()),
                "sheet_name": sheet_name
            })
        
        logger.debug(f"Screenshots cache hit for {file_name}: {len(results)} sheets")
        return results
    
    def save_screenshots_cache(
        self,
        file_name: str,
        screenshots: List[Dict[str, str]]
    ) -> None:
        """
        Save screenshots to cache by copying them.
        
        Args:
            file_name: Excel file name
            screenshots: List of dicts with "path" and "sheet_name"
        """
        import shutil
        
        base_name = Path(file_name).stem
        
        for screenshot in screenshots:
            source_path = Path(screenshot["path"])
            sheet_name = screenshot["sheet_name"]
            
            # Sanitize sheet name
            safe_sheet_name = "".join(
                c if c.isalnum() or c in (' ', '-', '_') else '_'
                for c in sheet_name
            )
            
            dest_path = self.screenshots_dir / f"{base_name}_{safe_sheet_name}.png"
            
            try:
                shutil.copy2(source_path, dest_path)
            except Exception as e:
                logger.warning(f"Failed to cache screenshot {source_path.name}: {e}")
        
        logger.debug(f"Screenshots cache saved for {file_name}: {len(screenshots)} sheets")
    
    # ==================== Cache Inspection ====================
    
    def get_cache_info(self) -> Dict[str, Any]:
        """
        Get information about cached items.
        
        Returns:
            Dict with cache statistics
        """
        info = {
            "cache_dir": str(self.cache_dir),
            "diff_files": [],
            "snapshot_files": [],
            "screenshot_count": 0
        }
        
        if not self.cache_dir.exists():
            return info
        
        # Count cache files
        for file in self.cache_dir.glob("diff_*.json"):
            info["diff_files"].append(file.name)
        
        for file in self.cache_dir.glob("*_snapshot_*.txt"):
            info["snapshot_files"].append(file.name)
        
        if self.screenshots_dir.exists():
            info["screenshot_count"] = len(list(self.screenshots_dir.glob("*.png")))
        
        return info
