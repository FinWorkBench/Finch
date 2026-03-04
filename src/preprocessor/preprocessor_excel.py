#!/usr/bin/env python3
"""
Excel file preprocessor with recalculation functionality.

Recalculates all formulas in Excel workbooks and saves them.
Includes both preprocessor integration and standalone script functionality.
"""

import os
import glob
import time
import shutil
import logging
import argparse
from pathlib import Path
from typing import Dict, List, Any, Optional

from preprocessor_base import BasePreprocessor

try:
    import xlwings as xw
except ImportError:
    xw = None
try:
    import openpyxl
except ImportError:
    openpyxl = None


logger = logging.getLogger(__name__)


class ExcelPreprocessor(BasePreprocessor):
    """Preprocessor for Excel files."""
    
    EXCEL_EXTENSIONS = {'.xlsx', '.xls', '.xlsm', '.csv'}
    
    def __init__(self, config, excel_app=None):
        """
        Initialize Excel preprocessor.
        
        Args:
            config: Configuration object
            excel_app: xlwings App instance (shared across all processing)
        """
        super().__init__(config)
        self.excel_app = excel_app
        self._owns_app = False
        
        if self.excel_app is None and xw is not None:
            # Create our own app if not provided
            try:
                self.excel_app = xw.App(visible=False, add_book=False)
                self.excel_app.display_alerts = False
                self.excel_app.screen_updating = False
                self._owns_app = True
                logger.info("Created Excel application instance")
            except Exception as e:
                logger.warning(f"Failed to create Excel app: {e}")
                self.excel_app = None
    
    def __del__(self):
        """Clean up Excel app if we own it."""
        if self._owns_app and self.excel_app is not None:
            try:
                self.excel_app.quit()
                logger.info("Closed Excel application instance")
            except:
                pass
    
    def can_process(self, file_path: Path) -> bool:
        """Check if file is an Excel file."""
        return file_path.suffix.lower() in self.EXCEL_EXTENSIONS
    
    def process(
        self,
        file_path: Path,
        output_dir: Path,
        metadata: Dict[str, Any],
        model_name: str,
        id_str: str,
        kind: str
    ) -> List[Dict[str, Any]]:
        """Recalculate Excel file and save to original location."""
        
        if xw is None:
            logger.warning(f"xlwings not installed, cannot process Excel: {file_path}")
            self.special_cases.append(
                f"[NO_XLWINGS]\tmodel={model_name}\tid={id_str}\tkind={kind}\t"
                f"file={file_path.name}"
            )
            return []
        
        if self.excel_app is None:
            logger.warning(f"Excel app not available, cannot process: {file_path}")
            self.special_cases.append(
                f"[NO_EXCEL_APP]\tmodel={model_name}\tid={id_str}\tkind={kind}\t"
                f"file={file_path.name}"
            )
            return []
        
        # Pre-flight checks
        if not file_path.exists():
            logger.warning(f"File does not exist: {file_path}")
            self.special_cases.append(
                f"[FILE_NOT_EXIST]\tmodel={model_name}\tid={id_str}\tkind={kind}\t"
                f"file={file_path.name}"
            )
            return []
        
        # Check file size (skip if too large or zero)
        file_size = file_path.stat().st_size
        if file_size == 0:
            logger.warning(f"File is empty (0 bytes): {file_path}")
            self.special_cases.append(
                f"[EMPTY_FILE]\tmodel={model_name}\tid={id_str}\tkind={kind}\t"
                f"file={file_path.name}"
            )
            return []
        
        if file_size > 100 * 1024 * 1024:  # 100MB
            logger.warning(f"File too large ({file_size / 1024 / 1024:.1f}MB), skipping: {file_path}")
            self.special_cases.append(
                f"[FILE_TOO_LARGE]\tmodel={model_name}\tid={id_str}\tkind={kind}\t"
                f"file={file_path.name}\tsize={file_size}"
            )
            return []
        
        # Work directly with the original file
        file_path_abs = str(file_path.absolute())
        temp_repair_path = str(file_path.with_suffix(f".repair{file_path.suffix}").absolute())
        
        # Check if file is already open
        temp_file = file_path.parent / f"~${file_path.name}"
        if temp_file.exists():
            logger.warning(f"File appears to be open in Excel: {file_path}")
            self.special_cases.append(
                f"[FILE_ALREADY_OPEN]\tmodel={model_name}\tid={id_str}\tkind={kind}\t"
                f"file={file_path.name}"
            )
            return []
        
        try:
            logger.info(
                f"Processing Excel file: {model_name} / {id_str}: {kind} {file_path.name} "
                f"({file_size / 1024:.1f}KB)"
            )
            
            # Try to open the file with Excel's built-in repair
            wb = None
            open_success = False
            last_error = None
            
            open_strategies = [
                {"read_only": False, "corrupt_load": 1, "label": "repair (read/write)"},
                {"read_only": False, "corrupt_load": 2, "label": "extract (read/write)"},
                {"read_only": True, "corrupt_load": 1, "label": "repair (read-only)"},
                {"read_only": True, "corrupt_load": 2, "label": "extract (read-only)"},
                {"read_only": False, "corrupt_load": 0, "label": "normal (read/write)"},
            ]

            for strat in open_strategies:
                logger.info(f"  Opening file with Excel {strat['label']}...")
                try:
                    wb = self.excel_app.books.open(
                        file_path_abs,
                        update_links=False,
                        read_only=strat["read_only"],
                        ignore_read_only_recommended=True,
                        notify=False,
                        corrupt_load=strat["corrupt_load"]
                    )
                    open_success = True
                    logger.info(f"  OK File opened successfully with {strat['label']}")
                    break
                except Exception as e:
                    last_error = e
                    logger.warning(
                        f"  Failed to open file with {strat['label']}: {type(e).__name__}: {e}"
                    )

            repaired_via_openpyxl = False
            openpyxl_rewrite_ok = False
            if not open_success and openpyxl is not None:
                logger.info("  Excel open failed, attempting openpyxl rewrite repair...")
                try:
                    keep_vba = file_path.suffix.lower() == '.xlsm'
                    wb_py = openpyxl.load_workbook(file_path_abs, data_only=False, keep_vba=keep_vba)
                    wb_py.save(temp_repair_path)
                    wb_py.close()
                    openpyxl_rewrite_ok = True
                    try:
                        wb = self.excel_app.books.open(
                            temp_repair_path,
                            update_links=False,
                            read_only=False,
                            ignore_read_only_recommended=True,
                            notify=False,
                            corrupt_load=0
                        )
                        open_success = True
                        repaired_via_openpyxl = True
                        logger.info("  OK File opened after openpyxl rewrite")
                    except Exception as e:
                        last_error = e
                        logger.warning(f"  Failed to open openpyxl-rewritten file in Excel: {type(e).__name__}: {e}")
                except Exception as e:
                    last_error = e
                    logger.warning(f"  Failed openpyxl repair: {type(e).__name__}: {e}")

            if not open_success and openpyxl_rewrite_ok:
                try:
                    try:
                        if os.path.exists(file_path_abs):
                            os.remove(file_path_abs)
                        shutil.move(temp_repair_path, file_path_abs)
                    except PermissionError:
                        # If delete/rename is blocked, overwrite contents in-place
                        shutil.copyfile(temp_repair_path, file_path_abs)
                        try:
                            os.remove(temp_repair_path)
                        except Exception:
                            pass

                    logger.info("  OK Rewrote file with openpyxl; skipping Excel recalculation")
                    self.special_cases.append(
                        f"[EXCEL_OPENPYXL_REWRITE]\tmodel={model_name}\tid={id_str}\t"
                        f"kind={kind}\tfile={file_path.name}"
                    )
                    return []
                except Exception as e:
                    last_error = e
                    logger.warning(f"  Failed to replace file after openpyxl rewrite: {e}")
                    try:
                        if os.path.exists(temp_repair_path):
                            os.remove(temp_repair_path)
                    except Exception:
                        pass

            if not open_success:
                # All strategies failed including repair
                error_msg = str(last_error) if last_error else "Unknown error"

                # Categorize the error
                error_type_lower = error_msg.lower()
                if "unexpected" in error_type_lower or "???" in error_msg:
                    error_type = "EXCEL_UNEXPECTED_ERROR"
                elif "invalid" in error_type_lower or "???" in error_msg:
                    error_type = "EXCEL_INVALID_FILE"
                elif "password" in error_type_lower or "???" in error_msg:
                    error_type = "EXCEL_PASSWORD_PROTECTED"
                elif "corrupt" in error_type_lower or "???" in error_msg:
                    error_type = "EXCEL_CORRUPTED"
                else:
                    error_type = "EXCEL_OPEN_FAIL"

                logger.warning(f"Cannot open Excel file after trying all strategies: {file_path_abs}")
                logger.warning(f"Last error: {repr(last_error)}")
                self.special_cases.append(
                    f"[{error_type}]	model={model_name}	id={id_str}	"
                    f"kind={kind}	file={file_path.name}	size={file_size}	error={repr(last_error)}"
                )
                return []
            
            # Recalculate and save
            try:
                # Check if workbook is read-only
                if wb.api.ReadOnly:
                    logger.warning(f"Workbook opened read-only, attempting SaveAs repair: {file_path}")
                    try:
                        if os.path.exists(temp_repair_path):
                            os.remove(temp_repair_path)
                        wb.api.SaveAs(
                            temp_repair_path,
                            FileFormat=51,  # xlOpenXMLWorkbook (.xlsx)
                            ConflictResolution=2,  # xlLocalSessionChanges
                            Local=True
                        )
                        wb.close()
                        wb = None
                        shutil.move(temp_repair_path, file_path_abs)
                        logger.info("  ✓ Saved repaired copy from read-only workbook")

                        wb = self.excel_app.books.open(
                            file_path_abs,
                            update_links=False,
                            read_only=False,
                            ignore_read_only_recommended=True,
                            notify=False,
                            corrupt_load=0
                        )
                    except Exception as e:
                        logger.warning(f"Failed to SaveAs repaired copy from read-only workbook: {e}")
                        try:
                            if wb:
                                wb.close()
                        except Exception:
                            pass
                        self.special_cases.append(
                            f"[EXCEL_READ_ONLY]\tmodel={model_name}\tid={id_str}\t"
                            f"kind={kind}\tfile={file_path.name}\terror={repr(e)}"
                        )
                        return []
                
                # Force full recalculation
                logger.debug(f"  Starting recalculation...")
                self.excel_app.calculation = "manual"
                self.excel_app.api.CalculateFullRebuild()
                self.excel_app.calculation = "automatic"
                
                # Small delay to let Excel finish background work
                time.sleep(0.2)
                
                # Save the workbook to original path
                # This will save the repaired and recalculated version
                logger.debug(f"  Saving workbook to {file_path_abs}...")
                wb.save(file_path_abs)
                if 'repaired_via_openpyxl' in locals() and repaired_via_openpyxl:
                    try:
                        if os.path.exists(temp_repair_path):
                            os.remove(temp_repair_path)
                    except Exception:
                        pass
                logger.debug(f"  Save completed")
                
                logger.info(
                    f"[OK] {model_name} / {id_str}: {kind} Excel repaired, recalculated and saved -> "
                    f"{file_path}"
                )
                
                # Return empty list - don't add to metadata
                return []
                
            except Exception as e:
                # If Excel recalculation fails (e.g., RPC errors), try openpyxl rewrite as a fallback.
                if openpyxl is not None:
                    try:
                        try:
                            if wb:
                                wb.close()
                        except Exception:
                            pass
                        keep_vba = file_path.suffix.lower() == '.xlsm'
                        wb_py = openpyxl.load_workbook(file_path_abs, data_only=False, keep_vba=keep_vba)
                        wb_py.save(temp_repair_path)
                        wb_py.close()
                        shutil.copyfile(temp_repair_path, file_path_abs)
                        try:
                            os.remove(temp_repair_path)
                        except Exception:
                            pass
                        self.special_cases.append(
                            f"[EXCEL_OPENPYXL_REWRITE_AFTER_RECALC_FAIL]\tmodel={model_name}\t"
                            f"id={id_str}\tkind={kind}\tfile={file_path.name}\terror={repr(e)}"
                        )
                        return []
                    except Exception:
                        pass

                error_msg = f"Error during Excel processing: {file_path_abs}, error: {repr(e)}"
                logger.warning(error_msg)
                self.special_cases.append(
                    f"[EXCEL_RECALC_FAIL]\tmodel={model_name}\tid={id_str}\t"
                    f"kind={kind}\tfile={file_path.name}\terror={repr(e)}"
                )
                return []
            
            finally:
                # Always close the workbook
                try:
                    if wb:
                        logger.debug(f"  Closing workbook...")
                        wb.close()
                        logger.debug(f"  Workbook closed")
                except Exception as close_error:
                    logger.debug(f"  Error closing workbook: {close_error}")
        
        except Exception as e:
            logger.warning(f"Failed to process Excel file: {file_path}, error: {e}")
            self.special_cases.append(
                f"[EXCEL_PROCESS_FAIL]\tmodel={model_name}\tid={id_str}\t"
                f"kind={kind}\tfile={file_path.name}\terror={e}"
            )
            return []


# Standalone Excel recalculation functionality
def setup_logging_for_standalone(log_file: str = "recalc_excel.log"):
    """Setup logging configuration for standalone mode."""
    log_dir = Path(log_file).parent
    if log_dir != Path('.') and not log_dir.exists():
        log_dir.mkdir(parents=True, exist_ok=True)
    
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[
            logging.FileHandler(log_file, mode="a"),
            logging.StreamHandler()
        ]
    )


def recalculate_excel_files(folder: str, pattern: str = "**/*.xlsx"):
    """
    Recalculate all Excel workbooks under a folder (standalone mode).
    
    Args:
        folder: Root folder to search
        pattern: Glob pattern for Excel files (default: **/*.xlsx)
    """
    if xw is None:
        print("ERROR: xlwings is not installed. Please install it with: pip install xlwings")
        return False
    
    logger_inst = logging.getLogger(__name__)
    
    # Build full pattern
    full_pattern = os.path.join(folder, pattern)
    
    # Start hidden Excel application
    logger_inst.info("Starting Excel application...")
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    
    processed_count = 0
    skipped_count = 0
    error_count = 0
    
    try:
        # Recursively iterate over all .xlsx files
        excel_files = sorted(glob.iglob(full_pattern, recursive=True))
        total_files = len(excel_files)
        
        logger_inst.info(f"Found {total_files} Excel file(s) to process")
        
        for i, fp in enumerate(excel_files, 1):
            # Skip non-regular files
            if not os.path.isfile(fp):
                continue
            
            # Skip temporary Excel files
            if os.path.basename(fp).startswith("~$"):
                logger_inst.info(f"[SKIP] Temporary file: {fp}")
                skipped_count += 1
                continue
            
            fp_abs = os.path.abspath(fp)
            print(f"\n[{i}/{total_files}] Processing: {fp_abs}")
            logger_inst.info(f"Processing: {fp_abs}")
            
            # Open workbook
            try:
                wb = app.books.open(fp_abs, update_links=False, read_only=False)
            except Exception as e:
                msg = f"[SKIP] Cannot open: {fp_abs}, error: {repr(e)}"
                print(msg)
                logger_inst.warning(msg)
                skipped_count += 1
                continue
            
            # Recalculate and save
            try:
                # Force full recalculation
                app.calculation = "manual"
                app.api.CalculateFullRebuild()
                app.calculation = "automatic"
                
                # Small delay to let Excel finish background work
                time.sleep(0.2)
                
                # Save the workbook
                wb.save()
                
                logger_inst.info(f"[OK] Recalculated and saved: {fp_abs}")
                processed_count += 1
                
            except Exception as e:
                msg = f"[ERROR] Processing failed: {fp_abs}, error: {repr(e)}"
                print(msg)
                logger_inst.warning(msg)
                error_count += 1
            
            finally:
                # Always close the workbook
                try:
                    wb.close()
                except:
                    pass
        
        # Summary
        print("\n" + "=" * 70)
        print("PROCESSING SUMMARY")
        print("=" * 70)
        print(f"Total files found: {total_files}")
        print(f"Successfully processed: {processed_count}")
        print(f"Skipped: {skipped_count}")
        print(f"Errors: {error_count}")
        print("=" * 70)
        
        logger_inst.info("All workbooks processed.")
        logger_inst.info(f"Summary - Total: {total_files}, Processed: {processed_count}, "
                        f"Skipped: {skipped_count}, Errors: {error_count}")
        
        return True
        
    finally:
        # Ensure Excel is closed
        try:
            app.quit()
            logger_inst.info("Excel application closed")
        except Exception as e:
            logger_inst.warning(f"Error closing Excel application: {e}")


def main():
    """Main entry point for standalone Excel recalculation."""
    parser = argparse.ArgumentParser(
        description="Recalculate all Excel workbooks under a folder (recursively) using xlwings."
    )
    parser.add_argument(
        "--folder",
        "-f",
        required=True,
        help="Root folder to search for Excel files"
    )
    parser.add_argument(
        "--pattern",
        "-p",
        default="**/*.xlsx",
        help="Glob pattern for Excel files (default: **/*.xlsx)"
    )
    parser.add_argument(
        "--log-file",
        "-l",
        default="logs/recalc_excel.log",
        help="Log file path (default: logs/recalc_excel.log)"
    )
    
    args = parser.parse_args()
    
    # Setup logging
    setup_logging_for_standalone(args.log_file)
    logger_inst = logging.getLogger(__name__)
    
    logger_inst.info("=" * 70)
    logger_inst.info("Excel Recalculation Script Started")
    logger_inst.info("=" * 70)
    logger_inst.info(f"Folder: {args.folder}")
    logger_inst.info(f"Pattern: {args.pattern}")
    logger_inst.info(f"Log file: {args.log_file}")
    
    # Run recalculation
    success = recalculate_excel_files(args.folder, args.pattern)
    
    logger_inst.info("Script completed")
    print("\nDone.")
    
    return 0 if success else 1


if __name__ == "__main__":
    exit(main())
