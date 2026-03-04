"""
Excel Sheet Screenshot Generator

Generate screenshots for specified sheets in an Excel workbook.
Screenshots are saved in a folder named <workbook>_shot/ next to the Excel file.

Usage:
    from excel_shot_generator import ExcelShotGenerator

    generator = ExcelShotGenerator(visible=False)
    results = generator.generate_screenshots(
        excel_path="report.xlsx",
        sheet_names=["Sheet1", "Summary"]
    )

    # Results: [{"path": "/path/to/report_shot/Sheet1.png", "sheet_name": "Sheet1"}, ...]

Requirements:
    pip install xlwings psutil
"""

import os
import time
import logging
from typing import List, Dict, Optional

import psutil
import xlwings as xw
from xlwings.constants import CellType, Calculation


# ==================== Process Management ====================

def kill_excel_process(pid: int) -> None:
    """Kill a single Excel process by pid (only if it is Excel)."""
    try:
        proc = psutil.Process(pid)
    except psutil.NoSuchProcess:
        return
    try:
        name = (proc.name() or "").lower()
        if name == "excel.exe":
            logging.debug(f"Killing Excel process pid={pid}")
            proc.kill()
            time.sleep(0.2)
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        pass


# ==================== Screenshot Generator ====================

class ExcelShotGenerator:
    """
    Excel sheet screenshot generator using xlwings.

    Interface kept same as original code1.

    Implementation updated:
    - Use xlwings Range.to_png() (code2 approach)
    - Fallback to ActiveWindow.VisibleRange if visual range screenshot fails
    """

    def __init__(
        self,
        visible: bool = False,
        max_rows: int = 200,
        max_cols: int = 50,
        skip_hidden_sheets: bool = True,   # ✅ NEW (backward compatible)
        kill_orphan_excels: bool = True,
    ):
        self.visible = visible
        self.max_rows = max_rows
        self.max_cols = max_cols
        self.skip_hidden_sheets = skip_hidden_sheets
        self.kill_orphan_excels = kill_orphan_excels

    def _is_sheet_visible(self, sheet: xw.Sheet) -> bool:
        """
        Excel COM visibility:
          -1: xlSheetVisible
           0: xlSheetHidden
           2: xlSheetVeryHidden
        """
        try:
            return int(sheet.api.Visible) == -1
        except Exception:
            # If we can't query visibility, be conservative and treat as visible
            return True


    def _get_visual_range(self, sheet: xw.Sheet) -> xw.main.Range:
        """
        Compute the visual range for screenshot.

        Algorithm (same spirit as code1, but compatible with code2 screenshot method):
        1. Start with UsedRange (cells with data)
        2. Expand to include SpecialCells (constants and formulas)
        3. Expand to include all shapes (charts, images, etc.)
        4. Add padding (3 cols, 2 rows)
        5. Clamp to max_rows/max_cols
        """
        used = sheet.api.UsedRange
        if used is None or (used.Count == 1 and used.Value in (None, "")):
            return sheet.range("A1:A1")

        # Initial bounds from UsedRange
        orig_first_row = used.Row
        orig_first_col = used.Column
        orig_last_row = used.Row + used.Rows.Count - 1
        orig_last_col = used.Column + used.Columns.Count - 1

        first_row = orig_first_row
        first_col = orig_first_col
        last_row = orig_last_row
        last_col = orig_last_col

        # Expand to include all data areas (constants and formulas)
        data_areas = []
        try:
            for cell_type in (CellType.xlCellTypeConstants, CellType.xlCellTypeFormulas):
                try:
                    rng = used.SpecialCells(cell_type)
                except Exception:
                    continue

                try:
                    areas = list(rng.Areas)
                    if not areas:
                        data_areas.append(rng)
                    else:
                        data_areas.extend(areas)
                except Exception:
                    data_areas.append(rng)

            if data_areas:
                first_row = min(a.Row for a in data_areas)
                first_col = min(a.Column for a in data_areas)
                last_row = max(a.Row + a.Rows.Count - 1 for a in data_areas)
                last_col = max(a.Column + a.Columns.Count - 1 for a in data_areas)
        except Exception:
            first_row = orig_first_row
            first_col = orig_first_col
            last_row = orig_last_row
            last_col = orig_last_col

        # Add padding
        last_col += 3
        last_row += 2

        # Expand to include shapes and charts
        try:
            shapes = sheet.api.Shapes
            for shape in shapes:
                tl = shape.TopLeftCell
                br = shape.BottomRightCell
                first_row = min(first_row, tl.Row)
                first_col = min(first_col, tl.Column)
                last_row = max(last_row, br.Row)
                last_col = max(last_col, br.Column)
        except Exception:
            pass

        # Clamp to sheet bounds and max limits
        sheet_max_rows = sheet.cells.rows.count
        sheet_max_cols = sheet.cells.columns.count

        first_row = max(1, first_row)
        first_col = max(1, first_col)
        last_row = min(last_row, first_row + self.max_rows - 1, sheet_max_rows)
        last_col = min(last_col, first_col + self.max_cols - 1, sheet_max_cols)

        return sheet.range((first_row, first_col), (last_row, last_col))

    def generate_screenshots(
        self,
        excel_path: str,
        sheet_names: Optional[List[str]] = None
    ) -> List[Dict[str, str]]:
        excel_path = os.path.abspath(excel_path)

        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"Excel file not found: {excel_path}")

        excel_dir = os.path.dirname(excel_path)
        excel_basename = os.path.splitext(os.path.basename(excel_path))[0]
        output_folder = os.path.join(excel_dir, f"{excel_basename}_shot")
        os.makedirs(output_folder, exist_ok=True)

        results: List[Dict[str, str]] = []

        app_pid = None
        try:
            print(f"[INFO] Opening workbook: {os.path.basename(excel_path)}")

            with xw.App(visible=self.visible, add_book=False) as app:
                app_pid = getattr(app, "pid", None)
                # best-effort settings
                try:
                    app.display_alerts = False
                except Exception:
                    pass

                # ⚠️ Do NOT set screen_updating = False (may cause blank screenshots)
                # Ensure it's ON for rendering
                try:
                    app.api.ScreenUpdating = True
                except Exception:
                    pass

                # reduce heavy calc where possible
                try:
                    app.calculation = "manual"
                except Exception:
                    try:
                        app.api.Calculation = Calculation.xlCalculationManual
                    except Exception:
                        pass

                wb = app.books.open(excel_path, read_only=True, update_links=False)

                available_sheets = [s.name for s in wb.sheets]

                if sheet_names is None:
                    sheets_to_process = available_sheets
                    print(f"[INFO] Processing all {len(sheets_to_process)} sheets")
                else:
                    invalid_sheets = set(sheet_names) - set(available_sheets)
                    if invalid_sheets:
                        raise ValueError(
                            f"Invalid sheet names: {invalid_sheets}. "
                            f"Available sheets: {available_sheets}"
                        )
                    sheets_to_process = sheet_names
                    print(f"[INFO] Processing {len(sheets_to_process)} selected sheets")

                for sheet_name in sheets_to_process:
                    try:
                        sheet = wb.sheets[sheet_name]

                        # ✅ NEW: optionally skip hidden sheets
                        if self.skip_hidden_sheets and (not self._is_sheet_visible(sheet)):
                            print(f"[INFO] Skipping hidden sheet: {sheet_name}")
                            continue

                        # Activate and allow Excel to render
                        sheet.activate()
                        time.sleep(0.25)

                        safe_sheet_name = "".join(
                            c if c.isalnum() or c in (" ", "-", "_") else "_"
                            for c in sheet_name
                        )
                        screenshot_path = os.path.join(output_folder, f"{safe_sheet_name}.png")

                        try:
                            visual_range = self._get_visual_range(sheet)
                            visual_range.to_png(screenshot_path)
                        except Exception as screenshot_error:
                            logging.warning(
                                f"Visual range screenshot failed for '{sheet_name}', "
                                f"trying VisibleRange fallback: {screenshot_error}"
                            )
                            try:
                                visible_range = app.api.ActiveWindow.VisibleRange
                                sheet.range(visible_range.Address).to_png(screenshot_path)
                            except Exception as fallback_error:
                                raise RuntimeError(
                                    f"All screenshot methods failed: {fallback_error}"
                                ) from fallback_error

                        print(f"[SUCCESS] Generated screenshot: {safe_sheet_name}.png")
                        results.append({"path": os.path.abspath(screenshot_path), "sheet_name": sheet_name})

                    except Exception as e:
                        print(f"[ERROR] Failed to screenshot sheet '{sheet_name}': {e}")
                        logging.error(f"Sheet screenshot error for '{sheet_name}': {e}", exc_info=True)
                        continue

                try:
                    wb.close()
                except Exception:
                    pass

        except Exception as e:
            print(f"[ERROR] Failed to process workbook: {e}")
            logging.error(f"Workbook processing error: {e}", exc_info=True)
            raise

        finally:
            time.sleep(0.5)
            if self.kill_orphan_excels:
                try:
                    if app_pid:
                        kill_excel_process(app_pid)
                except Exception:
                    pass

        print(f"[INFO] Total screenshots generated: {len(results)}")
        print(f"[INFO] Output folder: {output_folder}")

        return results


# ==================== Convenience Function ====================

def generate_excel_screenshots(
    excel_path: str,
    sheet_names: Optional[List[str]] = None,
    visible: bool = False,
    max_rows: int = 200,
    max_cols: int = 50
) -> List[Dict[str, str]]:
    """
    Convenience function to generate Excel screenshots.

    (Interface unchanged from code1)
    """
    generator = ExcelShotGenerator(visible=visible, max_rows=max_rows, max_cols=max_cols)
    return generator.generate_screenshots(excel_path, sheet_names)


# ==================== Command Line Interface ====================

def main():
    """Command line interface for the screenshot generator."""
    import argparse

    parser = argparse.ArgumentParser(
        description="Generate screenshots for Excel worksheet sheets"
    )
    parser.add_argument(
        "excel_file",
        type=str,
        help="Path to the Excel file"
    )
    parser.add_argument(
        "-s", "--sheets",
        type=str,
        nargs="+",
        default=None,
        help="Sheet names to screenshot (default: all sheets)"
    )
    parser.add_argument(
        "-v", "--visible",
        action="store_true",
        help="Show Excel UI during processing"
    )
    parser.add_argument(
        "--max-rows",
        type=int,
        default=200,
        help="Maximum rows per screenshot (default: 200)"
    )
    parser.add_argument(
        "--max-cols",
        type=int,
        default=50,
        help="Maximum columns per screenshot (default: 50)"
    )
    parser.add_argument(
        "--log-level",
        type=str,
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        default="INFO",
        help="Logging level (default: INFO)"
    )

    args = parser.parse_args()

    # Configure logging
    logging.basicConfig(
        level=getattr(logging, args.log_level),
        format="%(asctime)s - %(levelname)s - %(message)s"
    )

    try:
        results = generate_excel_screenshots(
            excel_path=args.excel_file,
            sheet_names=args.sheets,
            visible=args.visible,
            max_rows=args.max_rows,
            max_cols=args.max_cols
        )

        print("\n" + "="*60)
        print("SCREENSHOT GENERATION COMPLETE")
        print("="*60)

        for result in results:
            print(f"  • {result['sheet_name']}: {result['path']}")

        print(f"\nTotal: {len(results)} screenshots generated")

    except Exception as e:
        print(f"\n[ERROR] Screenshot generation failed: {e}")
        import traceback
        traceback.print_exc()
        return 1

    return 0


if __name__ == "__main__":
    exit(main())
