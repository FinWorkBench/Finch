import os
import time
import json
import logging
from pathlib import Path
from typing import List, Tuple, Optional, Sequence, Set

import psutil
import xlwings as xw
from xlwings.constants import CellType, Calculation

from .metadata_extractor import EnhancedExcelExtractor
import tqdm
from .diff_utils import _compute_focus_sheets_for_modify, _get_file_mode_from_property, get_changed_sheet_names_for_pair, get_formatting_diff_sheet_names, _get_wb_normal

def kill_excel_processes() -> None:
    """Kill any orphaned Excel processes more efficiently."""
    excel_processes = []
    for proc in psutil.process_iter(["name", "pid"]):
        try:
            if proc.info["name"] and proc.info["name"].lower() == "excel.exe":
                excel_processes.append(proc)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass

    if excel_processes:
        logging.debug(f"Killing {len(excel_processes)} Excel processes...")
        for proc in excel_processes:
            try:
                proc.kill()
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        time.sleep(0.5)
    else:
        logging.debug("No Excel processes found, skipping kill step")


class ExcelScreenshotGenerator:
    """Excel screenshot generator using xlwings for better resource management."""

    def __init__(self, visible: bool = False):
        """
        Args:
            visible: If True, Excel UI will be visible (useful for debugging).
        """
        self.visible = visible

    # ---------- internal helpers ----------

    def _get_visual_range(
        self,
        sheet: xw.Sheet,
        max_rows: int = 200,
        max_cols: int = 50,
    ) -> xw.main.Range:
        """
        Compute a visual range for screenshot using UsedRange + SpecialCells
        and shapes (charts, etc.), then clamp by max_rows/max_cols.
        """
        used = sheet.api.UsedRange
        if used is None or (used.Count == 1 and used.Value in (None, "")):
            return sheet.range("A1:A1")

        orig_first_row = used.Row
        orig_first_col = used.Column
        orig_last_row = used.Row + used.Rows.Count - 1
        orig_last_col = used.Column + used.Columns.Count - 1

        first_row = orig_first_row
        first_col = orig_first_col
        last_row = orig_last_row
        last_col = orig_last_col

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
            # fallback to original UsedRange bounds
            first_row = orig_first_row
            first_col = orig_first_col
            last_row = orig_last_row
            last_col = orig_last_col

        # add some padding
        last_col += 3
        last_row += 2

        # include shapes/charts
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

        sheet_max_rows = sheet.cells.rows.count
        sheet_max_cols = sheet.cells.columns.count

        first_row = max(1, first_row)
        first_col = max(1, first_col)
        last_row = min(last_row, first_row + max_rows - 1, sheet_max_rows)
        last_col = min(last_col, first_col + max_cols - 1, sheet_max_cols)

        return sheet.range((first_row, first_col), (last_row, last_col))

    # ---------- public APIs ----------

    def convert_workbook_to_screenshots_and_csv(
        self,
        excel_path: str,
        format: str = "png",
        export_csv: bool = False,
        export_metadata: bool = False,
    ) -> Tuple[List[str], List[str], Optional[str]]:
        """
        Convert a single workbook to sheet screenshots, CSVs, and metadata.json.

        All files are saved next to the Excel file:
        - <workbook>_<sheet_index>.png
        - <workbook>_<sheet_index>.csv (if enabled)
        - metadata.json (if enabled)

        Returns:
            (list_of_screenshot_paths, list_of_csv_paths, metadata_path_or_None)
        """
        excel_path = os.path.abspath(excel_path)
        output_folder = os.path.dirname(excel_path)

        if format.lower() != "png":
            logging.warning("xlwings only supports PNG format. Using PNG.")
            format = "png"

        created_screenshots: List[str] = []
        created_csvs: List[str] = []
        metadata_file: Optional[str] = None

        try:
            # clean up Excel processes before starting
            kill_excel_processes()
            time.sleep(1)

            print(f"    >> Opening workbook: {os.path.basename(excel_path)}")

            # metadata via openpyxl
            if export_metadata:
                metadata_path = os.path.join(output_folder, "metadata.json")
                if not os.path.exists(metadata_path):
                    try:
                        print("    >> Extracting comprehensive metadata...")
                        extractor = EnhancedExcelExtractor(excel_path)
                        metadata = extractor.extract_all_metadata()

                        with open(metadata_path, "w", encoding="utf-8") as f:
                            json.dump(metadata, f, indent=2, default=str)

                        metadata_file = metadata_path
                        print("    OK Generated metadata.json")
                    except ValueError as ve:
                        if "old Excel 97-2003 format" in str(ve):
                            print(
                                f"    !! Skipping metadata extraction: "
                                f"{os.path.basename(excel_path)} is in old .xls format"
                            )
                            print(
                                "       (Screenshots and CSV will still be generated via xlwings)"
                            )
                        else:
                            print(
                                f"    !! Failed to extract metadata from "
                                f"{os.path.basename(excel_path)}: {ve}"
                            )
                            print(f"       File path: {excel_path}")
                    except Exception as metadata_error:
                        print(
                            f"    !! Failed to extract metadata from "
                            f"{os.path.basename(excel_path)}: {metadata_error}"
                        )
                        print(f"       File path: {excel_path}")
                else:
                    print("    OK Metadata already exists: metadata.json")
                    metadata_file = metadata_path

            # use xlwings for screenshots/CSV
            with xw.App(visible=self.visible, add_book=False) as app:
                # best-effort: disable alerts & heavy calculation
                try:
                    app.display_alerts = False
                except Exception:
                    pass

                try:
                    app.calculation = "manual"
                except Exception:
                    try:
                        app.api.Calculation = Calculation.xlCalculationManual
                    except Exception:
                        pass

                try:
                    wb = app.books.open(
                        excel_path, read_only=True, update_links=False
                    )

                    sheet_count = len(wb.sheets)
                    print(f"    -> Found {sheet_count} worksheet(s)")
                    print(f"    -> Processing {sheet_count} sheet(s)...")

                    for sheet_idx, sheet in enumerate(wb.sheets):
                        try:
                            sheet_name = sheet.name

                            # skip helper/hidden-like sheets by name convention
                            if (
                                "Assistant2302a3b1df77" in sheet_name
                                or sheet_name.startswith("_")
                            ):
                                print(f"    >> Skipping sheet '{sheet_name}'")
                                continue

                            sheet.activate()

                            workbook_stem = os.path.splitext(
                                os.path.basename(excel_path)
                            )[0]
                            screenshot_name = f"_{workbook_stem}_{sheet_idx}.png"
                            csv_name = f"_{workbook_stem}_{sheet_idx}.csv"

                            screenshot_path = os.path.join(
                                output_folder, screenshot_name
                            )
                            csv_path = os.path.join(output_folder, csv_name)

                            # screenshot
                            if not os.path.exists(screenshot_path):
                                try:
                                    visual_range = self._get_visual_range(sheet)
                                    visual_range.to_png(screenshot_path)
                                    created_screenshots.append(screenshot_path)
                                    print(
                                        f"    OK Generated screenshot "
                                        f"for sheet '{sheet_name}' -> {screenshot_name}"
                                    )
                                except Exception as screenshot_error:
                                    print(
                                        f"    !! Failed to screenshot sheet "
                                        f"'{sheet_name}': {screenshot_error}"
                                    )
                                    # fallback: visible range
                                    try:
                                        visible_range = app.api.ActiveWindow.VisibleRange
                                        sheet.range(
                                            visible_range.Address
                                        ).to_png(screenshot_path)
                                        created_screenshots.append(screenshot_path)
                                        print(
                                            "    OK Generated screenshot using "
                                            f"visible range for sheet '{sheet_name}' -> "
                                            f"{screenshot_name}"
                                        )
                                    except Exception as fallback_error:
                                        print(
                                            "    !! All screenshot methods failed "
                                            f"for sheet '{sheet_name}': {fallback_error}"
                                        )
                            else:
                                print(
                                    f"    OK Screenshot already exists: {screenshot_name}"
                                )
                                created_screenshots.append(screenshot_path)

                            # CSV
                            if export_csv:
                                if not os.path.exists(csv_path):
                                    try:
                                        used_range = sheet.used_range
                                        if used_range is not None:
                                            data = used_range.value

                                            # normalize to 2D list
                                            if not isinstance(data, list):
                                                data = [[data]]
                                            elif data and not isinstance(
                                                data[0], list
                                            ):
                                                data = [data]

                                            import csv

                                            with open(
                                                csv_path,
                                                "w",
                                                newline="",
                                                encoding="utf-8",
                                            ) as csvfile:
                                                writer = csv.writer(csvfile)
                                                for row in data:
                                                    cleaned_row = [
                                                        ""
                                                        if cell is None
                                                        else str(cell)
                                                        for cell in (
                                                            row
                                                            if isinstance(
                                                                row, list
                                                            )
                                                            else [row]
                                                        )
                                                    ]
                                                    writer.writerow(cleaned_row)

                                            created_csvs.append(csv_path)
                                            print(
                                                "    OK Generated CSV "
                                                f"for sheet '{sheet_name}' -> {csv_name}"
                                            )
                                        else:
                                            # empty sheet
                                            with open(
                                                csv_path,
                                                "w",
                                                newline="",
                                                encoding="utf-8",
                                            ):
                                                pass
                                            created_csvs.append(csv_path)
                                            print(
                                                "    OK Generated empty CSV "
                                                f"for sheet '{sheet_name}' -> {csv_name}"
                                            )
                                    except Exception as csv_error:
                                        print(
                                            "    !! Failed to generate CSV for "
                                            f"sheet '{sheet_name}': {csv_error}"
                                        )
                                else:
                                    print(f"    OK CSV already exists: {csv_name}")
                                    created_csvs.append(csv_path)

                        except Exception as sheet_error:
                            print(
                                f"    !! Error processing sheet "
                                f"{sheet_idx + 1}: {sheet_error}"
                            )
                            continue

                except Exception as workbook_error:
                    print(f"    !! Error opening workbook: {workbook_error}")
                    raise

        except Exception as e:
            print(f"!! Error processing {excel_path}: {e}")
            raise

        return created_screenshots, created_csvs, metadata_file

    def screenshot_selected_sheets(
        self,
        excel_path: str,
        sheet_names: List[str],
        max_rows: int = 200,
        max_cols: int = 50,
    ) -> List[str]:
        """
        Screenshot only the given sheet names.

        Images are saved as <workbook_stem>_<sheet_index>.png
        in the same folder as the workbook, where sheet_index is
        the original index (0-based) in the workbook.

        Returns:
            List of generated screenshot paths.
        """
        excel_path = os.path.abspath(excel_path)
        output_folder = os.path.dirname(excel_path)
        workbook_stem = os.path.splitext(os.path.basename(excel_path))[0]

        created_screenshots: List[str] = []
        sheet_name_set = set(sheet_names)

        if not sheet_name_set:
            print(
                f"    >> No sheet names specified for {excel_path}, "
                "nothing to screenshot."
            )
            return created_screenshots

        with xw.App(visible=self.visible, add_book=False) as app:
            try:
                app.display_alerts = False
            except Exception:
                pass

            try:
                app.calculation = "manual"
            except Exception:
                try:
                    app.api.Calculation = Calculation.xlCalculationManual
                except Exception:
                    pass

            wb = app.books.open(excel_path, read_only=True, update_links=False)

            for sheet_idx, sheet in enumerate(wb.sheets):
                if sheet.name not in sheet_name_set:
                    continue

                sheet.activate()

                screenshot_name = f"_{workbook_stem}_{sheet_idx}.png"
                screenshot_path = os.path.join(output_folder, screenshot_name)

                try:
                    visual_range = self._get_visual_range(
                        sheet, max_rows=max_rows, max_cols=max_cols
                    )
                    visual_range.to_png(screenshot_path)
                    created_screenshots.append(screenshot_path)
                    print(
                        "    OK Generated diff screenshot for sheet "
                        f"'{sheet.name}' -> {screenshot_name}"
                    )
                except Exception as e:
                    print(
                        "    !! Failed to screenshot sheet "
                        f"'{sheet.name}' in {excel_path}: {e}"
                    )

        return created_screenshots

def convert_excel_to_screenshots(
    input_folder: str, 
    format: str = "png", 
    skip_existing: bool = True,
    visible: bool = False,
    export_csv: bool = False,
    export_metadata: bool = False
) -> None:
    """Batch convert all Excel workbooks to screenshots, CSV files, and metadata saved in the same folder as each file.

    Args:
        input_folder: Path to folder containing Excel files
        format: Format to save screenshots (png only supported)
        skip_existing: Skip processing files that already have screenshots/CSV/metadata
        visible: Make Excel visible (useful for debugging)
        export_csv: Whether to also export CSV files
        export_metadata: Whether to also export metadata.json
    """
    input_folder = os.path.abspath(input_folder)
    
    # Get list of Excel files
    excel_files = [
        file for file in os.listdir(input_folder)
        if file.lower().endswith((".xlsx", ".xls", ".xlsm"))
    ]
    
    if not excel_files:
        print(f"No Excel files found in {input_folder}")
        return
    
    # Check which files need processing
    files_to_process = []
    skipped_files = 0
    
    for file in excel_files:
        # Excel ，/CSV 
        workbook_stem = os.path.splitext(file)[0]
        screenshot_path = os.path.join(input_folder, f"_{workbook_stem}_0.png")
        csv_path = os.path.join(input_folder, f"_{workbook_stem}_0.csv") if export_csv else None
        metadata_path = os.path.join(input_folder, "metadata.json") if export_metadata else None
        
        files_exist = os.path.exists(screenshot_path)
        if export_csv:
            files_exist = files_exist and os.path.exists(csv_path)
        if export_metadata:
            files_exist = files_exist and os.path.exists(metadata_path)
        
        if skip_existing and files_exist:
            skipped_files += 1
        else:
            files_to_process.append(file)

    
    if skipped_files > 0:
        export_types = []
        if True:  # screenshots always
            export_types.append("screenshots")
        if export_csv:
            export_types.append("CSV files")
        if export_metadata:
            export_types.append("metadata")
        print(f"Skipping {skipped_files} files that already have {' and '.join(export_types)}")
    
    if not files_to_process:
        print("All files already processed. Nothing to do.")
        return
    
    export_types = []
    if True:  # screenshots always
        export_types.append("screenshots")
    if export_csv:
        export_types.append("CSV files")
    if export_metadata:
        export_types.append("metadata")
    print(f"Processing {len(files_to_process)} files to generate {' and '.join(export_types)}...")
    
    # Create generator instance
    generator = ExcelScreenshotGenerator(visible=visible)
    
    # Process each Excel file
    for file in tqdm(files_to_process, desc="Converting workbooks", unit="file"):
        excel_path = os.path.join(input_folder, file)
        
        try:
            created_screenshots, created_csvs, metadata_file = generator.convert_workbook_to_screenshots_and_csv(
                excel_path, format, export_csv, export_metadata
            )
            
            output_parts = [f"{len(created_screenshots)} screenshots"]
            if export_csv:
                output_parts.append(f"{len(created_csvs)} CSV files")
            if export_metadata and metadata_file:
                output_parts.append("metadata")
            
            tqdm.write(f"  OK Created {' and '.join(output_parts)} for {file}")
            
        except Exception as e:
            msg = f"Error processing {file}: {e}"
            tqdm.write(f"  !! {msg}")
            logging.warning(msg)
            
            continue
    
    export_types = []
    if True:  # screenshots always
        export_types.append("Screenshots")
    if export_csv:
        export_types.append("CSV files")
    if export_metadata:
        export_types.append("metadata")
    print(f"\n{' and '.join(export_types)} saved alongside Excel files.")
    
    
def process_dataset_folders(
    dataset_path: str = "dataset",
    visible: bool = False,
    export_csv: bool = False,       
    export_metadata: bool = False,  
) -> None:

    dataset_path = os.path.abspath(dataset_path)

    if not os.path.exists(dataset_path):
        print(f"!! Dataset directory not found: {dataset_path}")
        return

    print(f">> Processing dataset (diff screenshots only) under: {dataset_path}")

    gen = ExcelScreenshotGenerator(visible=visible)
    total_cases = 0

    # ， case 
    for entry in sorted(os.listdir(dataset_path)):
        case_dir = os.path.join(dataset_path, entry)
        if not os.path.isdir(case_dir):
            continue

        case_path = Path(case_dir)
        input_xlsx = case_path / "input.xlsx"
        answer_xlsx = case_path / "answer.xlsx"
        output_xlsx = case_path / "output.xlsx"

        # property.json file_mode（ modify） 
        file_mode = _get_file_mode_from_property(case_path)

        # ===== NEW: generate “ modify” ===== 
        
        # 1) file_mode generate 
        # 2) input.xlsx / answer.xlsx / output.xlsx 
        # 3) input answer sheet name 
        # 4) input output sheet name 
        if file_mode == "generate":
            if input_xlsx.is_file() and answer_xlsx.is_file() and output_xlsx.is_file():
                try:
                    wb_input = _get_wb_normal(input_xlsx)
                    wb_answer = _get_wb_normal(answer_xlsx)
                    wb_output = _get_wb_normal(output_xlsx)

                    input_sheets = set(wb_input.sheetnames)
                    answer_inter = input_sheets & set(wb_answer.sheetnames)
                    output_inter = input_sheets & set(wb_output.sheetnames)

                    if answer_inter and output_inter:
                        print(
                            "   >> file_mode=generate 但发现："
                            "input & answer 的 sheet 有交集，且 input & output 的 sheet 也有交集，"
                            "本 case 按 modify 逻辑处理。"
                        )
                        file_mode = "modify"
                except Exception as e:
                    print(
                        f"   !! Failed to evaluate generate->modify condition in case '{entry}': {e}"
                    )
        # ===== NEW END =====

        print(f"\n>> Case: {entry} (file_mode={file_mode})")

        # ===================== generate ===================== 
        if file_mode == "generate":
            # generate ， input.xlsx ； 
            # output / answer sheet ： 
            # - （0-based ） 
            # - CSV（ export_csv=True） 
            # - metadata（ export_metadata=True） 
            for wb_path, label in [
                (output_xlsx, "output.xlsx"),
                (answer_xlsx, "answer.xlsx"),
            ]:
                if not wb_path.is_file():
                    continue

                try:
                    created_screenshots, created_csvs, metadata_file = gen.convert_workbook_to_screenshots_and_csv(
                        str(wb_path),
                        format="png",
                        export_csv=export_csv,
                        export_metadata=export_metadata,
                    )
                    parts = [f"{len(created_screenshots)} screenshots"]
                    if export_csv:
                        parts.append(f"{len(created_csvs)} CSV files")
                    if export_metadata and metadata_file:
                        parts.append(f"metadata ({os.path.basename(metadata_file)})")

                    print(f"   - Generated {' and '.join(parts)} for {label}")
                except Exception as e:
                    msg = f"Failed to process {label} in case '{entry}' (generate mode): {e}"
                    print(f"   !! {msg}")
                    logging.warning(msg)

            if (not output_xlsx.is_file()) and (not answer_xlsx.is_file()):
                print("   -> No output.xlsx or answer.xlsx found in generate mode, skip case.")

            total_cases += 1
            # generate diff ， case 
            continue

        # ===================== modify （） ===================== 
        # input 
        if not input_xlsx.is_file():
            print(f"   -> No input.xlsx found in modify mode, skip case.")
            continue

        # answer.xlsx， sheet input ， 
        focus_sheets: Optional[Set[str]] = None
        if answer_xlsx.is_file():
            focus_sheets = _compute_focus_sheets_for_modify(input_xlsx, answer_xlsx)

        # 1) input vs output：“”（///）， 
        changed_io_core: List[str] = []
        if output_xlsx.is_file():
            try:
                changed_io_core = get_changed_sheet_names_for_pair(
                    input_xlsx, output_xlsx, include_formatting=False, focus_sheets=focus_sheets
                )

                if changed_io_core:
                    print(f"   - Core changed sheets (input vs output, no formatting): {changed_io_core}")
                else:
                    print(f"   - No core changes between input.xlsx and output.xlsx")
            except Exception as e:
                msg = f"Failed to diff input vs output in {entry}: {e}"
                print(f"   !! {msg}")
                logging.warning(msg)

        # 2) input vs answer：（） 
        changed_ia: List[str] = []
        if answer_xlsx.is_file():
            try:
                changed_ia = get_changed_sheet_names_for_pair(
                    input_xlsx, answer_xlsx, include_formatting=True, focus_sheets=focus_sheets
                )

                if changed_ia:
                    print(f"   - Changed sheets (input vs answer, with formatting): {changed_ia}")
                else:
                    print(f"   - No changes between input.xlsx and answer.xlsx")
            except Exception as e:
                msg = f"Failed to diff input vs output in {entry}: {e}"
                print(f"   !! {msg}")
                logging.warning(msg)

        # 3) ：input vs answer（） 
        formatting_ia: List[str] = []
        if answer_xlsx.is_file():
            try:
                formatting_ia = get_formatting_diff_sheet_names(
                    input_xlsx, answer_xlsx, focus_sheets=focus_sheets
                )

                if formatting_ia:
                    print(f"   - Sheets with task-relevant formatting changes (input vs answer): {formatting_ia}")
            except Exception as e:
                msg = f"Failed to diff input vs output in {entry}: {e}"
                print(f"   !! {msg}")
                logging.warning(msg)

        # 4) ：input vs output（） 
        formatting_io: List[str] = []
        if output_xlsx.is_file():
            try:
                formatting_io = get_formatting_diff_sheet_names(
                    input_xlsx, output_xlsx, focus_sheets=focus_sheets
                )

                if formatting_io:
                    print(f"   - Sheets with formatting changes (input vs output): {formatting_io}")
            except Exception as e:
                msg = f"Failed to diff input vs output in {entry}: {e}"
                print(f"   !! {msg}")
                logging.warning(msg)

        
        if not changed_io_core and not changed_ia and not formatting_io and not formatting_ia:
            print(f"   -> No meaningful changes found for this case, skip screenshots.")
            continue

        
        formatting_task_sheets = set(formatting_ia)
        formatting_io_task_sheets = set(formatting_io) & formatting_task_sheets
        changed_io_for_output = sorted(set(changed_io_core) | formatting_io_task_sheets)
        changed_for_input = sorted(set(changed_ia) | set(changed_io_for_output))

        if changed_for_input:
            try:
                gen.screenshot_selected_sheets(str(input_xlsx), changed_for_input)
            except Exception as e:
                msg = f"Failed to screenshot input.xlsx in case '{entry}': {e}"
                print(f"   !! {msg}")
                logging.warning(msg)

        if changed_io_for_output and output_xlsx.is_file():
            try:
                gen.screenshot_selected_sheets(str(output_xlsx), changed_io_for_output)
            except Exception as e:
                msg = f"Failed to screenshot output.xlsx in case '{entry}': {e}"
                print(f"   !! {msg}")
                logging.warning(msg)

        if changed_ia and answer_xlsx.is_file():
            try:
                gen.screenshot_selected_sheets(str(answer_xlsx), changed_ia)
            except Exception as e:
                msg = f"Failed to screenshot answer.xlsx in case '{entry}': {e}"
                print(f"   !! {msg}")
                logging.warning(msg)

        total_cases += 1

    print(f"\nOK Diff screenshots generated for {total_cases} case folder(s).")



