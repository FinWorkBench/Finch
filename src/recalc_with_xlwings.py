# recalc_with_xlwings.py
import os
import glob
import time
import xlwings as xw
import argparse
import logging


# Basic logging configuration; you can adjust level, filename, or mode as needed
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    filename="logs/recalc_with_xlwings.log",
    filemode="a",
)

def main():
    parser = argparse.ArgumentParser(
        description="Recalculate all Excel workbooks under a folder (recursively) using xlwings."
    )
    parser.add_argument(
        "--folder",
        "-f",
        required=True,
        help="Root folder, e.g. LLM eval test data-20251113T111824Z-1-001",
    )
    args = parser.parse_args()

    # Root folder passed from CLI
    FOLDER = args.folder
    pattern = os.path.join(FOLDER, "**", "*.xlsx")

    # Start a hidden Excel application to process all workbooks
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        # Recursively iterate over all .xlsx files under the given folder
        for fp in sorted(glob.iglob(pattern, recursive=True)):
            # Skip non-regular files (just in case)
            if not os.path.isfile(fp):
                continue
            # Skip temporary Excel files
            if os.path.basename(fp).startswith("~$"):
                continue

            fp_abs = os.path.abspath(fp)
            print("Processing:", fp_abs)
            logging.info("Processing: %s", fp_abs)

            # Open workbook; on failure, log and continue with the next file
            try:
                wb = app.books.open(fp_abs, update_links=False, read_only=False)
            except Exception as e:
                msg = f"[SKIP] Failed to open: {fp_abs}, error: {repr(e)}"
                print(msg)
                logging.warning(msg)
                continue

            # Trigger full recalculation and save; on failure, log and continue
            try:
                # Force a full recalculation to ensure all formulas are updated
                app.calculation = "manual"
                app.api.CalculateFullRebuild()
                app.calculation = "automatic"
                time.sleep(0.2)  # Small delay to let Excel finish background work
                wb.save()
                logging.info("Recalculated and saved: %s", fp_abs)
            except Exception as e:
                msg = f"[SKIP] Error during processing: {fp_abs}, error: {repr(e)}"
                print(msg)
                logging.warning(msg)
            finally:
                # Always close the workbook to avoid leaking handles
                wb.close()

        print("Done.")
        logging.info("All workbooks processed.")
    finally:
        # Ensure Excel is closed even if an exception occurs
        app.quit()


if __name__ == "__main__":
    main()
