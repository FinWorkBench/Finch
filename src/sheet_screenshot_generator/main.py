import argparse

from .logging_config import configure_logging
from .screenshot_generator import kill_excel_processes
from .screenshot_generator import (
    convert_excel_to_screenshots,
    process_dataset_folders,
)


def main() -> None:
    configure_logging()  # single place to configure logging

    parser = argparse.ArgumentParser(
        description="Excel screenshot generator using xlwings for better reliability"
    )
    parser.add_argument(
        "--input-folder",
        help="Folder containing Excel files (screenshots/CSVs/metadata saved in same folder)",
    )
    parser.add_argument(
        "--dataset",
        default="dataset",
        help="Process all Excel files in the dataset directory",
    )
    parser.add_argument(
        "--format",
        choices=["png"],
        default="png",
        help="Output image format (only PNG supported by xlwings)",
    )
    parser.add_argument(
        "--skip-existing",
        action="store_true",
        default=True,
        help="Skip processing files that already have screenshots",
    )
    parser.add_argument(
        "--force-reprocess",
        action="store_true",
        help="Force reprocessing of files even if they already exist",
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        help="Make Excel visible (useful for debugging)",
    )
    # ---- CSV / metadata options: only keep export_* and default to False ----
    parser.add_argument(
        "--export-csv",
        action="store_true",
        default=False,
        help="Export sheet data as CSV files (default: False)",
    )
    parser.add_argument(
        "--export-metadata",
        action="store_true",
        default=False,
        help="Export comprehensive Excel metadata as metadata.json (default: False)",
    )

    args = parser.parse_args()

    # Clean up any existing Excel processes before starting
    kill_excel_processes()

    # Decide export flags (now directly from arguments, default False)
    export_csv = args.export_csv
    export_metadata = args.export_metadata

    if args.input_folder:
        # Process specific input folder
        convert_excel_to_screenshots(
            args.input_folder,
            args.format,
            skip_existing=not args.force_reprocess,
            visible=args.visible,
            export_csv=export_csv,
            export_metadata=export_metadata,
        )
    else:
        # Default: process dataset folders
        process_dataset_folders(
            args.dataset,
            visible=args.visible,
            export_csv=export_csv,
            export_metadata=export_metadata,
        )


if __name__ == "__main__":
    main()
