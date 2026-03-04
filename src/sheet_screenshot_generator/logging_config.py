import logging


def configure_logging(log_file: str = "excel_processor.py.log") -> None:
    """Configure root logger for the Excel processing utilities."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        filename="logs/screenshot_generator.log",
        filemode="a",
    )
