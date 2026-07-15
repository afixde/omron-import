from pathlib import Path
import logging

LOG_DIR = Path(__file__).resolve().parent.parent / "logs"
LOG_DIR.mkdir(exist_ok=True)

LOG_FILE = LOG_DIR / "omron-import.log"


def get_logger() -> logging.Logger:

    logger = logging.getLogger("omron-import")

    if logger.hasHandlers():
        return logger

    logger.setLevel(logging.INFO)

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(message)s",
        "%Y-%m-%d %H:%M:%S"
    )

    file_handler = logging.FileHandler(
        LOG_FILE,
        encoding="utf-8"
    )
    file_handler.setFormatter(formatter)

#    console_handler = logging.StreamHandler()
#    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
#    logger.addHandler(console_handler)

    return logger