import logging
from pathlib import Path
from datetime import datetime


def setup_logging(logs_dir: Path, level: int = logging.INFO) -> logging.Logger:
    logs_dir.mkdir(parents=True, exist_ok=True)
    log_file = logs_dir / f"run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

    logger = logging.getLogger("ageing_report")
    logger.setLevel(level)
    logger.handlers.clear()
    logger.propagate = False

    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")

    # Console
    ch = logging.StreamHandler()
    ch.setLevel(level)
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    # File
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setLevel(level)
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    logger.info(f"Logging initialized. Log file: {log_file}")
    return logger
