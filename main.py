import argparse
from pathlib import Path
import logging

from src.config import AppConfig
from src.runner import run
from src.logging_setup import setup_logging


def parse_args():
    p = argparse.ArgumentParser(description="Document Ageing Report Automation")
    p.add_argument("--input", default="data/export.xls", help="Path to export.xls")
    p.add_argument("--output-dir", default="output", help="Folder to save generated workbook")
    p.add_argument("--logs-dir", default="logs", help="Folder to save log files")
    p.add_argument("--log-level", default="INFO", help="DEBUG, INFO, WARNING, ERROR")
    return p.parse_args()


def main():
    args = parse_args()
    level = getattr(logging, args.log_level.upper(), logging.INFO)

    cfg = AppConfig(
        input_xls=Path(args.input),
        output_dir=Path(args.output_dir),
        logs_dir=Path(args.logs_dir),
        summary_sheet="Summary",
    )

    logger = setup_logging(cfg.logs_dir, level=level)
    out_path = run(cfg, logger)
    print(f" Done! Created: {out_path}")


if __name__ == "__main__":
    main()
