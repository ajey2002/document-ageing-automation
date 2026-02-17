from datetime import date
from pathlib import Path
from openpyxl import Workbook

from .config import AppConfig
from .data_loader import load_and_validate_export
from .report_builder import build_summary_sheet, build_account_sheet


def build_output_filename() -> str:
    return f"Final_Report_Automated_{date.today().isoformat()}.xlsx"


def run(config: AppConfig, logger) -> Path:
    logger.info("Starting Document Ageing Automation...")

    df = load_and_validate_export(str(config.input_xls), logger=logger)
    logger.info(f"Valid rows after cleaning: {len(df)}")

    wb = Workbook()
    # remove default sheet
    wb.remove(wb.active)

    # Summary
    build_summary_sheet(wb, config.summary_sheet, df)
    logger.info("Summary sheet created.")

    # Accounts
    accounts = sorted(df["_AccountClean"].unique())
    logger.info(f"Accounts found: {len(accounts)}")

    for acct in accounts:
        sub = df[df["_AccountClean"] == acct].copy()
        build_account_sheet(wb, acct, sub)
        logger.info(f"Created sheet for account: {acct} | rows: {len(sub)}")

    config.output_dir.mkdir(parents=True, exist_ok=True)
    out_path = config.output_dir / build_output_filename()
    wb.save(str(out_path))

    logger.info(f"Workbook saved: {out_path}")
    logger.info("Completed successfully.")
    return out_path
