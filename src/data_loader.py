import pandas as pd
from typing import Optional
from .config import (
    REQUIRED_COLS,
    COL_ACCOUNT,
    COL_DOC_DATE,
    COL_AMT_DOC,
    COL_AMT_LOC,
    COL_COMPANY,
    COL_DOC_CURR,
    COL_LOC_CURR,
    COL_DOC_TYPE,
    COL_TEXT,
)


def clean_account(x) -> Optional[str]:
    """Convert 63010001.0 -> 63010001, return None for missing."""
    if pd.isna(x) or str(x).strip() == "":
        return None
    try:
        return str(int(float(x)))
    except Exception:
        s = str(x).strip()
        return s if s else None


def _parse_document_date(series: pd.Series) -> pd.Series:
    """
    Robust date parsing for SAP/Excel exports:
    - If values are Excel serial numbers (e.g., 44000), convert using origin '1899-12-30'
    - Otherwise parse as normal datetime
    Returns python `date` objects (or NaT -> None).
    """
    s = series.copy()

    # If numeric-looking (Excel serial dates), convert them
    # (Works even if the column is object with mixed types)
    s_numeric = pd.to_numeric(s, errors="coerce")

    # Convert numeric serial dates to datetime (Excel origin)
    dt_from_serial = pd.to_datetime(s_numeric, unit="D", origin="1899-12-30", errors="coerce")

    # Parse as normal datetime (handles strings like 6/18/2020, 2020-06-18, etc.)
    dt_from_text = pd.to_datetime(s, errors="coerce")

    # Prefer text parse if available, else fallback to serial parse
    dt = dt_from_text.combine_first(dt_from_serial)

    return dt.dt.date


def is_totals_row(row: pd.Series) -> bool:
    """
    Export file contains a last row that is a TOTAL line for amount columns.
    Typically: Account empty/NaN + many fields blank except amount columns.
    We should drop it so we don't create an UNKNOWN account sheet.
    """
    acct = row.get(COL_ACCOUNT, None)
    acct_clean = clean_account(acct)

    doc_date = row.get(COL_DOC_DATE, None)
    doc_type = str(row.get(COL_DOC_TYPE, "")).strip().lower()
    text = str(row.get(COL_TEXT, "")).strip().lower()
    company = str(row.get(COL_COMPANY, "")).strip().lower()

    if acct_clean is not None:
        return False

    indicators = ["total", "grand total", "totals"]
    if any(x in doc_type for x in indicators):
        return True
    if any(x in text for x in indicators):
        return True
    if company in indicators:
        return True

    # If account empty and date empty and at least one amount exists 
    date_missing = pd.isna(doc_date) or str(doc_date).strip() == ""
    amt_doc = row.get(COL_AMT_DOC, None)
    amt_loc = row.get(COL_AMT_LOC, None)
    has_amount = (
        (pd.notna(amt_doc) and str(amt_doc).strip() != "")
        or (pd.notna(amt_loc) and str(amt_loc).strip() != "")
    )

    return bool(date_missing and has_amount)


def load_and_validate_export(input_xls_path: str, logger=None) -> pd.DataFrame:
    df = pd.read_excel(input_xls_path)
    df.columns = df.columns.str.strip()

    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"Missing columns in export file: {missing}\nFound: {list(df.columns)}")

    if logger:
        logger.info(f"Loaded export rows: {len(df)}")

    # Preserve original order for SAP "visually grouped" rows
    df["_row"] = range(len(df))

    # Drop totals row(s)
    mask_totals = df.apply(is_totals_row, axis=1)
    totals_count = int(mask_totals.sum())
    if totals_count:
        df = df.loc[~mask_totals].copy()
        if logger:
            logger.info(f"Removed totals rows from export: {totals_count}. Remaining rows: {len(df)}")

    # Clean account
    df["_AccountClean"] = df[COL_ACCOUNT].apply(clean_account)

    # Keep only rows with valid account (prevents UNKNOWN)
    before = len(df)
    df = df[df["_AccountClean"].notna()].copy()
    dropped = before - len(df)
    if logger and dropped:
        logger.warning(f"Dropped rows without valid account: {dropped}")

    # Parse document date robustly
    missing_before = int(df[COL_DOC_DATE].isna().sum())
    df[COL_DOC_DATE] = _parse_document_date(df[COL_DOC_DATE])
    missing_after_parse = int(pd.isna(df[COL_DOC_DATE]).sum())

    if logger:
        logger.info(f"Document Date missing before parse: {missing_before}")
        logger.info(f"Document Date missing after parse : {missing_after_parse}")

    # Forward-fill dates inside Company + Account, preserving original export row order
    df.sort_values(by=[COL_COMPANY, "_AccountClean", "_row"], inplace=True)
    missing_before_ffill = int(pd.isna(df[COL_DOC_DATE]).sum())

    df[COL_DOC_DATE] = df.groupby([COL_COMPANY, "_AccountClean"])[COL_DOC_DATE].ffill()

    missing_after_ffill = int(pd.isna(df[COL_DOC_DATE]).sum())
    if logger:
        logger.info(f"Document Date missing before ffill: {missing_before_ffill}")
        logger.info(f"Document Date missing after  ffill: {missing_after_ffill}")

    # Restore original order
    df.sort_values(by="_row", inplace=True)
    df.drop(columns=["_row"], inplace=True)

    return df
