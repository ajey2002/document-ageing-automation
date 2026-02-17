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


def is_totals_row(row: pd.Series) -> bool:
    """
    Export file contains a last row that is a TOTAL line for amount columns.
    Typically: Account empty/NaN + many fields blank except amount columns.
    We should drop it so we don't create an UNKNOWN account sheet.
    """
    acct = row.get(COL_ACCOUNT, None)
    acct_clean = clean_account(acct)

    # If account missing AND either document date missing OR text/type indicates total => treat as totals row
    doc_date = row.get(COL_DOC_DATE, None)
    doc_type = str(row.get(COL_DOC_TYPE, "")).strip().lower()
    text = str(row.get(COL_TEXT, "")).strip().lower()
    company = str(row.get(COL_COMPANY, "")).strip().lower()

    if acct_clean is not None:
        return False

    # heuristics
    indicators = ["total", "grand total", "totals"]
    if any(x in doc_type for x in indicators):
        return True
    if any(x in text for x in indicators):
        return True
    if company in indicators:
        return True

    # if account empty and date empty and at least one amount exists -> likely totals row
    date_missing = pd.isna(doc_date) or str(doc_date).strip() == ""
    amt_doc = row.get(COL_AMT_DOC, None)
    amt_loc = row.get(COL_AMT_LOC, None)
    has_amount = (pd.notna(amt_doc) and str(amt_doc).strip() != "") or (pd.notna(amt_loc) and str(amt_loc).strip() != "")

    return bool(date_missing and has_amount)


def load_and_validate_export(input_xls_path: str, logger=None) -> pd.DataFrame:
    df = pd.read_excel(input_xls_path)
    df.columns = df.columns.str.strip()

    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"Missing columns in export file: {missing}\nFound: {list(df.columns)}")

    if logger:
        logger.info(f"Loaded export rows: {len(df)}")

    # Drop totals row(s)
    mask_totals = df.apply(is_totals_row, axis=1)
    totals_count = int(mask_totals.sum())
    if totals_count:
        df = df.loc[~mask_totals].copy()
        if logger:
            logger.info(f"Removed totals rows from export: {totals_count}. Remaining rows: {len(df)}")

    # Clean account and parse document date
    df["_AccountClean"] = df[COL_ACCOUNT].apply(clean_account)
    # Only keep rows that have a real account (safety)
    before = len(df)
    df = df[df["_AccountClean"].notna()].copy()
    if logger and len(df) != before:
        logger.warning(f"Dropped rows without valid account: {before - len(df)}")

    df[COL_DOC_DATE] = pd.to_datetime(df[COL_DOC_DATE], errors="coerce").dt.date

    return df
