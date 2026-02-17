from dataclasses import dataclass
from pathlib import Path

COL_COMPANY = "Comapany"
COL_ACCOUNT = "Account"
COL_DOC_DATE = "Document Date"
COL_DOC_TYPE = "Document Type"
COL_DOC_CURR = "Document currency"
COL_AMT_DOC = "Amount in doc. curr."
COL_LOC_CURR = "Local Currency"
COL_AMT_LOC = "Amount in local currency"
COL_TEXT = "Text"

REQUIRED_COLS = [
    COL_COMPANY, COL_ACCOUNT, COL_DOC_DATE, COL_DOC_TYPE,
    COL_DOC_CURR, COL_AMT_DOC, COL_LOC_CURR, COL_AMT_LOC, COL_TEXT
]


@dataclass(frozen=True)
class AppConfig:
    input_xls: Path
    output_dir: Path
    logs_dir: Path
    summary_sheet: str = "Summary"
