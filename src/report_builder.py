from datetime import date
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from .excel_styles import setup_table_styles
from .config import (
    COL_COMPANY,
    COL_DOC_DATE,
    COL_DOC_TYPE,
    COL_DOC_CURR,
    COL_AMT_DOC,
    COL_LOC_CURR,
    COL_AMT_LOC,
    COL_TEXT,
)


def autosize_columns(ws, min_width=12, max_width=60):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(val))
        ws.column_dimensions[col_letter].width = max(min_width, min(max_width, max_len + 2))


def build_summary_sheet(wb: Workbook, sheet_name: str, df: pd.DataFrame):
    styles = setup_table_styles()
    ws = wb.create_sheet(sheet_name)

    # Title row (merged)
    ws.merge_cells("A1:F1")
    ws["A1"].value = f"Document Ageing Report as at {date.today().strftime('%d.%m.%Y')}"
    ws["A1"].font = styles["title_font"]
    ws["A1"].alignment = styles["title_align"]

    # Table headers (row 3)
    headers = [
        "Comapany",
        "Account",
        "Document currency",
        "Amount in doc. curr.",
        "Local Currency",
        "Amount in local currency",
    ]

    header_row = 3
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=c, value=h)
        cell.fill = styles["header_fill"]
        cell.font = styles["header_font"]
        cell.alignment = styles["header_align"]
        cell.border = styles["border"]

    # Build summary (group + sum)
    summary_df = (
        df.groupby([COL_COMPANY, "_AccountClean", COL_DOC_CURR, COL_LOC_CURR], dropna=False)
          .agg({COL_AMT_DOC: "sum", COL_AMT_LOC: "sum"})
          .reset_index()
          .sort_values(["_AccountClean", COL_COMPANY])
    )

    start_row = header_row + 1
    r = start_row
    for i in range(len(summary_df)):
        ws.cell(r, 1).value = summary_df.iloc[i][COL_COMPANY]
        ws.cell(r, 2).value = summary_df.iloc[i]["_AccountClean"]
        ws.cell(r, 3).value = summary_df.iloc[i][COL_DOC_CURR]
        ws.cell(r, 4).value = float(summary_df.iloc[i][COL_AMT_DOC])
        ws.cell(r, 5).value = summary_df.iloc[i][COL_LOC_CURR]
        ws.cell(r, 6).value = float(summary_df.iloc[i][COL_AMT_LOC])

        for c in range(1, 7):
            ws.cell(r, c).border = styles["border"]
        r += 1

    # number formats
    for rr in range(start_row, r):
        ws.cell(rr, 4).number_format = "#,##0.00;(#,##0.00)"
        ws.cell(rr, 6).number_format = "#,##0.00;(#,##0.00)"

    ws.freeze_panes = "A4"
    autosize_columns(ws)
    return ws


def build_account_sheet(wb: Workbook, account: str, sub_df: pd.DataFrame):
    styles = setup_table_styles()
    ws = wb.create_sheet(title=account[:31])

    # Title
    ws.merge_cells("A1:J1")
    ws["A1"].value = f"Account: {account}  |  As at {date.today().strftime('%d.%m.%Y')}"
    ws["A1"].font = styles["title_font"]
    ws["A1"].alignment = styles["title_align"]

    headers = [
        "Comapany", "Account", "Document Date", "Document Type", "Document currency",
        "Amount in doc. curr.", "Local Currency", "Amount in local currency", "Text", "Doc Ageing"
    ]

    header_row = 3
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=c, value=h)
        cell.fill = styles["header_fill"]
        cell.font = styles["header_font"]
        cell.alignment = styles["header_align"]
        cell.border = styles["border"]

    start_row = header_row + 1

    for i in range(len(sub_df)):
        r = start_row + i

        ws.cell(r, 1).value = sub_df.iloc[i][COL_COMPANY]
        ws.cell(r, 2).value = sub_df.iloc[i]["_AccountClean"]
        ws.cell(r, 3).value = sub_df.iloc[i][COL_DOC_DATE]
        ws.cell(r, 4).value = sub_df.iloc[i][COL_DOC_TYPE]
        ws.cell(r, 5).value = sub_df.iloc[i][COL_DOC_CURR]
        ws.cell(r, 6).value = sub_df.iloc[i][COL_AMT_DOC]
        ws.cell(r, 7).value = sub_df.iloc[i][COL_LOC_CURR]
        ws.cell(r, 8).value = sub_df.iloc[i][COL_AMT_LOC]
        ws.cell(r, 9).value = sub_df.iloc[i][COL_TEXT]

        # Doc Ageing (TODAY - Document Date col C)
        ws.cell(r, 10).value = f"=TODAY()-C{r}"

        for c in range(1, 11):
            ws.cell(r, c).border = styles["border"]

    last_data_row = start_row + len(sub_df) - 1
    total_row = last_data_row + 1 if len(sub_df) else start_row

    # Totals in F and H
    if len(sub_df):
        ws.cell(total_row, 6).value = f"=SUM(F{start_row}:F{last_data_row})"
        ws.cell(total_row, 8).value = f"=SUM(H{start_row}:H{last_data_row})"

    for c in range(1, 11):
        cell = ws.cell(total_row, c)
        cell.border = styles["border"]
        cell.font = styles["total_font"]

    # Formats
    ws.freeze_panes = "A4"
    ws.column_dimensions["I"].width = 60

    for rr in range(start_row, total_row + 1):
        ws.cell(rr, 6).number_format = "#,##0.00;(#,##0.00)"
        ws.cell(rr, 8).number_format = "#,##0.00;(#,##0.00)"

    autosize_columns(ws, min_width=12, max_width=70)
    return ws
