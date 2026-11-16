import os
import re
import pandas as pd
from edgar import Company, set_identity
from edgar.xbrl import XBRLS

set_identity("Your name your.email@example.com")  # your real name and email

TICKER = "AAPL"
N = 8  # number of 10-Q filings to stitch

company = Company(TICKER)
filings = company.get_filings(form="10-Q", amendments=False).head(N)
xbrls = XBRLS.from_filings(filings)
st = xbrls.statements

income = st.income_statement(max_periods=N)
balance = st.balance_sheet(max_periods=N)
try:
    cashflow = st.cash_flow_statement(max_periods=N)
except AttributeError:
    cashflow = st.cashflow_statement(max_periods=N)

YEAR = re.compile(r"20\d\d")


def clean(stmt):
    df = stmt.to_dataframe()
    if "label" not in df.columns:
        df = df.reset_index()
    # remove segment / product / region breakdown rows and empty header rows
    if "dimension" in df.columns:
        df = df[df["dimension"] != True]
    if "is_abstract" in df.columns:
        df = df[df["is_abstract"] != True]
    # keep the label plus period columns, skipping year-to-date columns
    period_cols = [c for c in df.columns
                   if YEAR.search(str(c)) and "YTD" not in str(c).upper()]
    print("Period columns kept:", period_cols)
    out = df[["label"] + period_cols].copy()
    out[period_cols] = out[period_cols].apply(pd.to_numeric, errors="coerce")
    out = out.dropna(how="all", subset=period_cols)
    # show large numbers in millions (EPS and other small values stay as they are)
    out[period_cols] = out[period_cols].where(out[period_cols].abs() < 1e5,
                                              out[period_cols] / 1e6)
    return out.drop_duplicates(subset="label").set_index("label").round(2)


filename = f"{TICKER}_quarterly_from_10Q.xlsx"
with pd.ExcelWriter(filename) as writer:
    for name, stmt in [("Income Statement", income),
                       ("Balance Sheet", balance),
                       ("Cash Flow", cashflow)]:
        df = clean(stmt)
        df.to_excel(writer, sheet_name=name)
        ws = writer.sheets[name]
        ws.column_dimensions["A"].width = 55
        for i in range(2, len(df.columns) + 2):
            ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = 16
        for row in ws.iter_rows(min_row=2, min_col=2):
            for cell in row:
                cell.number_format = "#,##0.00"

print("Saved to", os.path.abspath(filename))