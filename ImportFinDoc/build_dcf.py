"""
build_dcf.py - build an Excel DCF valuation model from a company's latest 10-K.

Step By Step:
    1. Downloads the latest annual financial statements from SEC EDGAR (EdgarTools).
    2. Picks out the numbers a DCF needs (revenue, EBIT, tax, D&A, capex, cash, debt...).
    3. Writes <TICKER>_DCF_Model.xlsx a formula-driven model you can edit in Excel.

How to run
    python build_dcf.py            uses TICKER from the settings below
    python build_dcf.py MSFT       or pass a ticker
    python build_dcf.py MSFT 410   ticker and current share price

The SEC asks everyone who downloads filings to identify themselves. Set the
EDGAR_IDENTITY environment variable ("Your Name your.email@example.com"), or
just run the script and type it when asked. It is never saved in this file.

Requires: edgartools, pandas, openpyxl
Educational project.
"""

import os
import re
import sys
import warnings
from datetime import date

import pandas as pd
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# -----------------------------------------------------------------------------
# Settings: change these, then run the script
# -----------------------------------------------------------------------------
TICKER = "AAPL"       # any US company that files a 10-K
SHARE_PRICE = None    # current share price in $, e.g. 331.90 (or type it into Excel later)

# Discount-rate inputs. These are PLACEHOLDERS, not recommendations.
# They land in yellow cells in the workbook so you can change them there too.
RISK_FREE_RATE = 0.0485        # 10-year US Treasury yield: update to today's
EQUITY_RISK_PREMIUM = 0.05     # commonly 4-6%
BETA = 1.0                     # 1.0 = same risk as the market: look up the company's beta
PRE_TAX_COST_OF_DEBT = 0.055   # roughly the yield on the company's long-term bonds
TAX_RATE_ON_INTEREST = 0.21    # US federal statutory rate
TERMINAL_GROWTH = 0.025        # long-run growth after the forecast period

N_FORECAST = 5                 # forecast years (the sheet layout assumes 5)


# =============================================================================
# Part 1: get the numbers from SEC EDGAR
# =============================================================================

# Columns EdgarTools adds next to the period columns. Everything else is a period.
META_COLUMNS = {
    "concept", "label", "standard_concept", "level", "abstract", "dimension",
    "is_breakdown", "dimension_axis", "dimension_member", "dimension_member_label",
    "dimension_label", "balance", "weight", "preferred_sign", "parent_concept",
    "parent_abstract_concept", "unit", "point_in_time",
}
DATE = re.compile(r"(\d{4}-\d{2}-\d{2})")


def local_name(concept):
    """'us-gaap_NetIncomeLoss' -> 'netincomeloss' (drops the taxonomy prefix)."""
    return re.split(r"[_:]", str(concept), maxsplit=1)[-1].lower()


class Table:
    """One financial statement as plain numbers, newest period first."""

    def __init__(self, df, periods):
        self.df = df            # one row per line item
        self.periods = periods  # period-end dates, newest first


def to_table(statement, annual=True):
    """Turn an EdgarTools statement into a Table of raw filed values.

    view="summary" keeps company totals only (no segment or product rows).
    presentation=False keeps the values as filed: costs and cash outflows are
    positive numbers, so the signs are handled explicitly further down.
    """
    if statement is None:
        return None
    try:
        df = statement.to_dataframe(view="summary", presentation=False)
    except TypeError:
        sys.exit("This script needs a newer EdgarTools. Run:  python -m pip install --upgrade edgartools")
    if df is None or isinstance(df, str) or len(df) == 0:
        return None
    if "concept" not in df.columns or "label" not in df.columns:
        return None

    # Drop segment breakdowns and header rows that carry no numbers.
    for flag in ("dimension", "abstract"):
        if flag in df.columns:
            is_set = df[flag].map(lambda v: bool(v) if pd.notna(v) else False)
            df = df[~is_set]
    df = df.reset_index(drop=True)

    # Period columns are named by end date, e.g. "2026-09-26 (FY)" or "2026-09-26".
    dated = {}
    for col in df.columns:
        match = DATE.search(str(col))
        if col not in META_COLUMNS and match:
            dated[col] = pd.Timestamp(match.group(1))
    if annual:
        full_year = [c for c in dated if "(FY)" in str(c)]
        part_year = [c for c in dated if "(Q" in str(c) or "(YTD)" in str(c)]
        chosen = full_year or [c for c in dated if c not in part_year]
    else:
        chosen = list(dated)
    chosen.sort(key=lambda c: dated[c], reverse=True)

    out = pd.DataFrame({
        "concept": df["concept"].astype(str),
        "label": df["label"].astype(str),
    })
    out["local"] = out["concept"].map(local_name)
    for extra in ("balance", "preferred_sign", "parent_concept"):
        out[extra] = df[extra] if extra in df.columns else None
    periods = []
    for col in chosen:
        end = dated[col]
        if end not in periods:
            out[end] = pd.to_numeric(df[col], errors="coerce")
            periods.append(end)
    return Table(out, periods) if periods else None


def find(table, concepts=(), labels=(), skip_labels=None):
    """Look up one line item.

    Tries the XBRL concept names in order (exact match), then the label
    patterns. Returns (values by period, where it came from) or (None, None).
    """
    if table is None:
        return None, None
    df = table.df
    candidates = [(df[df["local"] == name.lower()], False) for name in concepts]
    for pattern in labels:
        rows = df[df["label"].str.contains(pattern, case=False, regex=True, na=False)]
        if skip_labels:
            rows = rows[~rows["label"].str.contains(skip_labels, case=False, regex=True, na=False)]
        candidates.append((rows, True))
    for rows, by_label in candidates:
        for _, row in rows.iterrows():
            values = pd.to_numeric(row[table.periods], errors="coerce")
            if values.notna().any():
                source = row["concept"]
                if by_label:
                    source += f' (matched on the label "{row["label"]}")'
                return values, source
    return None, None


def from_getter(getter, periods):
    """Fallback lookup using EdgarTools' own helpers (0 = latest year, 1 = the year before...)."""
    values = []
    for offset in range(len(periods)):
        try:
            values.append(getter(offset))
        except Exception:
            values.append(None)
    series = pd.to_numeric(pd.Series(values, index=periods), errors="coerce")
    return series if series.notna().any() else None


def latest_total(table, concepts):
    """Add up every one of these concepts found on the latest balance sheet."""
    total, parts = 0.0, []
    for name in concepts:
        values, source = find(table, [name])
        if values is not None and pd.notna(values.iloc[0]):
            total += float(values.iloc[0])
            parts.append(source)
    return (total, " + ".join(parts)) if parts else (None, None)


def working_capital_change(table):
    """Cash effect of the changes in operating assets and liabilities.

    Negative = cash used (for example receivables or inventory went up).
    The filed values say "increase/(decrease) in X", so each line needs a sign:
    in XBRL a debit-balance cash flow item adds to cash and a credit-balance
    item uses cash.
    """
    if table is None:
        return None, None
    df = table.df
    rows = df[df["local"].str.startswith("increasedecreasein") & ~df["local"].str.contains("cash")]
    rows = rows[rows[table.periods].notna().any(axis=1)]
    if rows.empty:
        return None, None

    # If a subtotal and its components are both listed, count the components only.
    parents = set(rows["parent_concept"].dropna().astype(str))
    if parents:
        components = rows[~rows["concept"].isin(parents)]
    else:
        components = rows[rows["local"] != "increasedecreaseinoperatingcapital"]
        if len(components) < 3:  # a lone subtotal (plus the odd extra line): keep everything
            components = rows
    if not components.empty:
        rows = components

    def cash_sign(row):
        balance = str(row["balance"]).lower()
        if balance == "debit":
            return 1.0
        if balance == "credit":
            return -1.0
        if row["preferred_sign"] in (1, -1):  # how the company displays the line
            return float(row["preferred_sign"])
        return -1.0 if re.search(r"receivable|inventor|prepaid|asset", row["local"]) else 1.0

    signs = rows.apply(cash_sign, axis=1)
    total = rows[table.periods].mul(signs, axis=0).sum(axis=0, min_count=1)

    # Cross-check against the signs the company uses when it displays the statement.
    shown = rows["preferred_sign"].map(lambda v: -1.0 if v == -1 else 1.0)
    as_shown = rows[table.periods].mul(shown, axis=0).sum(axis=0, min_count=1)
    source = f"sum of {len(rows)} 'changes in operating assets and liabilities' lines"
    gap = (total - as_shown).abs().max()
    if pd.notna(gap) and gap > 0.005 * max(total.abs().max(), 1.0):
        source += " (CHECK THE SIGN against the cash flow statement)"
    return total, source


def set_sec_identity():
    """Tell the SEC who is asking, without ever storing it in the code."""
    from edgar import set_identity

    identity = os.environ.get("EDGAR_IDENTITY", "").strip()
    if not identity:
        print("The SEC asks everyone who downloads filings to give a name and email.")
        identity = input("Your name and email (e.g. Jane Doe jane@example.com): ").strip()
        if not identity:
            sys.exit("Nothing entered, so the SEC download cannot start. Run the script again.")
    set_identity(identity)


def get_inputs(ticker):
    """Download the latest annual statements and return the numbers the model needs ($ millions)."""
    try:
        from edgar import Company
    except ImportError:
        sys.exit("EdgarTools is not installed in this environment. Run:  python -m pip install edgartools")

    set_sec_identity()
    print(f"Downloading the latest annual report for {ticker} from SEC EDGAR...")
    try:
        company = Company(ticker)
    except Exception as error:
        if type(error).__name__ == "CompanyNotFoundError":
            sys.exit(f"SEC EDGAR has no company with the ticker {ticker}. Check the spelling.")
        raise
    financials = company.get_financials()
    if financials is None:
        sys.exit(f"No annual report found for {ticker}. Check the ticker.")

    income = to_table(financials.income_statement())
    cashflow = to_table(financials.cash_flow_statement())
    balance = to_table(financials.balance_sheet(), annual=False)
    if income is None:
        sys.exit(f"The latest annual report for {ticker} has no income statement data that can be read.")

    years = income.periods[:3]  # fiscal year-ends, newest first

    def aligned(values):
        """Line the values up with the income statement's fiscal years."""
        if values is None:
            return None
        by_date = values.reindex(years)
        if by_date.notna().any():
            return by_date
        if len(values) >= len(years):  # same years, differently dated: go by position
            return pd.Series(list(values.iloc[:len(years)]), index=years)
        return None

    def lookup(table, concepts, labels=(), skip_labels=None, getter=None):
        values, source = find(table, concepts, labels, skip_labels)
        if values is None and getter is not None:
            values = from_getter(getter, years)
            source = f"EdgarTools {getter.__name__}()" if values is not None else None
        return aligned(values), source

    # ---- income statement ----
    revenue, revenue_src = lookup(
        income,
        ["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax",
         "RevenueFromContractWithCustomerIncludingAssessedTax", "SalesRevenueNet",
         "RevenuesNetOfInterestExpense", "Revenue", "SalesRevenueGoodsNet"],
        labels=[r"^(?:total )?(?:net )?(?:revenues?|sales)$", r"^total (?:net )?(?:revenues?|sales)"],
        getter=financials.get_revenue)
    cost, cost_src = lookup(
        income,
        ["CostOfGoodsAndServicesSold", "CostOfRevenue", "CostOfGoodsSold", "CostOfSales",
         "CostOfGoodsAndServiceExcludingDepreciationDepletionAndAmortization"],
        labels=[r"^(?:total )?cost of (?:revenues?|sales|goods sold)"])
    pretax, pretax_src = lookup(
        income,
        ["IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
         "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments",
         "IncomeLossBeforeIncomeTaxes", "IncomeLossFromContinuingOperationsBeforeIncomeTaxes"],
        labels=[r"(?:income|earnings|loss|profit).*before.*tax"])
    tax, tax_src = lookup(
        income,
        ["IncomeTaxExpenseBenefit", "IncomeTaxExpenseBenefitContinuingOperations"],
        labels=[r"^(?:provision for |benefit from )?income tax"], skip_labels=r"deferred|paid|payable")
    net_income, net_income_src = lookup(
        income,
        ["NetIncomeLoss", "ProfitLoss", "NetIncomeLossAvailableToCommonStockholdersBasic"],
        labels=[r"^net (?:income|earnings|loss)"], skip_labels=r"noncontrolling|per share",
        getter=financials.get_net_income)
    shares, shares_src = lookup(
        income,
        ["WeightedAverageNumberOfDilutedSharesOutstanding", "WeightedAverageNumberOfSharesOutstandingBasic"],
        getter=financials.get_shares_outstanding_diluted)

    # Operating income (EBIT). Not every company reports this line, so fall back
    # to pre-tax income with interest expense added back.
    ebit, ebit_src = lookup(
        income,
        ["OperatingIncomeLoss", "IncomeLossFromContinuingOperationsBeforeInterestAndTaxes"],
        labels=[r"^operating (?:income|profit|loss)", r"^(?:income|loss|profit).*from operations$"])
    if ebit is None and pretax is not None:
        interest, interest_src = lookup(
            income,
            ["InterestExpense", "InterestExpenseNonoperating", "InterestAndDebtExpense", "InterestExpenseDebt"],
            labels=[r"^interest expense"])
        if interest is not None:
            ebit = pretax + interest.fillna(0).abs()
            ebit_src = f"DERIVED (no operating income line filed): {pretax_src} + {interest_src}"
        else:
            ebit = pretax
            ebit_src = f"DERIVED (no operating income or interest line filed): {pretax_src}"

    # ---- cash flow statement ----
    cfo, cfo_src = lookup(
        cashflow,
        ["NetCashProvidedByUsedInOperatingActivities",
         "NetCashProvidedByUsedInOperatingActivitiesContinuingOperations"],
        labels=[r"^net cash.*operating activities"],
        getter=financials.get_operating_cash_flow)
    da, da_src = lookup(
        cashflow,
        ["DepreciationDepletionAndAmortization", "DepreciationAndAmortization",
         "DepreciationAmortizationAndAccretionNet", "DepreciationAmortizationAndOther", "Depreciation"],
        labels=[r"^depreciation"])
    sbc, sbc_src = lookup(
        cashflow,
        ["ShareBasedCompensation", "AllocatedShareBasedCompensationExpense"],
        labels=[r"(?:share|stock)[- ]based compensation"])
    capex, capex_src = lookup(
        cashflow,
        ["PaymentsToAcquirePropertyPlantAndEquipment", "PaymentsToAcquireProductiveAssets"],
        labels=[r"capital expenditures", r"(?:purchases?|acquisitions?|additions?) (?:of|to) property",
                r"payments for.*property"],
        getter=financials.get_capital_expenditures)
    if capex is not None:
        capex = capex.abs()  # always shown as a positive amount spent
    wc, wc_src = working_capital_change(cashflow)
    wc = aligned(wc)

    def millions(values):
        """Series (newest first) -> list of $ millions, oldest first, None where missing."""
        if values is None:
            return [None] * len(years)
        return [None if pd.isna(v) else float(v) / 1e6 for v in reversed(list(values))]

    if revenue is None or pd.isna(revenue.iloc[0]):
        sys.exit(f"Could not find revenue in the latest annual report for {ticker}, so no model was built.")

    # ---- balance sheet (latest date) ----
    def bs_first(concepts, labels=(), skip_labels=None):
        values, source = find(balance, concepts, labels, skip_labels)
        if values is None or pd.isna(values.iloc[0]):
            return None, None
        return float(values.iloc[0]) / 1e6, source

    def bs_sum(concepts):
        total, source = latest_total(balance, concepts)
        return (None, None) if total is None else (total / 1e6, source)

    cash = bs_first(
        ["CashAndCashEquivalentsAtCarryingValue",
         "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
         "Cash", "CashCashEquivalentsAndShortTermInvestments"],
        labels=[r"^cash and cash equivalents"])
    short_investments = bs_first(
        ["MarketableSecuritiesCurrent", "AvailableForSaleSecuritiesDebtSecuritiesCurrent",
         "ShortTermInvestments"],
        labels=[r"^(?:short-term|current) (?:investments|marketable securities)"])
    long_investments = bs_first(
        ["MarketableSecuritiesNoncurrent", "AvailableForSaleSecuritiesDebtSecuritiesNoncurrent"])

    short_debt = bs_first(["DebtCurrent"])
    if short_debt[0] is None:
        short_debt = bs_sum(
            ["CommercialPaper", "ShortTermBorrowings", "OtherShortTermBorrowings", "LongTermDebtCurrent",
             "LongTermDebtAndCapitalLeaseObligationsCurrent", "NotesPayableCurrent",
             "ConvertibleNotesPayableCurrent"])
    if short_debt[0] is None:
        short_debt = bs_first(
            [], labels=[r"commercial paper", r"short-term (?:debt|borrowings)",
                        r"current (?:portion|maturities) of (?:long-term )?debt"])
    long_debt = bs_first(
        ["LongTermDebtNoncurrent", "LongTermDebtAndCapitalLeaseObligations", "LongTermDebt",
         "LongTermNotesPayable", "SeniorNotes"],
        labels=[r"long-term debt", r"^term debt"], skip_labels=r"current (?:portion|maturities)")
    convertible = bs_sum(["ConvertibleNotesPayableNoncurrent", "ConvertibleDebtNoncurrent"])
    if convertible[0] is not None:
        long_debt = ((long_debt[0] or 0.0) + convertible[0],
                     " + ".join(s for s in (long_debt[1], convertible[1]) if s))

    return {
        "ticker": ticker,
        "name": str(getattr(company, "name", ticker)),
        "cik": getattr(company, "cik", None),
        "years": list(reversed(years)),  # oldest first
        "rows": {
            "revenue": (millions(revenue), revenue_src),
            "cost": (millions(cost), cost_src),
            "ebit": (millions(ebit), ebit_src),
            "pretax": (millions(pretax), pretax_src),
            "tax": (millions(tax), tax_src),
            "net_income": (millions(net_income), net_income_src),
            "shares": (millions(shares), shares_src),
            "da": (millions(da), da_src),
            "sbc": (millions(sbc), sbc_src),
            "wc": (millions(wc), wc_src),
            "capex": (millions(capex), capex_src),
            "cfo": (millions(cfo), cfo_src),
        },
        "balance_date": balance.periods[0] if balance is not None else None,
        "balance": {
            "cash": cash, "short_investments": short_investments, "long_investments": long_investments,
            "short_debt": short_debt, "long_debt": long_debt,
        },
    }


def print_inputs(d):
    """Show what was found, so it can be checked against the 10-K."""
    labels = {
        "revenue": "Revenue", "cost": "Cost of revenue", "ebit": "Operating income (EBIT)",
        "pretax": "Pre-tax income", "tax": "Income tax", "net_income": "Net income",
        "shares": "Diluted shares (mm)", "da": "Depreciation & amortization",
        "sbc": "Share-based compensation", "wc": "Change in working capital",
        "capex": "Capital expenditures", "cfo": "Cash from operations",
    }
    years = [f"FY{y.year}" for y in d["years"]]
    print(f"\n{d['name']} ({d['ticker']}), $ millions")
    print(f"{'':30}" + "".join(f"{y:>14}" for y in years))
    for key, label in labels.items():
        values, source = d["rows"][key]
        cells = "".join(f"{'missing':>14}" if v is None else f"{v:>14,.0f}" for v in values)
        print(f"{label:30}{cells}   {source or 'NOT FOUND'}")
    when = d["balance_date"].date() if d["balance_date"] is not None else "n/a"
    print(f"\nBalance sheet at {when}")
    for key, label in [("cash", "Cash and equivalents"), ("short_investments", "Short-term investments"),
                       ("long_investments", "Long-term securities"), ("short_debt", "Short-term debt"),
                       ("long_debt", "Long-term debt")]:
        value, source = d["balance"][key]
        cell = f"{'none found':>14}" if value is None else f"{value:>14,.0f}"
        print(f"{label:30}{cell}   {source or ''}")


# =============================================================================
# Part 2: starting assumptions for the forecast (taken from the company's history)
# =============================================================================

def clamp(value, low, high):
    return max(low, min(high, value))


def default_drivers(d):
    """First-pass forecast inputs. They are placeholders for the yellow cells, nothing more."""
    get = lambda key: d["rows"][key][0]
    revenue = get("revenue")

    def ratios(values, negate=False):
        out = []
        for v, r in zip(values, revenue):
            if v is not None and r:
                out.append((-v if negate else v) / r)
        return out

    growth = [revenue[i] / revenue[i - 1] - 1 for i in range(1, len(revenue))
              if revenue[i] is not None and revenue[i - 1]]
    start = clamp(sum(growth) / len(growth), 0.0, 0.15) if growth else 0.05
    end = min(start, TERMINAL_GROWTH + 0.02)
    growth_path = [round(start + (end - start) * i / (N_FORECAST - 1), 3) for i in range(N_FORECAST)]

    margins = ratios(get("ebit"))
    margin = round(margins[-1], 3) if margins else 0.15

    tax_rate = 0.21
    tax, pretax = get("tax")[-1], get("pretax")[-1]
    if tax is not None and pretax and 0.05 <= tax / pretax <= 0.35:
        tax_rate = round(tax / pretax, 3)

    da = ratios(get("da"))
    da_pct = round(da[-1], 3) if da else 0.03
    capex = ratios(get("capex"))
    capex_pct = round(sum(capex) / len(capex), 3) if capex else da_pct
    nwc = ratios(get("wc"), negate=True)
    nwc_pct = round(clamp(sum(nwc) / len(nwc), 0.0, 0.05), 3) if nwc else 0.01

    return {
        "growth": (growth_path,
                   "Starts at the average historical growth (capped at 0-15%) and fades toward "
                   "terminal growth + 2 points."),
        "margin": ([margin] * N_FORECAST, "Latest year's EBIT margin, held flat."),
        "tax": ([tax_rate] * N_FORECAST,
                "Latest year's effective tax rate (21% is used if that falls outside 5-35%)."),
        "da": ([da_pct] * N_FORECAST, "Latest year's D&A as a share of revenue."),
        "capex": ([capex_pct] * N_FORECAST, "Average historical capex as a share of revenue."),
        "nwc": ([nwc_pct] * N_FORECAST,
                "Average historical working-capital investment as a share of revenue (kept within 0-5%)."),
    }


# =============================================================================
# Part 3: write the Excel workbook
# =============================================================================

FONT = "Arial"
BLUE, BLACK, GREEN, GREY = "0000FF", "000000", "008000", "555555"
YELLOW = PatternFill("solid", fgColor="FFFF00")
HEADER = PatternFill("solid", fgColor="1F3864")
SUBHEAD = PatternFill("solid", fgColor="D9E1F2")
MM = "#,##0;(#,##0);-"                # $ millions
PCT = "0.0%;(0.0%);-"
PCT2 = "0.00%"
PRICE = "$#,##0.00;($#,##0.00);-"
MULTIPLE = '0.0"x"'
THIN = Side(style="thin", color="999999")

# Row numbers that formulas on other sheets point at.
H_REV, H_COST, H_GP, H_EBIT, H_PRETAX, H_TAX, H_NI, H_SHARES = 5, 6, 7, 8, 9, 10, 11, 12
H_DA, H_SBC, H_WC, H_CAPEX, H_CFO, H_FCF, H_OTHER = 15, 16, 17, 18, 19, 20, 21
H_GROWTH, H_GM, H_MARGIN, H_TAXRATE, H_DAPCT, H_CAPEXPCT, H_NWCPCT, H_SBCPCT = 24, 25, 26, 27, 28, 29, 30, 31
A_PRICE, A_SHARES, A_MCAP = 9, 10, 11
A_CASH, A_DEBT, A_NETDEBT = 17, 21, 22
A_WACC, A_G, A_STEP_WACC, A_STEP_G = 36, 39, 40, 41


def put(ws, ref, value, kind="calc", fmt=None, bold=False, fill=None, note=None, align=None):
    """Write one cell. kind sets the colour: input = blue, calc = black, link = green."""
    cell = ws[ref]
    cell.value = value
    colour = {"input": BLUE, "calc": BLACK, "link": GREEN, "label": BLACK}[kind]
    cell.font = Font(name=FONT, size=10, color=colour, bold=bold)
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    if note:
        cell.comment = Comment(note, "build_dcf.py")
    if align:
        cell.alignment = Alignment(horizontal=align)
    return cell


def small(ws, ref, text):
    """Grey explanatory text."""
    ws[ref].value = text
    ws[ref].font = Font(name=FONT, size=9, italic=True, color=GREY)
    ws[ref].alignment = Alignment(indent=1)


def title(ws, text, subtitle):
    ws["A1"].value = text
    ws["A1"].font = Font(name=FONT, size=14, bold=True)
    small(ws, "A2", subtitle)


def section(ws, row, text, ncols):
    for i in range(1, ncols + 1):
        cell = ws.cell(row=row, column=i)
        cell.fill = HEADER
        cell.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
    ws.cell(row=row, column=1).value = text


def build_workbook(d, path, share_price=None):
    n = len(d["years"])                                           # historical years
    H = [get_column_letter(2 + i) for i in range(n)]              # historical columns
    F = [get_column_letter(2 + n + i) for i in range(N_FORECAST)]  # forecast columns (DCF sheet)
    SRC = get_column_letter(2 + n)                                # source column (Historicals)
    last, f1, fN = H[-1], F[0], F[-1]
    fy = [f"FY{y.year}" for y in d["years"]]
    fy_forecast = [f"FY{d['years'][-1].year + i + 1}" for i in range(N_FORECAST)]
    drivers = default_drivers(d)
    name, ticker = d["name"], d["ticker"]
    year_end = d["years"][-1].date()

    wb = Workbook()

    # ------------------------------------------------------------------ README
    rd = wb.active
    rd.title = "README"
    title(rd, f"Discounted Cash Flow (DCF) Model: {name} ({ticker})",
          "Built by build_dcf.py from SEC EDGAR data. Educational template, not investment advice.")
    cik = d.get("cik")
    edgar_link = (f"https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK={cik}&type=10-K"
                  if cik else "https://www.sec.gov/edgar/search/")
    lines = [
        ("HOW TO USE", True),
        ("1. Yellow cells on 'Assumptions' and 'DCF' are PLACEHOLDERS, not recommendations. Replace them with your own views.", False),
        ("2. Enter the current share price on 'Assumptions' if it is blank.", False),
        ("3. Check 'Historicals' against the 10-K: the last column shows which filed line each number came from.", False),
        ("4. Read the result on 'DCF': value per share, upside/(downside) and the sensitivity table.", False),
        ("", False),
        ("COLOUR CODE", True),
        ("Blue text = input", False),
        ("Black text = formula", False),
        ("Green text = link to another sheet", False),
        ("Yellow fill = placeholder or missing number that needs your attention", False),
        ("", False),
        ("METHOD AND SIMPLIFICATIONS", True),
        ("Unlevered free cash flow = EBIT x (1 - tax rate) + D&A - capex - investment in net working capital.", False),
        ("Share-based compensation is NOT added back (it is treated as a real cost).", False),
        (f"Cash flows are discounted at year-end for {N_FORECAST} years, with no adjustment for the part of the year already gone.", False),
        ("Terminal value uses the Gordon growth formula on the final forecast year.", False),
        ("Debt covers borrowings shown on the balance sheet. Operating leases are not included unless you add them.", False),
        ("Value per share uses the diluted weighted-average share count from the annual report.", False),
        ("", False),
        ("SOURCES", True),
        (f"Financials: {name} annual report (Form 10-K) for the fiscal year ended {year_end}, "
         f"SEC EDGAR XBRL data read with EdgarTools. {edgar_link}", False),
        ("Share price, risk-free rate, equity risk premium, beta and cost of debt: entered by the user in build_dcf.py or in the yellow cells.", False),
        ("Forecast growth, margin, tax, D&A, capex and working capital: starting points taken from the historical ratios (see the cell comments).", False),
    ]
    for i, (text, bold) in enumerate(lines, start=4):
        put(rd, f"A{i}", text, "label", bold=bold)
    rd.column_dimensions["A"].width = 140

    # ------------------------------------------------------------- Historicals
    hs = wb.create_sheet("Historicals")
    title(hs, "Historical Financials ($ millions)",
          f"{name}, fiscal years ending around {year_end:%d %b}. Blue = number from the 10-K. "
          "Yellow = not found automatically, type it in.")
    put(hs, "A4", "($ millions)", "label", bold=True, fill=SUBHEAD)
    for col, label in zip(H, fy):
        put(hs, f"{col}4", label, "label", bold=True, fill=SUBHEAD, align="right")
    put(hs, f"{SRC}4", "Source (line in the filing)", "label", bold=True, fill=SUBHEAD)

    optional = {"cost": "Not reported as a separate line (leave blank)",
                "sbc": "Not found on the cash flow statement"}
    input_rows = [
        (H_REV, "Revenue", "revenue", MM), (H_COST, "Cost of revenue", "cost", MM),
        (H_EBIT, "Operating income (EBIT)", "ebit", MM), (H_PRETAX, "Pre-tax income", "pretax", MM),
        (H_TAX, "Income tax expense", "tax", MM), (H_NI, "Net income", "net_income", MM),
        (H_SHARES, "Diluted shares (millions)", "shares", "#,##0.0"),
        (H_DA, "Depreciation and amortization", "da", MM),
        (H_SBC, "Share-based compensation", "sbc", MM),
        (H_WC, "Change in working capital (cash effect, negative = cash used)", "wc", MM),
        (H_CAPEX, "Capital expenditures", "capex", MM),
        (H_CFO, "Cash from operating activities", "cfo", MM),
    ]
    for row, label, key, fmt in input_rows:
        values, source = d["rows"][key]
        put(hs, f"A{row}", label, "label")
        for col, value in zip(H, values):
            missing = value is None and key not in optional
            put(hs, f"{col}{row}", value, "input", fmt, fill=YELLOW if missing else None)
        small(hs, f"{SRC}{row}", source or optional.get(key, "NOT FOUND: enter it from the 10-K"))

    put(hs, f"A{H_GP}", "Gross profit", "label", bold=True)
    put(hs, f"A{H_FCF}", "Free cash flow (operating cash flow - capex)", "label", bold=True)
    put(hs, f"A{H_OTHER}", "Other operating items, net (reconciling check)", "label")
    small(hs, f"{SRC}{H_OTHER}",
          "Operating cash flow - net income - D&A - SBC - change in working capital. "
          "A very large number here can mean the working-capital sign is wrong.")
    section(hs, 14, "Cash flow items ($ millions)", n + 2)
    section(hs, 23, "Ratios", n + 2)
    ratio_labels = {
        H_GROWTH: "Revenue growth", H_GM: "Gross margin %", H_MARGIN: "Operating (EBIT) margin %",
        H_TAXRATE: "Effective tax rate", H_DAPCT: "D&A % of revenue", H_CAPEXPCT: "Capex % of revenue",
        H_NWCPCT: "Investment in working capital % of revenue", H_SBCPCT: "Share-based compensation % of revenue",
    }
    for row, label in ratio_labels.items():
        put(hs, f"A{row}", label, "label")
    for i, c in enumerate(H):
        put(hs, f"{c}{H_GP}", f'=IF(ISNUMBER({c}{H_COST}),{c}{H_REV}-{c}{H_COST},"n/a")', "calc", MM, bold=True, align="right")
        put(hs, f"{c}{H_FCF}", f"={c}{H_CFO}-{c}{H_CAPEX}", "calc", MM, bold=True)
        put(hs, f"{c}{H_OTHER}", f"={c}{H_CFO}-{c}{H_NI}-{c}{H_DA}-{c}{H_SBC}-{c}{H_WC}", "calc", MM)
        if i > 0:
            p = H[i - 1]
            put(hs, f"{c}{H_GROWTH}", f'=IF(N({p}{H_REV})=0,"n/a",{c}{H_REV}/{p}{H_REV}-1)', "calc", PCT, align="right")
        per_revenue = {H_MARGIN: f"{c}{H_EBIT}", H_DAPCT: f"{c}{H_DA}", H_CAPEXPCT: f"{c}{H_CAPEX}",
                       H_NWCPCT: f"-{c}{H_WC}", H_SBCPCT: f"{c}{H_SBC}"}
        for row, top in per_revenue.items():
            put(hs, f"{c}{row}", f'=IF(N({c}{H_REV})=0,"n/a",{top}/{c}{H_REV})', "calc", PCT, align="right")
        put(hs, f"{c}{H_GM}", f'=IF(AND(ISNUMBER({c}{H_GP}),N({c}{H_REV})<>0),{c}{H_GP}/{c}{H_REV},"n/a")', "calc", PCT, align="right")
        put(hs, f"{c}{H_TAXRATE}", f'=IF(N({c}{H_PRETAX})=0,"n/a",{c}{H_TAX}/{c}{H_PRETAX})', "calc", PCT, align="right")
    hs.column_dimensions["A"].width = 60
    for c in H:
        hs.column_dimensions[c].width = 14
    hs.column_dimensions[SRC].width = 95
    hs.freeze_panes = "B5"

    # ------------------------------------------------------------- Assumptions
    a = wb.create_sheet("Assumptions")
    title(a, "Assumptions",
          "$ millions unless stated. Blue = input, black = formula, green = link. "
          "Yellow = placeholder: review it before relying on the result.")
    section(a, 4, "Company and market data", 3)
    put(a, "A5", "Company", "label"); put(a, "B5", name, "input", align="right")
    put(a, "A6", "Ticker", "label"); put(a, "B6", ticker, "input", align="right")
    put(a, "A7", "Valuation date", "label"); put(a, "B7", date.today(), "input", "yyyy-mm-dd", align="right")
    put(a, "A8", "Latest fiscal year end", "label"); put(a, "B8", year_end, "input", "yyyy-mm-dd", align="right")
    put(a, f"A{A_PRICE}", "Current share price ($)", "label")
    put(a, f"B{A_PRICE}", share_price, "input", PRICE, fill=YELLOW,
        note="Enter the current share price. It drives market cap, the WACC weights and the upside/(downside).")
    put(a, f"C{A_PRICE}", f'=IF(N(B{A_PRICE})>0,"","<- enter the current share price")', "calc")
    put(a, f"A{A_SHARES}", "Diluted shares outstanding (millions)", "label")
    put(a, f"B{A_SHARES}", f"=Historicals!{last}{H_SHARES}", "link", "#,##0.0")
    put(a, f"A{A_MCAP}", "Market capitalization", "label")
    put(a, f"B{A_MCAP}", f"=N(B{A_PRICE})*B{A_SHARES}", "calc", MM)

    when = d["balance_date"].date() if d["balance_date"] is not None else "n/a"
    section(a, 13, f"Net debt (balance sheet at {when})", 3)
    bal = d["balance"]
    net_debt_rows = [
        (14, "Cash and cash equivalents", "cash", True),
        (15, "Short-term investments", "short_investments", False),
        (16, "Long-term marketable securities", "long_investments", False),
        (18, "Short-term debt (including the current part of long-term debt)", "short_debt", False),
        (19, "Long-term debt", "long_debt", False),
    ]
    for row, label, key, required in net_debt_rows:
        value, source = bal[key]
        put(a, f"A{row}", label, "label")
        put(a, f"B{row}", value, "input", MM, fill=YELLOW if (value is None and required) else None)
        small(a, f"C{row}", source or ("NOT FOUND: enter it from the balance sheet" if required
                                       else "None found on the balance sheet (enter it here if the company has any)"))
    put(a, "A20", "Other debt-like items (optional, e.g. lease liabilities)", "label")
    put(a, "B20", None, "input", MM)
    small(a, "C20", "Left blank by the script. Add anything else you treat as debt.")
    put(a, f"A{A_CASH}", "Total cash and investments", "label", bold=True)
    put(a, f"B{A_CASH}", "=SUM(B14:B16)", "calc", MM, bold=True)
    put(a, f"A{A_DEBT}", "Total debt", "label", bold=True)
    put(a, f"B{A_DEBT}", "=SUM(B18:B20)", "calc", MM, bold=True)
    put(a, f"A{A_NETDEBT}", "Net debt / (net cash)", "label", bold=True)
    put(a, f"B{A_NETDEBT}", f"=B{A_DEBT}-B{A_CASH}", "calc", MM, bold=True)

    section(a, 24, "Discount rate (WACC)", 3)
    wacc_inputs = [
        (25, "Risk-free rate", RISK_FREE_RATE, PCT2, True, "PLACEHOLDER: use the current 10-year US Treasury yield."),
        (26, "Equity risk premium", EQUITY_RISK_PREMIUM, PCT2, True, "PLACEHOLDER: your own view, commonly 4-6%."),
        (27, "Beta", BETA, "0.00", True, "PLACEHOLDER: 1.0 means market risk. Look up the company's beta."),
        (29, "Pre-tax cost of debt", PRE_TAX_COST_OF_DEBT, PCT2, True, "PLACEHOLDER: roughly the yield on the company's long-term bonds."),
        (30, "Tax rate on interest", TAX_RATE_ON_INTEREST, PCT2, False, "US federal statutory rate, for the interest tax shield."),
    ]
    for row, label, value, fmt, placeholder, note in wacc_inputs:
        put(a, f"A{row}", label, "label")
        put(a, f"B{row}", value, "input", fmt, fill=YELLOW if placeholder else None, note=note)
    put(a, "A28", "Cost of equity (CAPM)", "label"); put(a, "B28", "=B25+B27*B26", "calc", PCT2)
    put(a, "A31", "After-tax cost of debt", "label"); put(a, "B31", "=B29*(1-B30)", "calc", PCT2)
    put(a, "A32", "Equity weight (market cap)", "label")
    put(a, "B32", f"=IF(N(B{A_PRICE})>0,B{A_MCAP}/(B{A_MCAP}+B{A_DEBT}),1)", "calc", PCT)
    put(a, "C32", f'=IF(N(B{A_PRICE})>0,"","Share price is blank, so WACC uses 100% equity until you enter it")', "calc")
    put(a, "A33", "Debt weight", "label"); put(a, "B33", "=1-B32", "calc", PCT)
    put(a, "A34", "WACC (calculated)", "label", bold=True)
    put(a, "B34", "=B32*B28+B33*B31", "calc", PCT2, bold=True)
    put(a, "A35", "WACC override (optional: leave blank to use the calculated one)", "label")
    put(a, "B35", None, "input", PCT2, fill=YELLOW)
    put(a, f"A{A_WACC}", "WACC used in the DCF", "label", bold=True)
    put(a, f"B{A_WACC}", "=IF(ISNUMBER(B35),B35,B34)", "calc", PCT2, bold=True)

    section(a, 38, "Terminal value and sensitivity", 3)
    put(a, f"A{A_G}", "Terminal growth rate", "label")
    put(a, f"B{A_G}", TERMINAL_GROWTH, "input", PCT2, fill=YELLOW,
        note="PLACEHOLDER: keep it at or below long-run nominal GDP growth and below the risk-free rate.")
    put(a, f"A{A_STEP_WACC}", "Sensitivity step: WACC", "label")
    put(a, f"B{A_STEP_WACC}", 0.005, "input", PCT2)
    put(a, f"A{A_STEP_G}", "Sensitivity step: terminal growth", "label")
    put(a, f"B{A_STEP_G}", 0.005, "input", PCT2)
    a.column_dimensions["A"].width = 62
    a.column_dimensions["B"].width = 16
    a.column_dimensions["C"].width = 95
    a.freeze_panes = "A4"

    # --------------------------------------------------------------------- DCF
    s = wb.create_sheet("DCF")
    ncols = 1 + n + N_FORECAST
    title(s, "DCF Valuation",
          f"$ millions unless stated. {fy[0]}-{fy[-1]} are actuals. "
          f"{fy_forecast[0]}-{fy_forecast[-1]} are forecasts driven by the yellow inputs in rows 8-13.")
    put(s, "A4", "($ millions)", "label", bold=True, fill=SUBHEAD)
    for col, label in zip(H + F, [f"{y}A" for y in fy] + [f"{y}E" for y in fy_forecast]):
        put(s, f"{col}4", label, "label", bold=True, fill=SUBHEAD, align="right")
    put(s, "A5", "Projection year", "label")
    for c in F:
        put(s, f"{c}5", f"=COLUMNS(${f1}5:{c}5)", "calc", "0")

    section(s, 7, "Operating assumptions", ncols)
    driver_rows = [
        (8, "Revenue growth", H_GROWTH, "growth"), (9, "EBIT margin", H_MARGIN, "margin"),
        (10, "Tax rate on EBIT", H_TAXRATE, "tax"), (11, "D&A % of revenue", H_DAPCT, "da"),
        (12, "Capex % of revenue", H_CAPEXPCT, "capex"), (13, "Investment in working capital % of revenue", H_NWCPCT, "nwc"),
    ]
    for row, label, hist_row, key in driver_rows:
        values, basis = drivers[key]
        put(s, f"A{row}", label, "label")
        for i, c in enumerate(H):
            if not (row == 8 and i == 0):  # no growth rate for the first year
                put(s, f"{c}{row}", f"=Historicals!{c}{hist_row}", "link", PCT, align="right")
        for c, value in zip(F, values):
            put(s, f"{c}{row}", value, "input", PCT, fill=YELLOW)
        s[f"{f1}{row}"].comment = Comment(f"PLACEHOLDER. {basis} Replace with your own forecast.", "build_dcf.py")

    section(s, 15, "Free cash flow build", ncols)
    fcf_labels = {
        16: "Revenue", 17: "EBIT (operating income)", 18: "Less: taxes on EBIT", 19: "NOPAT",
        20: "Plus: depreciation and amortization", 21: "Less: capital expenditures",
        22: "Less: investment in net working capital", 23: "Unlevered free cash flow", 24: "FCF margin",
    }
    for row, label in fcf_labels.items():
        put(s, f"A{row}", label, "label", bold=row in (16, 19, 23))
    for c in H:
        put(s, f"{c}16", f"=Historicals!{c}{H_REV}", "link", MM, bold=True)
        put(s, f"{c}17", f"=Historicals!{c}{H_EBIT}", "link", MM)
        put(s, f"{c}18", f"=-{c}17*N({c}10)", "calc", MM)
        put(s, f"{c}20", f"=Historicals!{c}{H_DA}", "link", MM)
        put(s, f"{c}21", f"=-Historicals!{c}{H_CAPEX}", "link", MM)
        put(s, f"{c}22", f"=Historicals!{c}{H_WC}", "link", MM)
    for i, c in enumerate(F):
        p = last if i == 0 else F[i - 1]
        put(s, f"{c}16", f"={p}16*(1+{c}8)", "calc", MM, bold=True)
        put(s, f"{c}17", f"={c}16*{c}9", "calc", MM)
        put(s, f"{c}18", f"=-{c}17*{c}10", "calc", MM)
        put(s, f"{c}20", f"={c}16*{c}11", "calc", MM)
        put(s, f"{c}21", f"=-{c}16*{c}12", "calc", MM)
        put(s, f"{c}22", f"=-{c}16*{c}13", "calc", MM)
    for c in H + F:
        put(s, f"{c}19", f"={c}17+{c}18", "calc", MM, bold=True)
        put(s, f"{c}23", f"={c}19+{c}20+{c}21+{c}22", "calc", MM, bold=True)
        put(s, f"{c}24", f'=IF(N({c}16)=0,"n/a",{c}23/{c}16)', "calc", PCT, align="right")
        s[f"{c}23"].border = Border(top=THIN)

    section(s, 26, "Discounting", ncols)
    put(s, "A27", "Discount factor", "label")
    put(s, "A28", "Present value of free cash flow", "label", bold=True)
    wacc, g = f"Assumptions!$B${A_WACC}", f"Assumptions!$B${A_G}"
    for c in F:
        put(s, f"{c}27", f"=1/(1+{wacc})^{c}5", "calc", "0.000")
        put(s, f"{c}28", f"={c}23*{c}27", "calc", MM, bold=True)

    section(s, 30, "Valuation", ncols)
    price = f"Assumptions!$B${A_PRICE}"
    valuation = [
        (31, f"Sum of PV of free cash flow ({fy_forecast[0]}-{fy_forecast[-1]})", f"=SUM({f1}28:{fN}28)", MM, "calc"),
        (32, f"Terminal-year free cash flow ({fy_forecast[-1]})", f"={fN}23", MM, "calc"),
        (33, "Terminal value (Gordon growth)", f"=IF({wacc}>{g},{fN}23*(1+{g})/({wacc}-{g}),0)", MM, "link"),
        (34, "PV of terminal value", f"=B33*{fN}27", MM, "calc"),
        (35, "Enterprise value", "=B31+B34", MM, "calc"),
        (36, "Less: total debt", f"=-Assumptions!B{A_DEBT}", MM, "link"),
        (37, "Plus: cash and investments", f"=Assumptions!B{A_CASH}", MM, "link"),
        (38, "Equity value", "=B35+B36+B37", MM, "calc"),
        (39, "Diluted shares (millions)", f"=Assumptions!B{A_SHARES}", "#,##0.0", "link"),
        (40, "Intrinsic value per share ($)", '=IF(N(B39)=0,"n/a",B38/B39)', PRICE, "calc"),
        (41, "Current share price ($)", f'=IF(N({price})>0,{price},"enter on Assumptions")', PRICE, "link"),
        (42, "Upside / (downside) vs current price", f'=IF(AND(ISNUMBER(B40),N({price})>0),B40/{price}-1,"n/a")', PCT, "calc"),
        (43, "Terminal value as % of enterprise value", '=IF(B35=0,"n/a",B34/B35)', PCT, "calc"),
        (44, "Implied terminal EV / EBITDA", f'=IF(({fN}17+{fN}20)=0,"n/a",B33/({fN}17+{fN}20))', MULTIPLE, "calc"),
        (45, "Check: WACC must exceed terminal growth", f'=IF({wacc}>{g},"OK","ERROR: WACC <= terminal growth")', None, "calc"),
    ]
    for row, label, formula, fmt, kind in valuation:
        put(s, f"A{row}", label, "label", bold=row in (35, 38, 40))
        put(s, f"B{row}", formula, kind, fmt, bold=row in (35, 38, 40), align="right")
    s["B40"].fill = YELLOW
    small(s, "C40", "Result (a formula). Highlighted so it is easy to find.")

    section(s, 47, "Sensitivity: value per share ($). WACC down the side, terminal growth across the top", ncols)
    step_w, step_g = f"Assumptions!$B${A_STEP_WACC}", f"Assumptions!$B${A_STEP_G}"
    put(s, "A48", "WACC \\ terminal growth", "label", bold=True, fill=SUBHEAD)
    put(s, "D48", f"=Assumptions!B{A_G}", "link", PCT2, bold=True, fill=SUBHEAD)
    put(s, "C48", f"=D48-{step_g}", "calc", PCT2, bold=True, fill=SUBHEAD)
    put(s, "B48", f"=C48-{step_g}", "calc", PCT2, bold=True, fill=SUBHEAD)
    put(s, "E48", f"=D48+{step_g}", "calc", PCT2, bold=True, fill=SUBHEAD)
    put(s, "F48", f"=E48+{step_g}", "calc", PCT2, bold=True, fill=SUBHEAD)
    put(s, "A51", f"=Assumptions!B{A_WACC}", "link", PCT2, bold=True, fill=SUBHEAD, align="right")
    put(s, "A50", f"=A51-{step_w}", "calc", PCT2, bold=True, fill=SUBHEAD, align="right")
    put(s, "A49", f"=A50-{step_w}", "calc", PCT2, bold=True, fill=SUBHEAD, align="right")
    put(s, "A52", f"=A51+{step_w}", "calc", PCT2, bold=True, fill=SUBHEAD, align="right")
    put(s, "A53", f"=A52+{step_w}", "calc", PCT2, bold=True, fill=SUBHEAD, align="right")
    fcf, years = f"${f1}$23:${fN}$23", f"${f1}$5:${fN}$5"
    shares = f"Assumptions!$B${A_SHARES}"
    for row in range(49, 54):
        for c in "BCDEF":
            formula = (
                f'=IF(AND($A{row}>{c}$48,N({shares})>0),'
                f'(SUMPRODUCT({fcf}/(1+$A{row})^{years})'
                f'+${fN}$23*(1+{c}$48)/($A{row}-{c}$48)/(1+$A{row})^${fN}$5'
                f'-Assumptions!$B${A_DEBT}+Assumptions!$B${A_CASH})/{shares},"n/a")'
            )
            put(s, f"{c}{row}", formula, "link", PRICE, align="right")
    s["D51"].fill = YELLOW
    small(s, "A55", "The highlighted centre cell equals the value per share above. Everything recalculates when you change an input.")

    s.column_dimensions["A"].width = 48
    for c in H + F:
        s.column_dimensions[c].width = 13
    s.freeze_panes = "B5"

    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
    wb.calculation.fullCalcOnLoad = True  # make Excel calculate every formula when the file opens
    wb.save(path)


# =============================================================================
# Run it
# =============================================================================

def main():
    warnings.filterwarnings("ignore", category=DeprecationWarning)
    ticker = (sys.argv[1] if len(sys.argv) > 1 else TICKER).upper()
    share_price = float(sys.argv[2]) if len(sys.argv) > 2 else SHARE_PRICE

    data = get_inputs(ticker)
    print_inputs(data)

    path = os.path.abspath(f"{ticker}_DCF_Model.xlsx")
    try:
        build_workbook(data, path, share_price)
    except PermissionError:
        sys.exit(f"\nCould not save {path}. It is probably open in Excel: close it and run again.")
    print(f"\nSaved to {path}")
    print("Open it in Excel. Yellow cells are placeholders: replace them with your own assumptions.")
    ebit, revenue = data["rows"]["ebit"][0][-1], data["rows"]["revenue"][0][-1]
    if ebit is not None and revenue and ebit / revenue <= 0:
        print("Note: the latest operating margin is negative, so the starting forecast gives a negative value. "
              "Set your own EBIT margins on the DCF sheet.")
    if share_price is None:
        print("Enter the current share price on the Assumptions sheet to see the upside/(downside).")


if __name__ == "__main__":
    main()
