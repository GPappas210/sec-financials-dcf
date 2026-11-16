# SEC Financials 

Python tools that pull company financials from SEC EDGAR and builds a sample Excel DCF.

## What it does

- `ImportFinDoc/build_dcf.py` downloads a company's latest 10-K and builds an Excel DCF model.
- `ImportFinDoc/ImportQfiling.py` exports the statements from a company's last eight 10-Q filings to Excel. Income and cash flow figures are year-to-date, as filed.

- Export the quarterly statements only:
- `python ImportFinDoc/ImportQfiling.py` Set `TICKER` at the top of the file first. The export is saved as `<TICKER>_quarterly_from_10Q.xlsx`.

## Usage
```
python ImportFinDoc/build_dcf.py            # pulls the financial documents of set ticker
python ImportFinDoc/build_dcf.py MSFT 410   # ticker and current share price
```
The model is saved as `<TICKER>_DCF_Model.xlsx` in the current folder.

## Setup 
```
pip install edgartools pandas openpyxl
```
The SEC asks everyone who downloads filings to give a name and email. 
`build_dcf.py` asks for it when it runs.




## Notes

- Raw financial data is messy. First Excel export was full of segment rows,
  metadata columns and numbers in scientific notation. Most of the work in this project
  was cleaning, not calculating.

- "Quarterly" numbers aren't always quarterly. I thought I had eight quarters of
  revenue until I noticed the figures kept growing through the year. 10-Q income and
  cash flow figures are cumulative, so you have to subtract to get a single quarter,
  and Q4 doesn't exist as a filing at all. It has to be backed out of the annual report.

- A DCF is only as good as its assumptions. With placeholder inputs, my first model
  valued Apple at less than half its share price.
  It showed me how much the result depends on the discount rate and terminal growth,
  which is why the model has a sensitivity table.

- Small things break real projects. A script crashing because the output file was
  still open in Excel, or a `.gitignore` that didn't work because it was saved with the
  wrong name.

## Assumptions and limitations
- Growth, margin and discount-rate inputs are placeholders (yellow cells), not forecasts.
- Cash flows are discounted at year-end, and share-based compensation is not added back.
- Built for US companies that file a 10-K.

## Disclaimer
Educational project. 

## Credits
Data from SEC EDGAR, read with [EdgarTools](https://github.com/dgunning/edgartools).