# SEC Fundamental Analysis Toolkit

Python tools for extracting, structuring, and modelling financial statements
directly from SEC EDGAR filings — without relying on paid data vendors.

Built as part of a systematic research workflow targeting risk-adjusted outperformance
versus SPY / QQQ, with an emphasis on point-in-time data integrity and reproducibility.

---

## What this does

Pulls structured financial data from SEC EDGAR XBRL filings and exports it in a
format ready for financial modelling:

- **Income statement, balance sheet, and cash flow** — sourced directly from 10-K
  and 20-F filings, preserving the document's original line-item hierarchy
- **Multi-year time series** — up to 10 years of annual data stitched across filings,
  aligned by XBRL concept
- **USD millions, point-in-time** — values as originally reported in each filing,
  no restated data, no look-ahead

Covers both domestic filers (10-K: AAPL, AMZN, GOOG, INTC) and foreign private
issuers (20-F: ASML, TSM, GFS, TSEM) with automatic form-type detection.

---

## Scripts

### `examples/99_income_statement.py` — Single filing, structured Excel

Pulls the latest 10-K or 20-F for any ticker and exports all three financial
statements to a formatted Excel workbook (one sheet per statement).

```
TICKER = "AAPL"   # change to any ticker
```

Output preserves the filing's own structure — section headers, indentation, and
line-item hierarchy — matching what you would read in the actual SEC document.

```
uv run python examples\99_income_statement.py
```

### `examples/98_multiperiod_financials.py` — Multi-year time series, Excel

Stitches N annual filings into a single structured workbook with one column per
fiscal year. Uses the latest filing as the structural template and maps historical
values by XBRL concept — so the row order and section groupings reflect the
current filing, not an arbitrary sort.

```
TICKER = "AMZN"   # change to any ticker
YEARS  = 5        # number of annual filings to stitch
```

```
uv run python examples\98_multiperiod_financials.py
```

Sample output — income statement (USD millions):

| Label | FY2025 | FY2024 | FY2023 | FY2022 | FY2021 |
|---|---|---|---|---|---|
| Total net sales | 716,924 | 637,959 | 574,785 | 513,983 | 469,822 |
| *Operating expenses:* | | | | | |
| Cost of sales | 356,414 | 326,288 | 304,739 | 288,831 | 272,344 |
| Fulfillment | 109,074 | 98,505 | 90,619 | 84,299 | 75,111 |
| ... | | | | | |
| **Total operating expenses** | **636,949** | **569,366** | **537,933** | **501,735** | **444,943** |
| **Operating income** | **79,975** | **68,593** | **36,852** | **12,248** | **24,879** |
| **Net income** | **77,670** | **59,248** | **30,425** | **(2,722)** | **33,364** |

### `examples/02_qqq_owner_earnings.py` — Owner earnings yield screen

Computes Buffett's owner earnings yield across a basket of mega-cap names:

```
Owner Earnings = Net Income + D&A - CapEx
Owner Earnings Yield = Owner Earnings / Market Cap
```

Uses EDGAR financials for the earnings components and yfinance for current market
cap. Ranks names by yield — a starting filter for "expensive vs cheap" within QQQ.

> **Note on data integrity**: this snapshot mixes today's market cap with last
> reported financials. It illustrates the data plumbing; it is not a
> point-in-time backtestable signal.

---

## Design principles

**Point-in-time discipline** — values come from the as-filed document for each
period. Restated numbers from later filings are not silently substituted. This
matters for backtesting: using revised data as if it were known at the time
inflates historical performance.

**No paid data** — SEC EDGAR is a public dataset. The only credential required
is a User-Agent string identifying you to the SEC server (name + email).

**Structure over raw dumps** — `XBRLS.from_filings()` aligns concepts across
years but loses the filing's presentational hierarchy. These scripts re-attach
structure from the latest filing so the Excel output reads like the actual
document, not a flat spreadsheet.

---

## Setup

Requires Python 3.10+ and [uv](https://github.com/astral-sh/uv).

```powershell
uv venv .venv
.\.venv\Scripts\Activate.ps1
uv pip install -r requirements.txt
copy .env.example .env
# Edit .env: set SEC_USER_AGENT to "Your Name your.email@example.com"
```

Full install guide, common errors, and use-case walkthroughs:
[GETTING_STARTED.md](GETTING_STARTED.md)

---

## Data source and attribution

Financial data sourced from [SEC EDGAR](https://www.sec.gov/edgar/) — a free,
public dataset maintained by the U.S. Securities and Exchange Commission.

Built on top of [edgartools](https://github.com/dgunning/edgartools) by Dwight Gunning
(Apache 2.0). The scripts in this repository are original work; edgartools is
installed as a dependency and its source code is not included here.
