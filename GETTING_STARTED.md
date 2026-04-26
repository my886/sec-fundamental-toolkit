# edgartools — Starter Project

> **Pillar:** Fundamental Analysis & Financial Modelling
> **Goal fit:** **Core** — gives us free, point-in-time-capable access to every SEC filing (10-K, 10-Q, 8-K, 13F, Form 4 insider trades, S-1, proxy statements). Foundation for any fundamental screen, valuation model, or insider-flow signal applied to SPY/QQQ constituents.
> **Repo:** https://github.com/dgunning/edgartools
> **License:** Apache 2.0
> **Cost:** Free (SEC EDGAR is a public dataset; no API key required, just a User-Agent identifying you).

> **Attribution:** The scripts in this folder are original work built on top of
> [edgartools](https://github.com/dgunning/edgartools) by Dwight Gunning,
> licensed under the [Apache 2.0 License](https://www.apache.org/licenses/LICENSE-2.0).
> edgartools source code is not included here — it is installed as a dependency
> via `requirements.txt`.

## What it does
Pythonic wrapper around the SEC EDGAR full-text and XBRL endpoints. Lets you pull filings by company / form / date, parse XBRL financials into pandas DataFrames, extract insider transactions, fund holdings, and more — without scraping HTML.

## Edge / differentiation
- Active maintenance (Jason Dunning), thousands of stars, frequent releases.
- Cleaner API than `sec-edgar-downloader` and richer XBRL parsing than rolling your own.
- Built-in rate limiting + caching → won't get you banned by SEC's 10 req/sec rule.

## Honest limitations
- XBRL tag names are not 100% standardized across companies — same line item can use different tags. Expect to write per-company adapters for serious models.
- Restated financials: SEC EDGAR shows the as-reported numbers from each filing, but you need to be deliberate about which filing's data you use to avoid look-ahead bias in backtests.
- Rate-limited to SEC's public-API budget; for portfolio-wide pulls plan for caching.

---

## Install (Windows, PowerShell)

All commands run from `E:\github_awesome\fundamental\edgartools\`.

```powershell
# 1. Move into the repo folder
cd E:\github_awesome\fundamental\edgartools

# 2. Create an isolated virtual environment with uv
uv venv .venv

# 3. Activate it
.\.venv\Scripts\Activate.ps1
# If you get an execution-policy error, run once (as admin):
#   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# 4. Install pinned dependencies (includes edgartools, pandas, openpyxl, yfinance)
uv pip install -r requirements.txt

# 5. Configure your SEC User-Agent (EDGAR REQUIRES this)
copy .env.example .env
notepad .env
# Edit SEC_USER_AGENT to:  "Your Name your.email@example.com"
# SEC will block requests that don't identify a contact.
```

## Verify the install (hello world)

```powershell
python examples\01_hello_world.py
```

Expected output:
```
Company: Apple Inc.  |  CIK: 0000320193
Last 3 10-K filings:
  2024-11-01  10-K    0000320193-24-000123
  2023-11-03  10-K    0000320193-23-000106
  2022-10-28  10-K    0000320193-22-000108
OK — edgartools is working.
```

If this runs, the install is good.

---

## Common errors

| Symptom | Cause | Fix |
|---|---|---|
| `EDGAR User-Agent not set` | `.env` missing or `SEC_USER_AGENT` empty | Copy `.env.example` to `.env` and fill in name + email |
| `403 Forbidden` from sec.gov | User-Agent too generic, or rate-limited | Use a real name + email; if rate-limited, wait 60s |
| `KeyError: 'NetIncomeLoss'` in `02_qqq_owner_earnings.py` | XBRL tag name varies by filer | Print `income.index.tolist()` to find the actual tag for that company |
| `ModuleNotFoundError: edgar` | Wrong venv active | `.\.venv\Scripts\Activate.ps1` and re-run |
| `ModuleNotFoundError: openpyxl` | `openpyxl` not installed | `uv pip install -r requirements.txt` (it's listed there) |
| `AttributeError: 'Statement' object has no attribute 'to_excel'` | `get_financials()` returns `Statement` objects, not DataFrames | Call `.to_dataframe()` first: `stmt.to_dataframe().to_excel(writer, ...)` |

---

## Use cases mapped to the goal (beating SPY / QQQ)

### 1. Owner-earnings yield screen on QQQ mega-caps
File: `examples/02_qqq_owner_earnings.py`. Pulls latest 10-K financials and computes Owner Earnings Yield = (Net Income + D&A − CapEx) / Market Cap. A starting screen for "expensive vs cheap" inside QQQ. **Note:** the example is point-in-time-wrong (mixes today's mkt cap with last 10-K) — useful for showing the data plumbing, not a tradeable signal.

### 2. Insider buying signal
Pull Form 4 (insider transactions) for SPY/QQQ constituents, aggregate net insider buying vs selling per name on a rolling 90-day window. Backtest: rebalance monthly into the top-decile insider-buying names; compare to QQQ. Cluster buying by C-suite (CEO/CFO) historically front-runs positive earnings surprises.

### 3. Three core financial statements — console + Excel export
File: `examples/99_income_statement.py`. Pulls Apple's latest 10-K via XBRL, prints all three statements (income, balance sheet, cash flow) to the console, and writes them as separate sheets into `data/AAPL_financials.xlsx`.

Run it:
```powershell
uv run python examples\99_income_statement.py
```

Expected end of output:
```
[  6.5s] Writing Excel workbook to ...\data\AAPL_financials.xlsx...
[  7.3s] Done. Open AAPL_financials.xlsx in Excel - three tabs.
```

Key notes:
- First run downloads and parses the XBRL filing (~10 MB) — takes 30–90 s. Subsequent runs use cache and complete in ~5 s.
- `company.get_financials()` returns a `Financials` object. Calling `.income_statement()`, `.balance_sheet()`, or `.cash_flow_statement()` on it returns a `Statement` object — **not** a pandas DataFrame. Call `.to_dataframe()` on the `Statement` to get a DataFrame for further analysis or export.
- Change `TICKER = "AAPL"` at line 43 to pull any other company.

### 4. 10-K text-change signal
Parse the "Risk Factors" section of consecutive 10-Ks for each name. Year-over-year text similarity (cosine on TF-IDF) below the cohort median has been shown in academic literature (Cohen, Malloy, Nguyen 2020) to predict negative future returns — a short-side filter for an otherwise long-only QQQ momentum book.

---

## What's next
- Cache filings locally to a parquet store so re-runs are free (edgartools has built-in caching but tune the `EDGAR_LOCAL_DATA_DIR` env var).
- Add a point-in-time wrapper that, given an as-of date, returns only filings *available* on that date — required for honest backtests.
