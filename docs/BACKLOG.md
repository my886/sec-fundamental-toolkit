# Backlog — SEC Fundamental Analysis Toolkit

---

## Open work

### Tier 1 — Build next

- [ ] **`97_ratios.py` — Financial ratio time series**
  Compute standard ratios from the already-extracted three statements and export
  a multi-year trend table. Suggested metrics:
  - Profitability: gross margin, operating margin, net margin, ROIC, ROE
  - Leverage: net debt / EBITDA, interest coverage (EBIT / interest expense)
  - Liquidity: current ratio, DSO, DIO, DPO, cash conversion cycle
  - Valuation: FCF yield (use yfinance for market cap, same pattern as `02_qqq_owner_earnings.py`)

  Reuse `_prepare_structured_df()` from `98_multiperiod_financials.py` for raw data.
  Output: single Excel sheet, ratios as rows, fiscal years as columns.

- [ ] **`96_comps.py` — Comparable companies table**
  Run a peer group (list of tickers) and output one Excel sheet with key metrics
  side by side — the most recognisable equity research deliverable.
  Suggested columns per company: revenue, gross margin %, EBIT margin %, ROIC,
  net debt, EV/EBITDA (needs market cap + debt - cash).
  Handle mixed 10-K / 20-F peers automatically (reuse form-detection pattern).
  Output: `data/{PEER_GROUP}_comps.xlsx`

- [x] **`95_quarterly.py` — Quarterly data + LTM** ✓ 2026-04-26
  Use `company.get_quarterly_financials()` to pull 10-Q data and build:
  - LTM revenue and earnings: sum of last 4 pure-quarterly (Qn) periods
  - QoQ and YoY growth rates per line item
  Essential for valuations that need the most recent quarter, not the last annual filing.
  **Limitation**: 6-K filers (SPOT, ASML, TSM) have no XBRL in interim reports — script
  raises a clear error. Use yfinance (`yf.Ticker(ticker).quarterly_financials`) for these.

- [ ] **`95b_quarterly_yf.py` — Quarterly data for foreign issuers (yfinance fallback)**
  Same output format as `95_quarterly.py` but sourced from yfinance instead of EDGAR XBRL.
  Target tickers: SPOT, ASML, TSM, and any 6-K filer without XBRL interim reports.
  Key difference: yfinance returns restated numbers, not point-in-time SEC filings.
  Note that difference clearly in the Excel title row.

### Tier 2 — Meaningful extensions

- [ ] **`94_dcf.py` — DCF scaffolding**
  Historical data as inputs; assumption rows (revenue growth, EBIT margin,
  CapEx/revenue, WACC, terminal growth); output intrinsic value per share +
  sensitivity table (WACC vs terminal growth).
  5-year explicit period + Gordon growth terminal value is sufficient for a first version.

- [ ] **`93_insiders.py` — Insider transaction signal**
  edgartools parses Form 4 directly. Aggregate net buying vs selling by role
  (CEO/CFO weighted higher) over a rolling 90-day window per ticker.
  Academic basis: Seyhun 1992, Lakonishok & Lee 2001.
  Output: ranked table of net insider buying score across a ticker universe.

- [ ] **Point-in-time filing lookup**
  Add `as_of_date` parameter to `98_multiperiod_financials.py`.
  Filter `company.get_filings()` by `filing_date <= as_of_date` before stitching.
  Required for any honest backtest — without it, look-ahead bias is unquantified.

- [ ] **`92_segments.py` — Geographic / product segment breakdown**
  Use `to_dataframe(view="detailed")`, filter `dimension=True` rows to extract
  segment-level revenue and margin.
  Useful for: TSM (by technology node), ASML (by system type), AMZN (AWS vs retail vs ads).

### Tier 3 — Cross-pillar, bigger scope

- [ ] **FRED macro overlay**
  Map company revenue growth against ISM PMI, yield curve (2s10s), or HY credit
  spreads to identify macro regime sensitivity per company.
  Dependency: `E:\github_awesome\macro\` repo (not yet set up).

- [ ] **10-K Risk Factors text-change signal**
  Year-over-year TF-IDF cosine similarity on Risk Factors section.
  Similarity below cohort median predicts negative forward returns.
  Reference: Cohen, Malloy, Nguyen 2020 — "Lazy Prices".
  Use `filing.sections()` or `filing.text()` to extract the section;
  sklearn `TfidfVectorizer` + `cosine_similarity` for scoring.

- [ ] **Earnings transcript / 8-K parsing**
  Extract headline revenue, EPS, and management guidance from 8-K press releases.
  Compare guidance vs actuals over time as a management credibility signal.

---

## Decisions log

### 2026-04-26 — Add comprehensive income; defer statement of equity

| Decision | Choice | Rationale |
|---|---|---|
| Comprehensive income in multi-period | Added as 4th tab in both scripts | Meaningful data (9 rows AMZN, 51 rows TSM); critical for foreign filers where OCI from currency translation is material relative to net income |
| Statement of equity in multi-period | Deferred — single-period only | XBRLS stitching collapses equity rollforward to ~4 rows because beginning/ending balance rows share XBRL concepts with standalone balance sheet items; rollforward structure only makes sense in a single-period view |

---

### 2026-04-26 — Initial build session

| Decision | Choice | Rationale |
|---|---|---|
| Multi-period structure approach | Single-filing template + XBRLS join by `concept` | XBRLS stitched output loses row order and section headers; joining by XBRL concept key onto the latest filing's template preserves the document hierarchy |
| Value scaling | USD millions; detect per-share rows by `max(abs) < 1 000` | Consistent with financial modelling convention; avoids hardcoding concept names which vary across filers |
| Form type detection | Auto-detect 10-K then 20-F at runtime | Foreign private issuers (ASML, TSM, GFS, TSEM) file 20-F; hardcoding breaks silently for these |
| Repo visibility | Public | Portfolio signal for equity analyst transition — combination of financial statement fluency + engineering is differentiated |
| Excel export library | openpyxl directly (not `pd.ExcelWriter`) | Needed for cell-level formatting: indentation, section header colours, number formats per row |



----



## others
  1. Multi-period history via XBRLS (biggest gap in the current script)

  The current script gets 3 years from a single 10-K. For modelling you want 5–10 years. The docs show:

  from edgar.xbrl import XBRLS

  filings = company.get_filings(form="10-K", amendments=False).head(5)
  xbrls = XBRLS.from_filings(filings)
  income = xbrls.statements.income_statement()   # aligned across all 5 filings

  This stitches multiple 10-Ks into a single time-series-aligned statement — far more useful for trend analysis and DCF inputs than a single filing.

  ---
  2. Quarterly data for earnings momentum

  quarterly = company.get_quarterly_financials()
  q_income = quarterly.income_statement()

  Critical for any YoY / QoQ growth signal or earnings surprise model.

  ---
  3. Convenience metric accessors with period_offset

  Instead of parsing the DataFrame manually for a known line item:

  revenue      = financials.get_revenue()           # current period
  revenue_prev = financials.get_revenue(period_offset=1)  # prior period

  # Other ready-made methods:
  # get_net_income(), get_operating_income(), get_free_cash_flow(),
  # get_total_assets(), get_total_liabilities(), get_stockholders_equity(),
  # get_operating_cash_flow(), get_capital_expenditures(),
  # get_current_assets(), get_current_liabilities()

  Useful for quickly building a metrics table across a universe of tickers without wrangling the full DataFrame every time.

  ---
  4. view="detailed" for segment-level data

  The current standard view is right for the Excel display. For modelling, detailed gives you dimensional breakdowns (Products / Services / Geographic
  segments) that standard suppresses:

  df = income.to_dataframe(view="detailed")
  # rows with dimension=True are the segment splits

  ---
  5. Two extra statements the current script skips

  equity = financials.statement_of_equity()      # retained earnings, buybacks, AOCI
  comp   = financials.comprehensive_income()     # OCI items (currency, pension, hedges)

  Both matter for a complete balance-sheet model (e.g. reconciling equity movements, FX exposure).