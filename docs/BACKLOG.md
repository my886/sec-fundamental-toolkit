# Backlog — SEC Fundamental Analysis Toolkit

---

## Open work

### Tier 1 — Build next

- [x] **`97_ratios.py` — Financial ratio time series** ✓ 2026-04-26
  Computes 15 ratios (Profitability / Leverage / Liquidity / FCF) across N fiscal years
  using XBRLS stitching; exports styled single-sheet Excel. Tested: AAPL, MU, GOOG, TSM, ASML.
  **Limitations**:
  - FCF yield is wrong for non-USD 20-F filers (FCF in reporting currency vs USD market cap)
  - Some XBRLS concept-join gaps cause N/A for years where filers changed XBRL tags (GOOG FY2022-24, TSM FY2019-20)
  - Interest coverage N/A for companies that net interest income/expense (AAPL, GOOG)

- [x] **`96_comps.py` — Comparable companies table** ✓ 2026-04-28
  Run a peer group (list of tickers) and output one Excel sheet with key metrics
  side by side. 6 metric sections: Trading, Income, Returns, Capital Structure,
  Valuation, FCF. Market cap from yfinance; EDGAR XBRL for IS/BS/CF items.
  Handles mixed 10-K / 20-F peers automatically; EV multiples marked N/A for
  non-USD filers. Tested: 10-ticker semiconductor peer group (NVDA, AVGO, TSM, AMD,
  QCOM, TXN, MU, INTC, AMAT, LRCX). Output: `data/{PEER_GROUP}_comps.xlsx`
  **Limitations**:
  - 20-F filers (TSM): absolute values in local currency (TWD); EV/FCF yield N/A
  - Net cash balance uses cash & equivalents only — excludes short-term investments
  - Market cap is live from yfinance, not period-end

- [x] **`95_quarterly.py` — Quarterly data + LTM** ✓ 2026-04-26
  Use `company.get_quarterly_financials()` to pull 10-Q data and build:
  - LTM revenue and earnings: sum of last 4 pure-quarterly (Qn) periods
  - QoQ and YoY growth rates per line item
  Essential for valuations that need the most recent quarter, not the last annual filing.
  **Limitation**: 6-K filers (SPOT, ASML, TSM) have no XBRL in interim reports — script
  raises a clear error. Use yfinance (`yf.Ticker(ticker).quarterly_financials`) for these.

- [x] **`95b_quarterly_yf.py` — Quarterly data for foreign issuers (yfinance fallback)** ✓ 2026-04-30
  Same output format as `95_quarterly.py` but sourced from yfinance instead of EDGAR XBRL.
  Tested: ASML (EUR), SPOT (EUR), TSM (TWD). Curated row specs (IS/BS/CF), LTM = sum of
  last 4 quarters, shares use latest-Q for LTM, currency from `financialCurrency` field.
  **Limitations**:
  - yfinance returns restated numbers, not point-in-time SEC filings
  - Only ~6 quarters deep (yfinance limitation)
  - TSM Q1 2026 IS data incomplete — LTM shows N/A when most recent quarter is missing
  - No XBRL concept column (yfinance has no concept mapping)

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

### 2026-04-30 — 95b_quarterly_yf.py build

| Decision | Choice | Rationale |
|---|---|---|
| LTM = sum of last 4 quarters | Simple 4-quarter sum | yfinance quarterly data includes Q4 (unlike EDGAR 10-Q), so sum is always valid; no annual + stub formula needed |
| Curated row specs instead of XBRL template | IS_SPEC / BS_SPEC / CF_SPEC dicts | yfinance has no XBRL concept metadata; curated list gives clean, consistent output across all tickers |
| `financialCurrency` for currency label | `info.get("financialCurrency")` | `fast_info.currency` returns ADR trading currency (USD for TSM/ASML NYSE listings); `financialCurrency` correctly returns TWD and EUR |
| `is_avg=True` for shares rows | Latest-Q for LTM instead of sum | Summing weighted-average shares over 4 quarters is meaningless; latest Q is the correct period-end share count |
| Empty-section suppression | Trailing section headers dropped | Some companies lack certain rows (e.g. TSM missing R&D in some quarters); prevents orphan headers in Excel |

### 2026-04-28 — 96_comps.py build

| Decision | Choice | Rationale |
|---|---|---|
| Metrics source split | Market cap from yfinance; fundamentals from EDGAR XBRL | EDGAR has no real-time market data; yfinance `fast_info.market_cap` is a single API call |
| EV multiples for 20-F filers | Mark N/A | 20-F filers report in local currency; adding FX conversion would require a separate data source and introduce staleness risk |
| Net debt definition | Short-term debt + long-term debt - cash | Cash here is cash & equivalents only (not short-term investments) — simple and consistent across filers |
| EBITDA fallback | `operating_income` when D&A not found | EBITDA calculation requires D&A; if D&A lookup fails, show EBIT as a conservative fallback rather than N/A |
| Peer group structure | COMP_SECTIONS dict with per-row fmt strings | Allows arbitrary section headers and number formats per metric without any if/else in the render loop |
| Standalone script (no imports from 97_ratios) | Duplicate `_get_item` / `_prepare_structured_df` | Avoids inter-script coupling; each example is a self-contained, runnable artefact |

### 2026-04-26 — 97_ratios.py build and cross-filer tuning

| Decision | Choice | Rationale |
|---|---|---|
| sc-exact only; no sc-contains | Removed sc-contains from `_get_item` | "NonCurrentAssets".contains("CurrentAssets") was a systematic false positive; sc-exact + label patterns are sufficient |
| Tiebreaker: level then label length | Sort matches by (level asc, label_len asc) | Headline items are shallower (lower level) and shorter than supplemental disclosures; catches GOOG's capex supplemental note |
| capex exclude_patterns | ["included in accrued", "non-cash"] | Filters GOOG's "included in accrued liabilities" note and MU's "Non-cash equipment acquisitions" |
| Gross profit fallback | `revenue - cogs` when gross_prof not found | GOOG doesn't report a gross profit subtotal; identity always valid |
| ROIC tax rate fallback | `1 - (net_income / pretax)` when tax_exp not found | MU doesn't label tax expense in a way we can match; effective rate from actuals is equivalent |
| edgartools sc names | "CapitalExpenses", "DepreciationExpense", "TradeReceivables" | edgartools uses its own mapped sc names, not US GAAP concept names, for these three items |

---

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