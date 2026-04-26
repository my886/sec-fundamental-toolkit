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

- [ ] **`95_quarterly.py` — Quarterly data + LTM**
  Use `company.get_quarterly_financials()` to pull 10-Q data and build:
  - LTM revenue and earnings: last annual + last two quarters - same two quarters prior year
  - QoQ and YoY growth rates per line item
  Essential for valuations that need the most recent quarter, not the last annual filing.

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

### 2026-04-26 — Initial build session

| Decision | Choice | Rationale |
|---|---|---|
| Multi-period structure approach | Single-filing template + XBRLS join by `concept` | XBRLS stitched output loses row order and section headers; joining by XBRL concept key onto the latest filing's template preserves the document hierarchy |
| Value scaling | USD millions; detect per-share rows by `max(abs) < 1 000` | Consistent with financial modelling convention; avoids hardcoding concept names which vary across filers |
| Form type detection | Auto-detect 10-K then 20-F at runtime | Foreign private issuers (ASML, TSM, GFS, TSEM) file 20-F; hardcoding breaks silently for these |
| Repo visibility | Public | Portfolio signal for equity analyst transition — combination of financial statement fluency + engineering is differentiated |
| Excel export library | openpyxl directly (not `pd.ExcelWriter`) | Needed for cell-level formatting: indentation, section header colours, number formats per row |
