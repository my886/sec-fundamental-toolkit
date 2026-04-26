# edgartools — Project Context for Claude Code

> Loaded automatically whenever Claude Code is invoked from this directory.
> Contains API quirks, design decisions, and hard-won lessons from past sessions.

---

## What this repo is

Python wrapper around SEC EDGAR. We use it for the Fundamental Analysis pillar:
pulling 10-K / 20-F financials, stitching multi-year time series, and exporting
structured Excel workbooks for financial modelling.

---

## Scripts

| Script | Purpose |
|---|---|
| `examples/99_income_statement.py` | Single ticker, latest 10-K — prints all three statements to console + styled Excel (`view="standard"`) |
| `examples/98_multiperiod_financials.py` | Single ticker, N years — XBRLS stitched values merged onto single-filing template structure, styled Excel |

**Configuration** for both scripts: change `TICKER` and (for 98) `YEARS` at the top.

---

## Known API quirks

### Statement objects
- `company.get_financials()` returns a `Financials` object, **not** a DataFrame.
- Calling `.income_statement()`, `.balance_sheet()`, `.cashflow_statement()` returns a `Statement` object — **not** a DataFrame.
- `Statement` has **no** `.to_excel()`. Call `.to_dataframe()` first.
- Method is `cashflow_statement()`, **not** `cash_flow_statement()`.

### to_dataframe() views
| View | Rows | Use for |
|---|---|---|
| `view="standard"` | ~22–38 | Excel export, console display — matches filing document order, includes `level` / `abstract` / `dimension` metadata |
| `view="detailed"` | 50+ | Segment-level analysis, dimensional breakdowns |
| `view="summary"` | ~15–20 | Quick multi-company comparisons |

The `level`, `abstract`, and `dimension` columns only exist in the single-filing
`to_dataframe()` output, **not** in the XBRLS stitched output.

### XBRLS multi-period stitching
- `XBRLS.from_filings(filings)` aligns concepts across years but **loses all
  presentational structure** (row order, section headers, indentation levels).
- **Do not** use `xbrls.statements.income_statement().to_dataframe()` directly as
  the display structure — it produces a flat, arbitrarily ordered dump.
- **Correct pattern**: use single-filing template for structure + XBRLS for values:
  ```python
  # 1. Template: row order, section headers, levels
  tmpl_df = company.get_financials().balance_sheet().to_dataframe(view="standard")

  # 2. Stitched: multi-period values
  xbrls    = XBRLS.from_filings(filings)
  stitched = xbrls.statements.balance_sheet().to_dataframe()

  # 3. Join by concept (XBRL concept name)
  #    Abstract rows (section headers) have no concept in stitched — that is correct.
  #    Dimensional breakdown rows (dimension=True in template) are absent from stitched — skip them.
  ```
- Join key is `concept` (full XBRL tag, e.g. `us-gaap_AssetsCurrent`).
- Period column format: single-filing → `2025-09-27 (FY)` | stitched → `2025-09-27`.
- Use `amendments=False` when fetching filings for stitching (doc recommendation).

### Numeric type safety
- Some filings return **string values** in period columns instead of floats.
- Always coerce before any arithmetic:
  ```python
  for c in period_cols:
      df[c] = pd.to_numeric(df[c], errors="coerce")
  ```

### Filing object attributes
- Date field is `filing.filing_date`, **not** `filing.date`.

---

## Foreign private issuers (20-F)

Companies incorporated outside the US file 20-F, not 10-K.
`company.get_filings(form="10-K")` returns **empty** for these — always fall back:

```python
for form in ("10-K", "20-F"):
    filings = company.get_filings(form=form).head(YEARS)
    if len(filings):
        break
```

Known 20-F filers in our universe: ASML, TSM (TSMC), GFS (GlobalFoundries), TSEM (Tower Semiconductor).

`company.get_financials()` auto-detects the form type and works for both.

---

## Ticker gotchas

| Intended company | Wrong ticker | Correct ticker |
|---|---|---|
| TSMC | `TSMC` | `TSM` |

If `Company(ticker)` raises `CompanyNotFoundError`, try `find_company("name")` to
discover the correct ticker or CIK.

---

## Windows / environment

- **Encoding**: Windows terminal defaults to cp1252. Avoid non-ASCII characters
  (`→`, `—`, `×`) in `print()` output — use plain ASCII alternatives (`to`, `-`, `x`).
- **Dependencies**: `openpyxl` is listed in `requirements.txt`. If you see
  `ModuleNotFoundError: No module named 'openpyxl'`, run:
  ```powershell
  uv pip install -r requirements.txt
  ```
- **pandas ≥ 2.1**: `DataFrame.applymap()` is removed — use `.map()` instead.

---

## Value scaling convention

All exported values are in **USD millions** except per-share metrics.

Detection heuristic: if `max(abs(values)) < 1_000` after numeric coercion,
treat as a per-share metric (EPS ~ 7.49) and skip scaling. Everything else
is divided by 1 000 000.

---

## Data output location

`E:\github_awesome\fundamental\edgartools\data\`

File naming: `{TICKER}_financials.xlsx` (single period) and
`{TICKER}_multiperiod_financials.xlsx` (XBRLS stitched).
