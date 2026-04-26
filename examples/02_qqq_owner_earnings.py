"""
Use case: owner-earnings yield screen on a basket of QQQ-style mega-caps.

Owner earnings (Buffett's definition, simplified):
    Owner Earnings = Net Income + D&A - CapEx
    Owner Earnings Yield = Owner Earnings / Market Cap

Higher yield = cheaper on a cash-generation basis. We pull the latest 10-K
financials via edgartools and the current market cap via yfinance.

This is point-in-time WRONG (uses today's market cap with last reported financials)
and is meant only to demonstrate the data plumbing. For a real backtest you
need point-in-time financials AND point-in-time prices.

Run:
    python examples/02_qqq_owner_earnings.py
"""
import os
from dotenv import load_dotenv

load_dotenv()
ua = os.getenv("SEC_USER_AGENT")
if not ua:
    raise SystemExit("SEC_USER_AGENT not set. See .env.example")
os.environ["EDGAR_IDENTITY"] = ua

from edgar import Company
import yfinance as yf

TICKERS = ["AAPL", "MSFT", "NVDA", "GOOGL", "AMZN", "META", "AVGO", "COST"]


def owner_earnings_yield(ticker: str) -> dict:
    co = Company(ticker)
    fin = co.get_financials()  # latest financials parsed from filings
    if fin is None:
        return {"ticker": ticker, "error": "no financials"}

    try:
        income = fin.income_statement()
        cash = fin.cash_flow_statement()
        net_income = float(income.loc["NetIncomeLoss"].iloc[0])
        # Common XBRL tag names — edgartools normalizes some, may need adjusting
        dna = float(cash.loc["DepreciationDepletionAndAmortization"].iloc[0])
        capex = float(cash.loc["PaymentsToAcquirePropertyPlantAndEquipment"].iloc[0])
        owner_earn = net_income + dna - capex

        mkt_cap = yf.Ticker(ticker).info.get("marketCap")
        if not mkt_cap:
            return {"ticker": ticker, "error": "no market cap"}

        return {
            "ticker": ticker,
            "owner_earnings_$B": round(owner_earn / 1e9, 2),
            "market_cap_$B": round(mkt_cap / 1e9, 2),
            "oe_yield_%": round(100 * owner_earn / mkt_cap, 2),
        }
    except Exception as e:
        return {"ticker": ticker, "error": str(e)[:80]}


if __name__ == "__main__":
    rows = [owner_earnings_yield(t) for t in TICKERS]
    import pandas as pd
    df = pd.DataFrame(rows)
    if "oe_yield_%" in df.columns:
        df = df.sort_values("oe_yield_%", ascending=False, na_position="last")
    print(df.to_string(index=False))
    print(
        "\nReminder: this snapshot mixes today's market cap with last 10-K financials. "
        "Treat as illustrative, NOT as a backtestable signal."
    )
