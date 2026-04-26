"""
Hello-world for edgartools.

Verifies the install works end-to-end by:
  1. Loading SEC_USER_AGENT from .env
  2. Looking up Apple (AAPL) by ticker
  3. Listing the last 3 10-K filings

Run:
    python examples/01_hello_world.py
"""
import os
from dotenv import load_dotenv

load_dotenv()

# edgartools reads identity from EDGAR_IDENTITY env var.
# We load from SEC_USER_AGENT (our naming convention) and forward it.
ua = os.getenv("SEC_USER_AGENT")
if not ua:
    raise SystemExit(
        "SEC_USER_AGENT not set. Copy .env.example to .env and add your name/email."
    )
os.environ["EDGAR_IDENTITY"] = ua

from edgar import Company

c = Company("AAPL")
print(f"Company: {c.name}  |  CIK: {c.cik}")

print("\nLast 3 10-K filings:")
for f in c.get_filings(form="10-K").head(3):
    print(f"  {f.filing_date}  {f.form:<6}  {f.accession_no}")

print("\nOK — edgartools is working.")
