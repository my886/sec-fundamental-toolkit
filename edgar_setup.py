"""
Shared setup for edgartools scripts.

Import this at the top of any script that uses the `edgar` package, BEFORE
importing edgar:

    from edgar_setup import init_edgar
    init_edgar()
    from edgar import Company
    ...

It loads .env, validates SEC_USER_AGENT, and forwards it to edgartools'
EDGAR_IDENTITY env var (which the library actually reads).
"""
import os
from dotenv import load_dotenv


def init_edgar() -> str:
    """Load .env and set EDGAR_IDENTITY. Returns the identity string."""
    load_dotenv()
    ua = os.getenv("SEC_USER_AGENT", "").strip()
    if not ua or ua.startswith("Your Name"):
        raise SystemExit(
            "SEC_USER_AGENT is not set. Copy .env.example to .env and add your "
            "real name + email (SEC requires this in the User-Agent header)."
        )
    os.environ["EDGAR_IDENTITY"] = ua
    return ua


if __name__ == "__main__":
    print("EDGAR_IDENTITY =", init_edgar())
