"""
Month-to-month closing-balance persistence (INI file), so a new
month's opening balance can auto-fill from the previous month's
closing balance instead of requiring manual copy-paste.
"""
import configparser
import os

from config import MONTHS, BALANCE_STORE_PATH


def _month_key(month, year):
    return f"{str(year)}-{month.upper()}"


def _prev_month_year(month, year):
    idx = MONTHS.index(month.upper())
    if idx == 0:
        return MONTHS[-1], str(int(year) - 1)
    return MONTHS[idx - 1], str(year)


def _next_month_year(month, year):
    idx = MONTHS.index(month.upper())
    if idx == 11:
        return MONTHS[0], str(int(year) + 1)
    return MONTHS[idx + 1], str(year)


def load_balance_store():
    cp = configparser.ConfigParser()
    if os.path.exists(BALANCE_STORE_PATH):
        cp.read(BALANCE_STORE_PATH)
    return cp


def save_month_closing_balance(month, year, closing_balance):
    """Record this month's closing balance so next month can pick it up
    automatically as its opening balance."""
    cp = load_balance_store()
    key = _month_key(month, year)
    if key not in cp:
        cp[key] = {}
    cp[key]["closing_balance"] = str(int(closing_balance))
    try:
        with open(BALANCE_STORE_PATH, "w") as f:
            cp.write(f)
    except OSError:
        pass  # non-fatal — worst case, next month's auto-fill is skipped


def get_prev_month_closing_balance(month, year):
    """Returns the stored closing balance for the month immediately
    before (month, year), or None if it hasn't been generated/recorded."""
    prev_month, prev_year = _prev_month_year(month, year)
    cp = load_balance_store()
    key = _month_key(prev_month, prev_year)
    if key in cp and "closing_balance" in cp[key]:
        try:
            return int(cp[key]["closing_balance"])
        except ValueError:
            return None
    return None

