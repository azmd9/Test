"""Fetch currency exchange rates from the Frankfurter API."""

import logging
from datetime import date

import requests

import config

logger = logging.getLogger(__name__)


def _parse_pairs() -> tuple[set[str], list[tuple[str, str]]]:
    """Parse configured currency pairs into needed symbols and pair list.

    Returns:
        (all_symbols, pairs) where all_symbols is the set of currencies
        needed from the EUR-based API call, and pairs is the list of
        (from_currency, to_currency) tuples.
    """
    symbols = set()
    pairs = []
    for pair_str in config.CURRENCY_PAIRS:
        frm, to = pair_str.strip().split(":")
        pairs.append((frm, to))
        # We always fetch EUR-based rates, so collect all non-EUR currencies
        if frm != "EUR":
            symbols.add(frm)
        if to != "EUR":
            symbols.add(to)
    return symbols, pairs


def _fetch_eur_rates(endpoint: str) -> dict:
    """Fetch EUR-based rates for all symbols needed by configured pairs.

    Args:
        endpoint: API endpoint path (e.g. '/latest' or '/2025-01-15').

    Returns:
        Dict of currency -> rate (relative to EUR).
    """
    symbols, _ = _parse_pairs()
    if not symbols:
        return {}

    params = {"base": "EUR", "symbols": ",".join(sorted(symbols))}
    url = f"{config.API_BASE_URL}{endpoint}"
    logger.info("Fetching EUR-based rates from %s", url)

    resp = requests.get(url, params=params, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    return data.get("date", str(date.today())), data.get("rates", {})


def _compute_pair_rates(eur_rates: dict) -> dict:
    """Compute exchange rates for all configured pairs using EUR rates.

    For EUR:X pairs, the rate is taken directly.
    For X:Y cross-rate pairs, rate = EUR_Y / EUR_X.

    Args:
        eur_rates: Dict of currency -> rate relative to EUR.

    Returns:
        Dict mapping "FROM/TO" -> rate for each configured pair.
    """
    _, pairs = _parse_pairs()
    result = {}
    for frm, to in pairs:
        if frm == "EUR":
            rate = eur_rates.get(to)
        elif to == "EUR":
            eur_from = eur_rates.get(frm)
            rate = round(1.0 / eur_from, 6) if eur_from else None
        else:
            eur_from = eur_rates.get(frm)
            eur_to = eur_rates.get(to)
            if eur_from and eur_to:
                rate = round(eur_to / eur_from, 6)
            else:
                rate = None

        key = f"{frm}/{to}"
        if rate is not None:
            result[key] = rate
        else:
            logger.warning("Could not compute rate for %s", key)
    return result


def fetch_latest_rates() -> dict:
    """Fetch latest exchange rates for all configured currency pairs.

    Returns:
        Dict with keys: date, pairs (dict of "FROM/TO" -> rate).
    """
    rate_date, eur_rates = _fetch_eur_rates("/latest")
    pair_rates = _compute_pair_rates(eur_rates)
    logger.info("Retrieved %d pair rates for %s", len(pair_rates), rate_date)
    return {"date": rate_date, "pairs": pair_rates}


def fetch_historical_rates(target_date: str) -> dict:
    """Fetch historical rates for all configured currency pairs.

    Args:
        target_date: Date string in YYYY-MM-DD format.

    Returns:
        Dict with keys: date, pairs (dict of "FROM/TO" -> rate).
    """
    rate_date, eur_rates = _fetch_eur_rates(f"/{target_date}")
    pair_rates = _compute_pair_rates(eur_rates)
    logger.info("Retrieved %d historical pair rates for %s", len(pair_rates), rate_date)
    return {"date": rate_date, "pairs": pair_rates}
