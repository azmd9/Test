"""Fetch currency exchange rates from multiple API sources.

Primary source:   Frankfurter API (ECB data) - https://api.frankfurter.dev
Secondary source: ExchangeRate-API           - https://open.er-api.com
"""

import logging
from datetime import date

import requests

import config

logger = logging.getLogger(__name__)

# Source labels for audit trail
SOURCE_FRANKFURTER = "Frankfurter/ECB (api.frankfurter.dev)"
SOURCE_EXCHANGERATE = "ExchangeRate-API (open.er-api.com)"


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
        if frm != "EUR":
            symbols.add(frm)
        if to != "EUR":
            symbols.add(to)
    return symbols, pairs


def _fetch_frankfurter(endpoint: str, symbols: set[str]) -> tuple[str, dict]:
    """Fetch EUR-based rates from the Frankfurter API (ECB data).

    Args:
        endpoint: API endpoint path (e.g. '/latest' or '/2025-01-15').
        symbols: Set of currency codes to fetch.

    Returns:
        (rate_date, rates_dict) where rates_dict maps currency -> rate.
    """
    # Filter out currencies not supported by Frankfurter
    supported = symbols - config.SECONDARY_ONLY_CURRENCIES
    if not supported:
        return str(date.today()), {}

    params = {"base": "EUR", "symbols": ",".join(sorted(supported))}
    url = f"{config.API_BASE_URL}{endpoint}"
    logger.info("Fetching rates from Frankfurter API: %s", url)

    resp = requests.get(url, params=params, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    return data.get("date", str(date.today())), data.get("rates", {})


def _fetch_secondary(symbols: set[str]) -> dict:
    """Fetch EUR-based rates from the secondary ExchangeRate-API.

    Only fetches currencies listed in config.SECONDARY_ONLY_CURRENCIES.

    Args:
        symbols: Set of currency codes needed.

    Returns:
        Dict of currency -> rate (relative to EUR).
    """
    needed = symbols & config.SECONDARY_ONLY_CURRENCIES
    if not needed:
        return {}

    url = f"{config.SECONDARY_API_BASE_URL}/EUR"
    logger.info("Fetching rates from ExchangeRate-API: %s (for %s)", url, sorted(needed))

    resp = requests.get(url, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    all_rates = data.get("rates", {})

    # Return only the currencies we need from this source
    return {ccy: all_rates[ccy] for ccy in needed if ccy in all_rates}


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


def _build_source_map(eur_rates_primary: dict, eur_rates_secondary: dict) -> dict:
    """Build an audit trail mapping each currency to its data source.

    Args:
        eur_rates_primary: Rates fetched from Frankfurter/ECB.
        eur_rates_secondary: Rates fetched from ExchangeRate-API.

    Returns:
        Dict mapping "FROM/TO" -> source string for each configured pair.
    """
    _, pairs = _parse_pairs()
    source_map = {}
    for frm, to in pairs:
        key = f"{frm}/{to}"
        sources = set()

        for ccy in (frm, to):
            if ccy == "EUR":
                continue
            if ccy in eur_rates_secondary:
                sources.add(SOURCE_EXCHANGERATE)
            elif ccy in eur_rates_primary:
                sources.add(SOURCE_FRANKFURTER)

        if len(sources) == 0:
            source_map[key] = SOURCE_FRANKFURTER  # EUR/EUR edge case
        elif len(sources) == 1:
            source_map[key] = sources.pop()
        else:
            # Cross-rate using currencies from different sources
            source_map[key] = " + ".join(sorted(sources))
    return source_map


def _fetch_all_rates(endpoint: str) -> dict:
    """Fetch rates from both APIs, merge, compute pairs, and build audit trail.

    Args:
        endpoint: Frankfurter endpoint (e.g. '/latest' or '/2025-01-15').

    Returns:
        Dict with keys: date, pairs, sources.
    """
    symbols, _ = _parse_pairs()

    # Fetch from both sources
    rate_date, primary_rates = _fetch_frankfurter(endpoint, symbols)
    secondary_rates = _fetch_secondary(symbols)

    # Merge rates (secondary fills in what primary doesn't have)
    merged_rates = {**primary_rates, **secondary_rates}

    # Compute pair rates and build source audit trail
    pair_rates = _compute_pair_rates(merged_rates)
    source_map = _build_source_map(primary_rates, secondary_rates)

    logger.info("Retrieved %d pair rates for %s (%d from primary, %d from secondary)",
                len(pair_rates), rate_date, len(primary_rates), len(secondary_rates))

    return {"date": rate_date, "pairs": pair_rates, "sources": source_map}


def fetch_latest_rates() -> dict:
    """Fetch latest exchange rates for all configured currency pairs.

    Returns:
        Dict with keys: date, pairs (dict of "FROM/TO" -> rate),
        sources (dict of "FROM/TO" -> source string).
    """
    return _fetch_all_rates("/latest")


def fetch_historical_rates(target_date: str) -> dict:
    """Fetch historical rates for all configured currency pairs.

    Note: Historical rates from the secondary API are not date-specific;
    it always returns the latest available rate.

    Args:
        target_date: Date string in YYYY-MM-DD format.

    Returns:
        Dict with keys: date, pairs, sources.
    """
    return _fetch_all_rates(f"/{target_date}")
