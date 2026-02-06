"""Fetch currency exchange rates from the Frankfurter API."""

import logging
from datetime import date

import requests

import config

logger = logging.getLogger(__name__)


def fetch_latest_rates() -> dict:
    """Fetch latest exchange rates for configured currencies.

    Returns:
        Dict with keys: date, base, rates (dict of currency -> rate).

    Raises:
        requests.RequestException: On network or API errors.
    """
    params = {
        "base": config.BASE_CURRENCY,
        "symbols": ",".join(config.TARGET_CURRENCIES),
    }
    url = f"{config.API_BASE_URL}/latest"
    logger.info("Fetching latest rates from %s", url)

    resp = requests.get(url, params=params, timeout=30)
    resp.raise_for_status()
    data = resp.json()

    result = {
        "date": data.get("date", str(date.today())),
        "base": data.get("base", config.BASE_CURRENCY),
        "rates": data.get("rates", {}),
    }
    logger.info("Retrieved rates for %s: %d currencies", result["date"], len(result["rates"]))
    return result


def fetch_historical_rates(target_date: str) -> dict:
    """Fetch historical rates for a specific date (YYYY-MM-DD).

    Args:
        target_date: Date string in YYYY-MM-DD format.

    Returns:
        Dict with keys: date, base, rates.

    Raises:
        requests.RequestException: On network or API errors.
    """
    params = {
        "base": config.BASE_CURRENCY,
        "symbols": ",".join(config.TARGET_CURRENCIES),
    }
    url = f"{config.API_BASE_URL}/{target_date}"
    logger.info("Fetching historical rates for %s", target_date)

    resp = requests.get(url, params=params, timeout=30)
    resp.raise_for_status()
    data = resp.json()

    return {
        "date": data.get("date", target_date),
        "base": data.get("base", config.BASE_CURRENCY),
        "rates": data.get("rates", {}),
    }
