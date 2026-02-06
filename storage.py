"""Store and retrieve currency rate data as Excel and JSON."""

import json
import logging
import os
from pathlib import Path

from openpyxl import Workbook, load_workbook

import config

logger = logging.getLogger(__name__)


def _ensure_data_dir():
    """Create the data directory if it doesn't exist."""
    Path(config.DATA_DIR).mkdir(parents=True, exist_ok=True)


def save_rates_json(rate_data: dict) -> str:
    """Append rate data to a JSON Lines file (one JSON object per line).

    Args:
        rate_data: Dict with date, base, and rates.

    Returns:
        Path to the written file.
    """
    _ensure_data_dir()
    filepath = os.path.join(config.DATA_DIR, "rates.jsonl")

    with open(filepath, "a") as f:
        f.write(json.dumps(rate_data) + "\n")

    logger.info("Saved JSON rates to %s", filepath)
    return filepath


def save_rates_excel(rate_data: dict) -> str:
    """Append rate data as a row in an Excel file.

    Args:
        rate_data: Dict with date, base, and rates.

    Returns:
        Path to the written file.
    """
    _ensure_data_dir()
    filepath = os.path.join(config.DATA_DIR, "rates.xlsx")
    currencies = sorted(rate_data["rates"].keys())

    if os.path.exists(filepath):
        wb = load_workbook(filepath)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Currency Rates"
        ws.append(["date", "base"] + currencies)

    ws.append(
        [rate_data["date"], rate_data["base"]]
        + [rate_data["rates"].get(c, "") for c in currencies]
    )
    wb.save(filepath)

    logger.info("Saved Excel rates to %s", filepath)
    return filepath


def save_daily_snapshot(rate_data: dict) -> str:
    """Save a standalone JSON file for a single day's rates.

    Args:
        rate_data: Dict with date, base, and rates.

    Returns:
        Path to the written file.
    """
    _ensure_data_dir()
    filepath = os.path.join(config.DATA_DIR, f"rates_{rate_data['date']}.json")

    with open(filepath, "w") as f:
        json.dump(rate_data, f, indent=2)

    logger.info("Saved daily snapshot to %s", filepath)
    return filepath


def load_all_rates_json() -> list[dict]:
    """Load all rate entries from the JSON Lines file.

    Returns:
        List of rate data dicts.
    """
    filepath = os.path.join(config.DATA_DIR, "rates.jsonl")
    if not os.path.exists(filepath):
        return []

    entries = []
    with open(filepath) as f:
        for line in f:
            line = line.strip()
            if line:
                entries.append(json.loads(line))
    return entries
