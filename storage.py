"""Store and retrieve currency rate data as Excel and JSON."""

import json
import logging
import os
from datetime import datetime, timezone
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
        rate_data: Dict with date, pairs, and sources.

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
    """Save rate data as a row in a timestamped Excel file.

    The filename includes the run date and time (UTC) for traceability.
    Columns: date, then one column per currency pair, then source per pair.

    Args:
        rate_data: Dict with date, pairs, and sources.

    Returns:
        Path to the written file.
    """
    _ensure_data_dir()
    now_utc = datetime.now(timezone.utc)
    timestamp = now_utc.strftime("%Y-%m-%d_%H%M%S")
    filepath = os.path.join(config.DATA_DIR, f"rates_{timestamp}.xlsx")
    pair_keys = sorted(rate_data["pairs"].keys())
    sources = rate_data.get("sources", {})

    wb = Workbook()
    ws = wb.active
    ws.title = "Currency Rates"

    # Header row: date | rates... | source columns...
    headers = ["date"] + pair_keys + [f"{k} source" for k in pair_keys]
    ws.append(headers)

    # Data row
    row = (
        [rate_data["date"]]
        + [rate_data["pairs"].get(k, "") for k in pair_keys]
        + [sources.get(k, "") for k in pair_keys]
    )
    ws.append(row)
    wb.save(filepath)

    # Also save/update the cumulative rates.xlsx (append row)
    cumulative_path = os.path.join(config.DATA_DIR, "rates.xlsx")
    if os.path.exists(cumulative_path):
        wb_cum = load_workbook(cumulative_path)
        ws_cum = wb_cum.active
    else:
        wb_cum = Workbook()
        ws_cum = wb_cum.active
        ws_cum.title = "Currency Rates"
        ws_cum.append(headers)

    ws_cum.append(row)
    wb_cum.save(cumulative_path)

    logger.info("Saved Excel rates to %s and %s", filepath, cumulative_path)
    return filepath


def save_daily_snapshot(rate_data: dict) -> str:
    """Save a standalone JSON file for a single day's rates.

    Args:
        rate_data: Dict with date, pairs, and sources.

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
