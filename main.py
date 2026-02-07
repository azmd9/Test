#!/usr/bin/env python3
"""Currency rate daily retrieval - main entry point.

Usage:
    python main.py              Fetch once and exit
    python main.py --schedule   Run on a daily schedule
    python main.py --date 2025-01-15   Fetch historical rates for a date
"""

import argparse
import logging
import sys

from fetcher import fetch_latest_rates, fetch_historical_rates
from storage import save_rates_excel, save_rates_json, save_daily_snapshot
from scheduler import start_scheduler

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger(__name__)


def _print_rates(rate_data: dict):
    """Print fetched pair rates to stdout."""
    print(f"Rates for {rate_data['date']}:")
    for pair, rate in sorted(rate_data["pairs"].items()):
        print(f"  {pair}: {rate}")


def run_once():
    """Fetch latest rates, save them, and exit."""
    rates = fetch_latest_rates()
    save_rates_json(rates)
    save_rates_excel(rates)
    save_daily_snapshot(rates)
    _print_rates(rates)


def run_historical(target_date: str):
    """Fetch historical rates for a specific date."""
    rates = fetch_historical_rates(target_date)
    save_rates_json(rates)
    save_rates_excel(rates)
    save_daily_snapshot(rates)
    _print_rates(rates)


def main():
    parser = argparse.ArgumentParser(
        description="Automated daily currency rate retrieval"
    )
    parser.add_argument(
        "--schedule",
        action="store_true",
        help="Run continuously on a daily schedule",
    )
    parser.add_argument(
        "--date",
        type=str,
        help="Fetch historical rates for a specific date (YYYY-MM-DD)",
    )
    args = parser.parse_args()

    if args.schedule:
        logger.info("Starting in scheduled mode...")
        start_scheduler()
    elif args.date:
        run_historical(args.date)
    else:
        run_once()


if __name__ == "__main__":
    main()
