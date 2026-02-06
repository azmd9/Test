"""Schedule daily currency rate retrieval."""

import logging
import time

import schedule

import config
from fetcher import fetch_latest_rates
from storage import save_rates_excel, save_rates_json, save_daily_snapshot

logger = logging.getLogger(__name__)


def daily_job():
    """Fetch and store the latest currency rates."""
    logger.info("Running scheduled currency rate retrieval...")
    try:
        rates = fetch_latest_rates()
        save_rates_json(rates)
        save_rates_excel(rates)
        save_daily_snapshot(rates)
        logger.info("Daily retrieval complete for %s", rates["date"])
    except Exception:
        logger.exception("Failed to retrieve currency rates")


def start_scheduler():
    """Start the scheduler to run the daily job at the configured time."""
    schedule.every().day.at(config.SCHEDULE_TIME).do(daily_job)
    logger.info(
        "Scheduler started. Daily retrieval scheduled at %s", config.SCHEDULE_TIME
    )

    # Run immediately on first start, then follow schedule
    daily_job()

    while True:
        schedule.run_pending()
        time.sleep(60)
