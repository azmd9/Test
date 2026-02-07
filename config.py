"""Configuration for currency rate retrieval."""

import os
from dotenv import load_dotenv

load_dotenv()

# Currency pairs to retrieve (FROM -> TO)
# Pairs are defined as "FROM:TO" separated by commas.
# All pairs using EUR as base are fetched directly from the API.
# Cross-rate pairs (e.g. CNY:SGD) are computed via EUR as intermediary.
CURRENCY_PAIRS = os.getenv(
    "CURRENCY_PAIRS",
    "EUR:USD,EUR:CNY,EUR:BRL,EUR:BGN,EUR:GBP,EUR:SGD,"
    "EUR:TRY,EUR:SEK,EUR:ILS,EUR:CHF,EUR:KRW,EUR:PHP,"
    "EUR:JPY,EUR:HKD,EUR:IDR,EUR:THB,EUR:DKK,"
    "CNY:SGD,CNY:USD",
).split(",")

# Frankfurter API (free, no key required)
API_BASE_URL = "https://api.frankfurter.dev/v1"

# Data storage directory
DATA_DIR = os.getenv("DATA_DIR", "data")

# Schedule time (24h format)
SCHEDULE_TIME = os.getenv("SCHEDULE_TIME", "09:00")
