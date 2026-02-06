"""Configuration for currency rate retrieval."""

import os
from dotenv import load_dotenv

load_dotenv()

# Base currency to fetch rates against
BASE_CURRENCY = os.getenv("BASE_CURRENCY", "USD")

# Currencies to track
TARGET_CURRENCIES = os.getenv(
    "TARGET_CURRENCIES", "EUR,GBP,JPY,CAD,AUD,CHF,CNY,INR,BRL,MXN"
).split(",")

# Frankfurter API (free, no key required)
API_BASE_URL = "https://api.frankfurter.dev/v1"

# Data storage directory
DATA_DIR = os.getenv("DATA_DIR", "data")

# Schedule time (24h format)
SCHEDULE_TIME = os.getenv("SCHEDULE_TIME", "09:00")
