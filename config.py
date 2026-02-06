"""Configuration for currency rate retrieval."""

import os
from dotenv import load_dotenv

load_dotenv()

# Base currency to fetch rates against
BASE_CURRENCY = os.getenv("BASE_CURRENCY", "EUR")

# Currencies to track (all 30 non-EUR currencies supported by Frankfurter)
TARGET_CURRENCIES = os.getenv(
    "TARGET_CURRENCIES",
    "AUD,BGN,BRL,CAD,CHF,CNY,CZK,DKK,GBP,HKD,"
    "HUF,IDR,ILS,INR,ISK,JPY,KRW,MXN,MYR,NOK,"
    "NZD,PHP,PLN,RON,SEK,SGD,THB,TRY,USD,ZAR",
).split(",")

# Frankfurter API (free, no key required)
API_BASE_URL = "https://api.frankfurter.dev/v1"

# Data storage directory
DATA_DIR = os.getenv("DATA_DIR", "data")

# Schedule time (24h format)
SCHEDULE_TIME = os.getenv("SCHEDULE_TIME", "09:00")
