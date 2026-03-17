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
    "EUR:AED,EUR:CLP,EUR:RUB,EUR:SAR,EUR:PEN,EUR:MAD,EUR:COP,"
    "CNY:SGD,CNY:USD",
).split(",")

# --- API Sources ---
# Primary: Frankfurter API (ECB data, free, no key required)
# Supports: USD, JPY, BGN, CZK, DKK, GBP, HUF, PLN, RON, SEK, CHF, ISK, NOK,
#           TRY, AUD, BRL, CAD, CNY, HKD, IDR, ILS, INR, KRW, MXN, MYR, NZD,
#           PHP, SGD, THB, ZAR
API_BASE_URL = "https://api.frankfurter.dev/v1"

# Secondary: ExchangeRate-API (free, no key required)
# Used for currencies not available in Frankfurter/ECB:
#   AED (UAE Dirham), CLP (Chilean Peso), RUB (Russian Ruble),
#   SAR (Saudi Riyal), PEN (Peruvian Sol), MAD (Moroccan Dirham),
#   COP (Colombian Peso)
SECONDARY_API_BASE_URL = "https://open.er-api.com/v6/latest"

# Currencies only available from the secondary API source
SECONDARY_ONLY_CURRENCIES = {"AED", "CLP", "RUB", "SAR", "PEN", "MAD", "COP"}

# Data storage directory
DATA_DIR = os.getenv("DATA_DIR", "data")

# Schedule time (24h format)
SCHEDULE_TIME = os.getenv("SCHEDULE_TIME", "09:00")
