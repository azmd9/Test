# Automated Currency Rate Daily Retrieval

Fetches daily currency exchange rates from the [Frankfurter API](https://frankfurter.dev/) and stores them in Excel and JSON formats.

## Features

- **Daily automation** via GitHub Actions (runs at 09:00 UTC)
- **Local scheduling** for self-hosted continuous retrieval
- **Historical rate lookup** for any past date
- **Dual storage** — Excel for spreadsheets, JSON Lines for programmatic access, plus per-day snapshots
- **Configurable** currency pairs via environment variables, with cross-rate support

## Quick Start

```bash
pip install -r requirements.txt

# Fetch latest rates once
python main.py

# Fetch historical rates
python main.py --date 2025-01-15

# Run on a daily schedule locally
python main.py --schedule
```

## Configuration

Set via environment variables or a `.env` file (see `.env.example`):

| Variable | Default | Description |
|---|---|---|
| `CURRENCY_PAIRS` | `EUR:USD,EUR:CNY,...,CNY:SGD,CNY:USD` | Comma-separated `FROM:TO` pairs |
| `DATA_DIR` | `data` | Directory for stored rate files |
| `SCHEDULE_TIME` | `09:00` | Local time for daily retrieval (24h) |

EUR-based pairs are fetched directly from the API. Cross-rate pairs (e.g. `CNY:SGD`) are computed via EUR as an intermediary.

## GitHub Actions

The included workflow (`.github/workflows/daily-rates.yml`) runs daily at 09:00 UTC and commits updated rate data to the repository. You can also trigger it manually from the Actions tab.

Override pairs via the GitHub repository variable `CURRENCY_PAIRS`.

## Output Files

| File | Format | Description |
|---|---|---|
| `data/rates.jsonl` | JSON Lines | Append-only log of all fetched rates |
| `data/rates.xlsx` | Excel | Tabular rate history |
| `data/rates_YYYY-MM-DD.json` | JSON | Per-day snapshot |
