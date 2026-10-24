# Market Summary Dashboard

**Live page: https://ericliu76.github.io/market-dashboard/**

A daily market summary with two tabs:

- **Markets**: a card for every instrument, grouped into Broad Market Index (global, Asian, EMEA and LATAM sub-groups), Equity Stock and Commodities. Each card has a sparkline, the change over the chosen range, the last day's change, the 52-week range and badges for unusual moves (σ), 52-week highs and lows and 200-day line crosses. Click a card for a full chart with 50- and 200-day lines, a compare overlay and stats. A signals strip and a sortable table of every market sit above and below the cards.
- **Calendar**: earnings dates in month and agenda views, marked (Confirmed) or (Estimate), with actual revenue and EPS for reported quarters and consensus for upcoming ones.

The page is static: `index.html` loads `dashboard_page.html` (the dashboard's own code) and `dash-shim.js`, which reads the CSV files in `data/`. Filters and the open chart are kept in the URL, so a link reopens the same view.

## Data

| Item | Source |
|---|---|
| Daily closes (5 years), snapshot statistics | Wind Terminal (`build_data.py`) |
| Earnings dates, actual revenue, GAAP diluted EPS | Wind |
| Confirmed or estimated upcoming dates | Nasdaq.com (Zacks) |
| Consensus (quarterly and fiscal-year revenue and adjusted EPS), EPS beat or miss | Yahoo Finance |

Stock prices are forward-adjusted for splits and dividends; commodities are Wind continuous futures. The data belongs to its providers; this page is for information only and is not investment advice.

## Files

| File | What it is |
|---|---|
| `index.html`, `dash-shim.js`, `dashboard_page.html` | The page |
| `data/market_snapshot.csv` | Latest value, 1D to 5Y changes, moving averages, 52-week range, volatility and signals per instrument |
| `data/master_data_file.csv` | Daily history: rows 1 to 8 describe each column (type, name, code, frequency, field, currency, unit, source), row 10 names the columns, data from row 11 |
| `data/earnings_calendar.csv` | One row per stock and fiscal quarter |
| `universe.csv` | The instrument list |
| `build_data.py`, `build_calendar.py` | Build the data files (need Wind Terminal open and logged in) |
| `publish_github.py` | Copies the site and data into this repo and pushes |

## Refreshing

1. `python build_data.py` (also runs `build_calendar.py`)
2. `python publish_github.py`

GitHub Pages shows the new data a minute or two after the push.
