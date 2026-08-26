# NSE ETF RRG Mobile V1

Mobile-first Streamlit interface for the approved NSE ETF RRG scanner.

## Mobile pages
- Dashboard
- Intraday

## Dashboard shows only
- Top 3 Equity Groups
- Top 3 Swing Buy Priority
- Top 3 Near Buy Watchlist
- Top 3 Negative / Rebound candidates
- Deep Fall Watch >=1% only when available
- Top 3 Unusual Volume

## Intraday shows only
- Top 3 Intraday Groups
- Top 3 strict Intraday Buy candidates
- Top 3 Intraday Near Buy candidates

ETF-name colors:
- Green = qualified buy
- Amber = watch
- Red = deep fall / danger
- Blue = informational / unusual volume

The scanner logic is derived from the frozen V4.8.3 desktop scanner.
