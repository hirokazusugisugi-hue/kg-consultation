# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A Japanese stock backtesting framework with moving average crossover strategy analysis for 25 Tokyo Stock Exchange equities across 5 sectors (automotive, electronics, banking, telecom, retail). Includes CLI, automated demo, and Streamlit web interface.

All UI text is in Japanese.

## Commands

### Stock Backtest Application
```bash
# Install Python dependencies
pip install -r stock_backtest/requirements.txt

# Interactive CLI
python stock_backtest/main.py

# Automated demo (generates automotive sector charts)
python stock_backtest/demo.py

# Streamlit web app (requires: pip install streamlit)
streamlit run stock_backtest/app.py
```

### Root-level Scripts
```bash
python calculator.py    # Interactive calculator
python todo_app.py      # Todo list manager (persists to todos.json)
```

## Architecture

### stock_backtest/ - Pipeline Architecture

Data flows through a modular pipeline: **data → strategy → backtest → analysis → visualization**

- **config.py** — Central configuration: MA parameters (5/25 day), initial capital (¥1M), sector/stock definitions (25 stocks with `.T` suffix tickers), time period options
- **data_fetcher.py** — Yahoo Finance API wrapper via `yfinance`. Returns pandas DataFrames with OHLCV data
- **strategy.py** — Moving average crossover: calculates SMA, detects golden/dead crosses, outputs Signal and Position columns
- **backtester.py** — Simulates portfolio performance from signals, calculates cumulative returns, compares against buy-and-hold benchmark
- **analyzer.py** — Computes metrics: Sharpe ratio, max drawdown, profit factor, win rate, trade count, excess return
- **visualizer.py** — Matplotlib charts with `japanize_matplotlib` for Japanese text. Generates stock price + MA + signals charts, portfolio value comparisons, sector comparisons, drawdown plots

### Entry Points

- **main.py** — Interactive menu-driven CLI with sector/stock/period selection
- **demo.py** — Non-interactive batch run for automotive sector, generates PNG charts
- **app.py** — Streamlit web app with 4 strategies (MA crossover, RSI, Bollinger Bands, MACD), customizable parameters via sliders, and screening filters. Note: `app.py` implements its own strategy logic independently from `strategy.py`

### Root-level Scripts

- **calculator.py** — Standalone interactive calculator
- **todo_app.py** — JSON-backed todo list manager using `todos.json`

## Key Dependencies

- `yfinance` — Stock data from Yahoo Finance
- `pandas` / `numpy` — Data manipulation
- `matplotlib` + `japanize_matplotlib` — Charting with Japanese font support
- `streamlit` — Web interface (for app.py only, not in requirements.txt)
- `@playwright/mcp` — Node.js dependency in package.json
