"""
データ取得モジュール: yfinanceでヤフーファイナンスからデータ取得
"""

import yfinance as yf
import pandas as pd
from typing import Optional


def fetch_stock_data(
    ticker: str,
    period: str = "1y",
    interval: str = "1d"
) -> Optional[pd.DataFrame]:
    """
    指定した銘柄の株価データを取得

    Args:
        ticker: 銘柄コード（例: "7203.T"）
        period: 期間（例: "1y", "6mo", "2y"）
        interval: 間隔（例: "1d", "1wk"）

    Returns:
        株価データのDataFrame（Date, Open, High, Low, Close, Volume）
        取得失敗時はNone
    """
    try:
        stock = yf.Ticker(ticker)
        df = stock.history(period=period, interval=interval)

        if df.empty:
            print(f"警告: {ticker} のデータが取得できませんでした")
            return None

        # 必要なカラムのみ保持
        df = df[["Open", "High", "Low", "Close", "Volume"]]
        df.index = pd.to_datetime(df.index).tz_localize(None)

        return df

    except Exception as e:
        print(f"エラー: {ticker} のデータ取得中にエラーが発生しました: {e}")
        return None


def fetch_multiple_stocks(
    tickers: list[str],
    period: str = "1y"
) -> dict[str, pd.DataFrame]:
    """
    複数銘柄の株価データを取得

    Args:
        tickers: 銘柄コードのリスト
        period: 期間

    Returns:
        銘柄コードをキー、DataFrameを値とする辞書
    """
    result = {}
    for ticker in tickers:
        print(f"  {ticker} を取得中...")
        df = fetch_stock_data(ticker, period)
        if df is not None:
            result[ticker] = df
    return result
