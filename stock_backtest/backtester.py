"""
バックテストエンジン: 戦略のパフォーマンスをシミュレート
"""

import pandas as pd
import numpy as np
from config import INITIAL_CAPITAL


def run_backtest(
    df: pd.DataFrame,
    initial_capital: float = INITIAL_CAPITAL
) -> pd.DataFrame:
    """
    バックテストを実行

    Args:
        df: 戦略が適用されたDataFrame（Signal, Position列が必要）
        initial_capital: 初期資金

    Returns:
        バックテスト結果を含むDataFrame
    """
    df = df.copy()

    # NaNを除去（移動平均の計算で発生）
    df = df.dropna()

    if df.empty:
        return df

    # 日次リターンを計算
    df["Returns"] = df["Close"].pct_change()

    # 戦略リターン（前日のポジションに基づく）
    df["Strategy_Returns"] = df["Position"].shift(1) * df["Returns"]

    # 累積リターン
    df["Cumulative_Returns"] = (1 + df["Returns"]).cumprod()
    df["Cumulative_Strategy_Returns"] = (1 + df["Strategy_Returns"]).cumprod()

    # 資産推移
    df["Portfolio_Value"] = initial_capital * df["Cumulative_Strategy_Returns"]
    df["Buy_Hold_Value"] = initial_capital * df["Cumulative_Returns"]

    # 取引記録
    df["Trade"] = df["Signal"].apply(
        lambda x: "買い" if x == 1 else ("売り" if x == -1 else "")
    )

    return df


def get_trades(df: pd.DataFrame) -> pd.DataFrame:
    """
    取引履歴を抽出

    Args:
        df: バックテスト結果のDataFrame

    Returns:
        取引のみを含むDataFrame
    """
    trades = df[df["Signal"] != 0].copy()
    trades = trades[["Close", "Signal", "Trade", "Portfolio_Value"]]
    return trades


def calculate_trade_results(df: pd.DataFrame) -> list[dict]:
    """
    個別の取引結果を計算

    Args:
        df: バックテスト結果のDataFrame

    Returns:
        各取引の結果を含むリスト
    """
    trades = []
    position = None
    entry_price = None
    entry_date = None

    for date, row in df.iterrows():
        if row["Signal"] == 1:  # 買いシグナル
            if position is None:
                position = "long"
                entry_price = row["Close"]
                entry_date = date
        elif row["Signal"] == -1:  # 売りシグナル
            if position == "long":
                exit_price = row["Close"]
                pnl = (exit_price - entry_price) / entry_price * 100
                trades.append({
                    "entry_date": entry_date,
                    "exit_date": date,
                    "entry_price": entry_price,
                    "exit_price": exit_price,
                    "pnl_percent": pnl,
                    "win": pnl > 0
                })
                position = None
                entry_price = None

    return trades
