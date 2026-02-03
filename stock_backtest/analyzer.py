"""
分析モジュール: バックテスト結果の統計分析
"""

import pandas as pd
import numpy as np
from backtester import calculate_trade_results


def calculate_max_drawdown(portfolio_values: pd.Series) -> float:
    """
    最大ドローダウンを計算

    Args:
        portfolio_values: ポートフォリオ価値の時系列

    Returns:
        最大ドローダウン（パーセント）
    """
    peak = portfolio_values.expanding(min_periods=1).max()
    drawdown = (portfolio_values - peak) / peak * 100
    return drawdown.min()


def calculate_sharpe_ratio(
    returns: pd.Series,
    risk_free_rate: float = 0.0,
    periods_per_year: int = 252
) -> float:
    """
    シャープレシオを計算

    Args:
        returns: 日次リターンの時系列
        risk_free_rate: 年間リスクフリーレート
        periods_per_year: 年間取引日数

    Returns:
        シャープレシオ
    """
    if returns.std() == 0:
        return 0.0

    excess_returns = returns - risk_free_rate / periods_per_year
    return np.sqrt(periods_per_year) * excess_returns.mean() / returns.std()


def calculate_profit_factor(trades: list[dict]) -> float:
    """
    プロフィットファクターを計算

    Args:
        trades: 取引結果のリスト

    Returns:
        プロフィットファクター（総利益 / 総損失）
    """
    gross_profit = sum(t["pnl_percent"] for t in trades if t["pnl_percent"] > 0)
    gross_loss = abs(sum(t["pnl_percent"] for t in trades if t["pnl_percent"] < 0))

    if gross_loss == 0:
        return float("inf") if gross_profit > 0 else 0.0

    return gross_profit / gross_loss


def analyze_results(df: pd.DataFrame, ticker: str = "") -> dict:
    """
    バックテスト結果を分析

    Args:
        df: バックテスト結果のDataFrame
        ticker: 銘柄コード

    Returns:
        分析結果の辞書
    """
    if df.empty or "Strategy_Returns" not in df.columns:
        return {}

    df = df.dropna()

    # 取引結果を計算
    trades = calculate_trade_results(df)
    num_trades = len(trades)
    wins = sum(1 for t in trades if t["win"])
    win_rate = wins / num_trades * 100 if num_trades > 0 else 0

    # 各種指標を計算
    total_return = (df["Cumulative_Strategy_Returns"].iloc[-1] - 1) * 100
    buy_hold_return = (df["Cumulative_Returns"].iloc[-1] - 1) * 100
    max_drawdown = calculate_max_drawdown(df["Portfolio_Value"])
    sharpe_ratio = calculate_sharpe_ratio(df["Strategy_Returns"].dropna())
    profit_factor = calculate_profit_factor(trades)

    # 最終資産
    final_value = df["Portfolio_Value"].iloc[-1]
    initial_value = df["Portfolio_Value"].iloc[0] / df["Cumulative_Strategy_Returns"].iloc[0]

    return {
        "ticker": ticker,
        "period": f"{df.index[0].strftime('%Y-%m-%d')} ~ {df.index[-1].strftime('%Y-%m-%d')}",
        "num_trades": num_trades,
        "wins": wins,
        "losses": num_trades - wins,
        "win_rate": win_rate,
        "total_return": total_return,
        "buy_hold_return": buy_hold_return,
        "max_drawdown": max_drawdown,
        "sharpe_ratio": sharpe_ratio,
        "profit_factor": profit_factor,
        "initial_value": initial_value,
        "final_value": final_value,
        "trades": trades,
    }


def print_analysis(results: dict, stock_name: str = "") -> None:
    """
    分析結果をコンソールに出力

    Args:
        results: 分析結果の辞書
        stock_name: 銘柄名
    """
    if not results:
        print("分析結果がありません")
        return

    title = f"{stock_name} ({results['ticker']})" if stock_name else results["ticker"]
    print("\n" + "=" * 60)
    print(f"【バックテスト結果】{title}")
    print("=" * 60)
    print(f"分析期間: {results['period']}")
    print("-" * 60)
    print(f"取引回数: {results['num_trades']}回")
    print(f"勝ち: {results['wins']}回 / 負け: {results['losses']}回")
    print(f"勝率: {results['win_rate']:.1f}%")
    print("-" * 60)
    print(f"戦略リターン: {results['total_return']:+.2f}%")
    print(f"バイ&ホールド: {results['buy_hold_return']:+.2f}%")
    print(f"超過リターン: {results['total_return'] - results['buy_hold_return']:+.2f}%")
    print("-" * 60)
    print(f"最大ドローダウン: {results['max_drawdown']:.2f}%")
    print(f"シャープレシオ: {results['sharpe_ratio']:.2f}")
    print(f"プロフィットファクター: {results['profit_factor']:.2f}")
    print("-" * 60)
    print(f"初期資金: ¥{results['initial_value']:,.0f}")
    print(f"最終資産: ¥{results['final_value']:,.0f}")
    print(f"損益: ¥{results['final_value'] - results['initial_value']:+,.0f}")
    print("=" * 60)


def compare_sectors(sector_results: dict[str, list[dict]]) -> pd.DataFrame:
    """
    セクター間の比較

    Args:
        sector_results: セクター名をキー、分析結果リストを値とする辞書

    Returns:
        比較結果のDataFrame
    """
    comparison = []

    for sector, results in sector_results.items():
        if not results:
            continue

        avg_return = np.mean([r["total_return"] for r in results])
        avg_win_rate = np.mean([r["win_rate"] for r in results])
        avg_sharpe = np.mean([r["sharpe_ratio"] for r in results])

        comparison.append({
            "セクター": sector,
            "平均リターン(%)": avg_return,
            "平均勝率(%)": avg_win_rate,
            "平均シャープレシオ": avg_sharpe,
            "銘柄数": len(results),
        })

    return pd.DataFrame(comparison)
