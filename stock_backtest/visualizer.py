"""
可視化モジュール: グラフ描画
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import japanize_matplotlib  # 日本語フォント対応
from config import MA_SHORT, MA_LONG


def plot_stock_chart(
    df: pd.DataFrame,
    ticker: str,
    stock_name: str = "",
    save_path: str = None
) -> None:
    """
    株価チャートと移動平均線、売買シグナルを描画

    Args:
        df: バックテスト結果のDataFrame
        ticker: 銘柄コード
        stock_name: 銘柄名
        save_path: 保存先パス（Noneの場合は表示のみ）
    """
    fig, axes = plt.subplots(2, 1, figsize=(14, 10), height_ratios=[2, 1])

    title = f"{stock_name} ({ticker})" if stock_name else ticker

    # 上段: 株価チャート
    ax1 = axes[0]
    ax1.plot(df.index, df["Close"], label="終値", color="black", linewidth=1)
    ax1.plot(df.index, df["MA_Short"], label=f"MA{MA_SHORT}", color="blue", linewidth=1)
    ax1.plot(df.index, df["MA_Long"], label=f"MA{MA_LONG}", color="red", linewidth=1)

    # 売買シグナルをプロット
    buy_signals = df[df["Signal"] == 1]
    sell_signals = df[df["Signal"] == -1]

    ax1.scatter(buy_signals.index, buy_signals["Close"],
                marker="^", color="green", s=100, label="買い", zorder=5)
    ax1.scatter(sell_signals.index, sell_signals["Close"],
                marker="v", color="red", s=100, label="売り", zorder=5)

    ax1.set_title(f"{title} - 株価チャート", fontsize=14)
    ax1.set_ylabel("株価（円）")
    ax1.legend(loc="upper left")
    ax1.grid(True, alpha=0.3)

    # 下段: 資産推移
    ax2 = axes[1]
    ax2.plot(df.index, df["Portfolio_Value"], label="戦略", color="blue", linewidth=1.5)
    ax2.plot(df.index, df["Buy_Hold_Value"], label="バイ&ホールド",
             color="gray", linewidth=1, linestyle="--")

    ax2.set_title("資産推移", fontsize=14)
    ax2.set_ylabel("資産（円）")
    ax2.set_xlabel("日付")
    ax2.legend(loc="upper left")
    ax2.grid(True, alpha=0.3)

    # Y軸のフォーマット
    ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f"¥{x:,.0f}"))

    plt.tight_layout()

    if save_path:
        plt.savefig(save_path, dpi=150, bbox_inches="tight")
        print(f"グラフを保存しました: {save_path}")
    else:
        plt.show()

    plt.close()


def plot_sector_comparison(
    comparison_df: pd.DataFrame,
    save_path: str = None
) -> None:
    """
    セクター別比較グラフを描画

    Args:
        comparison_df: セクター比較のDataFrame
        save_path: 保存先パス
    """
    if comparison_df.empty:
        print("比較データがありません")
        return

    fig, axes = plt.subplots(1, 3, figsize=(15, 5))

    sectors = comparison_df["セクター"]

    # リターン比較
    ax1 = axes[0]
    colors = ["green" if x >= 0 else "red" for x in comparison_df["平均リターン(%)"]]
    ax1.barh(sectors, comparison_df["平均リターン(%)"], color=colors)
    ax1.set_xlabel("平均リターン（%）")
    ax1.set_title("セクター別平均リターン")
    ax1.axvline(x=0, color="black", linewidth=0.5)

    # 勝率比較
    ax2 = axes[1]
    ax2.barh(sectors, comparison_df["平均勝率(%)"], color="steelblue")
    ax2.set_xlabel("平均勝率（%）")
    ax2.set_title("セクター別平均勝率")
    ax2.axvline(x=50, color="red", linewidth=0.5, linestyle="--")

    # シャープレシオ比較
    ax3 = axes[2]
    colors = ["green" if x >= 0 else "red" for x in comparison_df["平均シャープレシオ"]]
    ax3.barh(sectors, comparison_df["平均シャープレシオ"], color=colors)
    ax3.set_xlabel("平均シャープレシオ")
    ax3.set_title("セクター別シャープレシオ")
    ax3.axvline(x=0, color="black", linewidth=0.5)

    plt.tight_layout()

    if save_path:
        plt.savefig(save_path, dpi=150, bbox_inches="tight")
        print(f"グラフを保存しました: {save_path}")
    else:
        plt.show()

    plt.close()


def plot_multiple_stocks(
    results_list: list[dict],
    sector_name: str,
    save_path: str = None
) -> None:
    """
    複数銘柄の資産推移を比較

    Args:
        results_list: 分析結果のリスト（各要素にdf, ticker, stock_nameを含む）
        sector_name: セクター名
        save_path: 保存先パス
    """
    if not results_list:
        print("データがありません")
        return

    fig, ax = plt.subplots(figsize=(14, 8))

    for result in results_list:
        label = f"{result['stock_name']} ({result['ticker']})"
        # 正規化（初期値を1に）
        normalized = result["df"]["Portfolio_Value"] / result["df"]["Portfolio_Value"].iloc[0]
        ax.plot(result["df"].index, normalized, label=label, linewidth=1.5)

    ax.axhline(y=1, color="black", linewidth=0.5, linestyle="--")
    ax.set_title(f"{sector_name}セクター - 銘柄別パフォーマンス比較", fontsize=14)
    ax.set_ylabel("リターン（倍）")
    ax.set_xlabel("日付")
    ax.legend(loc="upper left", fontsize=9)
    ax.grid(True, alpha=0.3)

    plt.tight_layout()

    if save_path:
        plt.savefig(save_path, dpi=150, bbox_inches="tight")
        print(f"グラフを保存しました: {save_path}")
    else:
        plt.show()

    plt.close()


def plot_drawdown(df: pd.DataFrame, ticker: str, stock_name: str = "", save_path: str = None) -> None:
    """
    ドローダウンチャートを描画

    Args:
        df: バックテスト結果のDataFrame
        ticker: 銘柄コード
        stock_name: 銘柄名
        save_path: 保存先パス
    """
    title = f"{stock_name} ({ticker})" if stock_name else ticker

    # ドローダウン計算
    peak = df["Portfolio_Value"].expanding(min_periods=1).max()
    drawdown = (df["Portfolio_Value"] - peak) / peak * 100

    fig, ax = plt.subplots(figsize=(14, 5))

    ax.fill_between(df.index, drawdown, 0, color="red", alpha=0.3)
    ax.plot(df.index, drawdown, color="red", linewidth=1)

    ax.set_title(f"{title} - ドローダウン", fontsize=14)
    ax.set_ylabel("ドローダウン（%）")
    ax.set_xlabel("日付")
    ax.grid(True, alpha=0.3)

    plt.tight_layout()

    if save_path:
        plt.savefig(save_path, dpi=150, bbox_inches="tight")
        print(f"グラフを保存しました: {save_path}")
    else:
        plt.show()

    plt.close()
