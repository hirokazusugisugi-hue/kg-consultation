"""
デモスクリプト: 自動車セクターのバックテストを自動実行
"""

import sys
import matplotlib
matplotlib.use('Agg')  # 非対話型バックエンド

from config import SECTORS, MA_SHORT, MA_LONG, INITIAL_CAPITAL
from data_fetcher import fetch_stock_data
from strategy import apply_strategy
from backtester import run_backtest
from analyzer import analyze_results, print_analysis, compare_sectors
from visualizer import plot_stock_chart, plot_multiple_stocks


def run_demo():
    """デモ実行: 自動車セクターの分析"""
    print("\n" + "=" * 60)
    print("   日本株バックテストアプリケーション - デモ")
    print("   移動平均クロスオーバー戦略")
    print(f"   （短期MA: {MA_SHORT}日 / 長期MA: {MA_LONG}日）")
    print("=" * 60)

    sector_name = "自動車"
    stocks = SECTORS[sector_name]
    period = "1y"

    print(f"\n【{sector_name}セクター - 1年間のバックテスト】")
    print("=" * 60)

    results_list = []

    for ticker, stock_name in stocks.items():
        print(f"\n{stock_name} ({ticker}) を分析中...")

        # データ取得
        df = fetch_stock_data(ticker, period)
        if df is None:
            continue

        # 戦略適用
        df = apply_strategy(df)

        # バックテスト実行
        df = run_backtest(df)

        # 分析
        results = analyze_results(df, ticker)
        results["df"] = df
        results["stock_name"] = stock_name

        # 結果出力
        print_analysis(results, stock_name)

        # グラフ保存
        chart_path = f"{ticker.replace('.', '_')}_chart.png"
        plot_stock_chart(df, ticker, stock_name, save_path=chart_path)

        results_list.append(results)

    # セクター比較グラフ
    if results_list:
        print("\n" + "=" * 60)
        print("【セクター内比較グラフを生成中...】")
        plot_multiple_stocks(results_list, sector_name, save_path="sector_comparison.png")

    print("\n" + "=" * 60)
    print("デモ完了!")
    print("生成されたグラフファイル:")
    for ticker in stocks.keys():
        print(f"  - {ticker.replace('.', '_')}_chart.png")
    print("  - sector_comparison.png")
    print("=" * 60)


if __name__ == "__main__":
    run_demo()
