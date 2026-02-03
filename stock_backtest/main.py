"""
日本株バックテストアプリケーション
移動平均クロスオーバー戦略によるバックテスト
"""

import sys
from config import SECTORS, PERIOD_OPTIONS, MA_SHORT, MA_LONG, INITIAL_CAPITAL
from data_fetcher import fetch_stock_data, fetch_multiple_stocks
from strategy import apply_strategy
from backtester import run_backtest, get_trades
from analyzer import analyze_results, print_analysis, compare_sectors
from visualizer import (
    plot_stock_chart,
    plot_sector_comparison,
    plot_multiple_stocks,
    plot_drawdown,
)


def print_header():
    """ヘッダーを表示"""
    print("\n" + "=" * 60)
    print("   日本株バックテストアプリケーション")
    print("   移動平均クロスオーバー戦略")
    print(f"   （短期MA: {MA_SHORT}日 / 長期MA: {MA_LONG}日）")
    print("=" * 60)
    print_usage()


def print_usage():
    """使い方を表示"""
    print("""
┌─────────────────────────────────────────────────────────┐
│                      使い方                              │
├─────────────────────────────────────────────────────────┤
│                                                         │
│  1. セクターを選択                                       │
│     → 自動車/電機/銀行/通信/小売 から選択               │
│     → 「全セクター分析」で一括比較も可能                │
│                                                         │
│  2. 分析期間を選択                                       │
│     → 6ヶ月/1年/2年/5年 から選択                        │
│                                                         │
│  3. 銘柄を選択                                           │
│     → 個別銘柄 または セクター内全銘柄                  │
│                                                         │
│  4. 結果を確認                                           │
│     → 勝率、リターン、最大ドローダウン等を表示          │
│     → 取引履歴やグラフも表示可能                        │
│                                                         │
├─────────────────────────────────────────────────────────┤
│  【戦略の説明】                                          │
│   ゴールデンクロス（MA5がMA25を上抜け）→ 買い           │
│   デッドクロス（MA5がMA25を下抜け）→ 売り               │
│                                                         │
│  【分析指標】                                            │
│   ・勝率: 利益が出た取引の割合                           │
│   ・最大ドローダウン: 最大の資産減少率                   │
│   ・シャープレシオ: リスク調整後リターン                 │
│   ・プロフィットファクター: 総利益÷総損失               │
│                                                         │
│  【操作方法】                                            │
│   ・数字を入力して選択                                   │
│   ・Enterでデフォルト値を使用                            │
│   ・Ctrl+C で終了                                        │
│                                                         │
└─────────────────────────────────────────────────────────┘
""")


def select_sector() -> tuple[str, dict]:
    """セクターを選択"""
    print("\n【セクター選択】")
    sectors = list(SECTORS.keys())
    for i, sector in enumerate(sectors, 1):
        print(f"  {i}. {sector}")
    print(f"  {len(sectors) + 1}. 全セクター分析")
    print("  0. 終了")

    while True:
        try:
            choice = input("\n選択してください: ").strip()
            if choice == "0":
                return None, None
            choice = int(choice)
            if 1 <= choice <= len(sectors):
                sector = sectors[choice - 1]
                return sector, SECTORS[sector]
            elif choice == len(sectors) + 1:
                return "全セクター", None
        except ValueError:
            pass
        print("無効な選択です。もう一度入力してください。")


def select_period() -> str:
    """分析期間を選択"""
    print("\n【期間選択】")
    for key, (_, label) in PERIOD_OPTIONS.items():
        print(f"  {key}. {label}")

    while True:
        choice = input("\n選択してください [デフォルト: 2]: ").strip()
        if choice == "":
            choice = "2"
        if choice in PERIOD_OPTIONS:
            return PERIOD_OPTIONS[choice][0]
        print("無効な選択です。もう一度入力してください。")


def select_stock(stocks: dict) -> tuple[str, str]:
    """銘柄を選択"""
    print("\n【銘柄選択】")
    tickers = list(stocks.keys())
    for i, (ticker, name) in enumerate(stocks.items(), 1):
        print(f"  {i}. {name} ({ticker})")
    print(f"  {len(tickers) + 1}. 全銘柄を分析")

    while True:
        try:
            choice = input("\n選択してください: ").strip()
            choice = int(choice)
            if 1 <= choice <= len(tickers):
                ticker = tickers[choice - 1]
                return ticker, stocks[ticker]
            elif choice == len(tickers) + 1:
                return "all", None
        except ValueError:
            pass
        print("無効な選択です。もう一度入力してください。")


def analyze_single_stock(ticker: str, stock_name: str, period: str) -> dict:
    """単一銘柄を分析"""
    print(f"\n{stock_name} ({ticker}) を分析中...")

    # データ取得
    df = fetch_stock_data(ticker, period)
    if df is None:
        return None

    # 戦略適用
    df = apply_strategy(df)

    # バックテスト実行
    df = run_backtest(df)

    # 分析
    results = analyze_results(df, ticker)
    results["df"] = df
    results["stock_name"] = stock_name

    return results


def analyze_sector_stocks(sector_name: str, stocks: dict, period: str) -> list[dict]:
    """セクター内の全銘柄を分析"""
    print(f"\n【{sector_name}セクター分析】")
    print("=" * 40)

    results_list = []
    for ticker, stock_name in stocks.items():
        result = analyze_single_stock(ticker, stock_name, period)
        if result:
            results_list.append(result)
            print_analysis(result, stock_name)

    return results_list


def run_all_sectors_analysis(period: str):
    """全セクター分析を実行"""
    print("\n【全セクター分析】")
    print("=" * 60)

    sector_results = {}

    for sector_name, stocks in SECTORS.items():
        print(f"\n--- {sector_name}セクター ---")
        results = []
        for ticker, stock_name in stocks.items():
            result = analyze_single_stock(ticker, stock_name, period)
            if result:
                results.append(result)
        sector_results[sector_name] = results

    # セクター比較
    comparison_df = compare_sectors(
        {k: [r for r in v] for k, v in sector_results.items()}
    )

    print("\n" + "=" * 60)
    print("【セクター別比較】")
    print("=" * 60)
    print(comparison_df.to_string(index=False))

    # グラフ表示
    show_graph = input("\nセクター比較グラフを表示しますか？ (y/n) [y]: ").strip().lower()
    if show_graph != "n":
        plot_sector_comparison(comparison_df)


def show_graphs_menu(results: dict):
    """グラフ表示メニュー"""
    print("\n【グラフ表示】")
    print("  1. 株価チャート（移動平均線・売買シグナル）")
    print("  2. ドローダウンチャート")
    print("  0. スキップ")

    while True:
        choice = input("\n選択してください: ").strip()
        if choice == "0":
            break
        elif choice == "1":
            plot_stock_chart(
                results["df"],
                results["ticker"],
                results["stock_name"]
            )
        elif choice == "2":
            plot_drawdown(
                results["df"],
                results["ticker"],
                results["stock_name"]
            )
        else:
            print("無効な選択です。")
            continue

        another = input("\n他のグラフを表示しますか？ (y/n) [n]: ").strip().lower()
        if another != "y":
            break


def show_sector_graphs(results_list: list[dict], sector_name: str):
    """セクター分析のグラフ表示"""
    show = input("\nセクター比較グラフを表示しますか？ (y/n) [y]: ").strip().lower()
    if show != "n":
        plot_multiple_stocks(results_list, sector_name)


def main():
    """メイン関数"""
    print_header()

    while True:
        # セクター選択
        sector_name, stocks = select_sector()
        if sector_name is None:
            print("\nアプリケーションを終了します。")
            break

        # 期間選択
        period = select_period()

        if sector_name == "全セクター":
            # 全セクター分析
            run_all_sectors_analysis(period)
        else:
            # 銘柄選択
            ticker, stock_name = select_stock(stocks)

            if ticker == "all":
                # セクター内全銘柄分析
                results_list = analyze_sector_stocks(sector_name, stocks, period)
                if results_list:
                    show_sector_graphs(results_list, sector_name)
            else:
                # 単一銘柄分析
                results = analyze_single_stock(ticker, stock_name, period)
                if results:
                    print_analysis(results, stock_name)

                    # 取引履歴表示
                    show_trades = input("\n取引履歴を表示しますか？ (y/n) [n]: ").strip().lower()
                    if show_trades == "y":
                        trades_df = get_trades(results["df"])
                        print("\n【取引履歴】")
                        print(trades_df.to_string())

                    # グラフ表示
                    show_graphs_menu(results)

        # 続行確認
        print("\n" + "-" * 40)
        continue_choice = input("別の分析を行いますか？ (y/n) [y]: ").strip().lower()
        if continue_choice == "n":
            print("\nアプリケーションを終了します。")
            break


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n中断されました。アプリケーションを終了します。")
        sys.exit(0)
