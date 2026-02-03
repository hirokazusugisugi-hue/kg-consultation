"""
日本株バックテスト Webアプリケーション
テクニカル指標でスクリーニング → バックテスト
"""

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import japanize_matplotlib

from config import SECTORS, PERIOD_OPTIONS, INITIAL_CAPITAL
from data_fetcher import fetch_stock_data
from backtester import run_backtest, calculate_trade_results
from analyzer import calculate_max_drawdown, calculate_sharpe_ratio, calculate_profit_factor


# ページ設定
st.set_page_config(
    page_title="日本株バックテスト",
    page_icon="📈",
    layout="wide"
)


def apply_ma_crossover_strategy(df, ma_short, ma_long):
    """移動平均クロスオーバー戦略を適用"""
    df = df.copy()
    df["MA_Short"] = df["Close"].rolling(window=ma_short).mean()
    df["MA_Long"] = df["Close"].rolling(window=ma_long).mean()

    df["Signal"] = 0
    df["Position"] = np.where(df["MA_Short"] > df["MA_Long"], 1, -1)
    df["Signal"] = df["Position"].diff()
    df["Signal"] = df["Signal"].apply(lambda x: 1 if x > 0 else (-1 if x < 0 else 0))

    return df


def apply_rsi_strategy(df, rsi_period, oversold, overbought):
    """RSI戦略を適用"""
    df = df.copy()

    # RSI計算
    delta = df["Close"].diff()
    gain = (delta.where(delta > 0, 0)).rolling(window=rsi_period).mean()
    loss = (-delta.where(delta < 0, 0)).rolling(window=rsi_period).mean()
    rs = gain / loss
    df["RSI"] = 100 - (100 / (1 + rs))

    # シグナル生成
    df["Signal"] = 0
    df["Position"] = 0

    # RSIが売られすぎから回復 → 買い、買われすぎから下落 → 売り
    for i in range(1, len(df)):
        if df["RSI"].iloc[i-1] < oversold and df["RSI"].iloc[i] >= oversold:
            df.iloc[i, df.columns.get_loc("Signal")] = 1
        elif df["RSI"].iloc[i-1] > overbought and df["RSI"].iloc[i] <= overbought:
            df.iloc[i, df.columns.get_loc("Signal")] = -1

    # ポジション計算
    position = 0
    for i in range(len(df)):
        if df["Signal"].iloc[i] == 1:
            position = 1
        elif df["Signal"].iloc[i] == -1:
            position = -1
        df.iloc[i, df.columns.get_loc("Position")] = position

    return df


def apply_bollinger_strategy(df, bb_period, bb_std):
    """ボリンジャーバンド戦略を適用"""
    df = df.copy()

    df["BB_Middle"] = df["Close"].rolling(window=bb_period).mean()
    df["BB_Std"] = df["Close"].rolling(window=bb_period).std()
    df["BB_Upper"] = df["BB_Middle"] + (df["BB_Std"] * bb_std)
    df["BB_Lower"] = df["BB_Middle"] - (df["BB_Std"] * bb_std)

    df["Signal"] = 0
    df["Position"] = 0

    # 下限タッチで買い、上限タッチで売り
    for i in range(1, len(df)):
        if df["Close"].iloc[i] <= df["BB_Lower"].iloc[i]:
            df.iloc[i, df.columns.get_loc("Signal")] = 1
        elif df["Close"].iloc[i] >= df["BB_Upper"].iloc[i]:
            df.iloc[i, df.columns.get_loc("Signal")] = -1

    position = 0
    for i in range(len(df)):
        if df["Signal"].iloc[i] == 1:
            position = 1
        elif df["Signal"].iloc[i] == -1:
            position = -1
        df.iloc[i, df.columns.get_loc("Position")] = position

    return df


def apply_macd_strategy(df, fast_period, slow_period, signal_period):
    """MACD戦略を適用"""
    df = df.copy()

    df["EMA_Fast"] = df["Close"].ewm(span=fast_period, adjust=False).mean()
    df["EMA_Slow"] = df["Close"].ewm(span=slow_period, adjust=False).mean()
    df["MACD"] = df["EMA_Fast"] - df["EMA_Slow"]
    df["MACD_Signal"] = df["MACD"].ewm(span=signal_period, adjust=False).mean()
    df["MACD_Hist"] = df["MACD"] - df["MACD_Signal"]

    df["Signal"] = 0
    df["Position"] = np.where(df["MACD"] > df["MACD_Signal"], 1, -1)
    df["Signal"] = df["Position"].diff()
    df["Signal"] = df["Signal"].apply(lambda x: 1 if x > 0 else (-1 if x < 0 else 0))

    return df


def run_backtest_with_strategy(df, initial_capital=INITIAL_CAPITAL):
    """バックテストを実行"""
    df = df.copy()
    df = df.dropna()

    if df.empty:
        return df

    df["Returns"] = df["Close"].pct_change()
    df["Strategy_Returns"] = df["Position"].shift(1) * df["Returns"]
    df["Cumulative_Returns"] = (1 + df["Returns"]).cumprod()
    df["Cumulative_Strategy_Returns"] = (1 + df["Strategy_Returns"]).cumprod()
    df["Portfolio_Value"] = initial_capital * df["Cumulative_Strategy_Returns"]
    df["Buy_Hold_Value"] = initial_capital * df["Cumulative_Returns"]

    return df


def analyze_backtest_results(df, ticker, stock_name):
    """バックテスト結果を分析"""
    df = df.dropna()

    if df.empty or "Strategy_Returns" not in df.columns:
        return None

    trades = calculate_trade_results(df)
    num_trades = len(trades)
    wins = sum(1 for t in trades if t["win"])
    win_rate = wins / num_trades * 100 if num_trades > 0 else 0

    total_return = (df["Cumulative_Strategy_Returns"].iloc[-1] - 1) * 100
    buy_hold_return = (df["Cumulative_Returns"].iloc[-1] - 1) * 100
    max_drawdown = calculate_max_drawdown(df["Portfolio_Value"])
    sharpe_ratio = calculate_sharpe_ratio(df["Strategy_Returns"].dropna())
    profit_factor = calculate_profit_factor(trades)

    final_value = df["Portfolio_Value"].iloc[-1]

    return {
        "ticker": ticker,
        "stock_name": stock_name,
        "num_trades": num_trades,
        "wins": wins,
        "win_rate": win_rate,
        "total_return": total_return,
        "buy_hold_return": buy_hold_return,
        "excess_return": total_return - buy_hold_return,
        "max_drawdown": max_drawdown,
        "sharpe_ratio": sharpe_ratio,
        "profit_factor": profit_factor,
        "final_value": final_value,
        "profit_loss": final_value - INITIAL_CAPITAL,
        "df": df
    }


def plot_chart(df, ticker, stock_name, strategy_type, params):
    """チャートを描画"""
    fig, axes = plt.subplots(3, 1, figsize=(12, 10), height_ratios=[2, 1, 1])

    # 上段: 株価チャート
    ax1 = axes[0]
    ax1.plot(df.index, df["Close"], label="終値", color="black", linewidth=1)

    if strategy_type == "移動平均クロスオーバー":
        ax1.plot(df.index, df["MA_Short"], label=f"MA{params['ma_short']}", color="blue", linewidth=1)
        ax1.plot(df.index, df["MA_Long"], label=f"MA{params['ma_long']}", color="red", linewidth=1)
    elif strategy_type == "ボリンジャーバンド":
        ax1.plot(df.index, df["BB_Middle"], label="中央線", color="blue", linewidth=1)
        ax1.fill_between(df.index, df["BB_Upper"], df["BB_Lower"], alpha=0.2, color="blue")

    buy_signals = df[df["Signal"] == 1]
    sell_signals = df[df["Signal"] == -1]
    ax1.scatter(buy_signals.index, buy_signals["Close"], marker="^", color="green", s=100, label="買い", zorder=5)
    ax1.scatter(sell_signals.index, sell_signals["Close"], marker="v", color="red", s=100, label="売り", zorder=5)

    ax1.set_title(f"{stock_name} ({ticker}) - {strategy_type}", fontsize=14)
    ax1.set_ylabel("株価（円）")
    ax1.legend(loc="upper left")
    ax1.grid(True, alpha=0.3)

    # 中段: テクニカル指標
    ax2 = axes[1]
    if strategy_type == "RSI":
        ax2.plot(df.index, df["RSI"], label="RSI", color="purple", linewidth=1)
        ax2.axhline(y=params["oversold"], color="green", linestyle="--", label=f"売られすぎ({params['oversold']})")
        ax2.axhline(y=params["overbought"], color="red", linestyle="--", label=f"買われすぎ({params['overbought']})")
        ax2.set_ylabel("RSI")
        ax2.set_ylim(0, 100)
    elif strategy_type == "MACD":
        ax2.plot(df.index, df["MACD"], label="MACD", color="blue", linewidth=1)
        ax2.plot(df.index, df["MACD_Signal"], label="シグナル", color="red", linewidth=1)
        ax2.bar(df.index, df["MACD_Hist"], label="ヒストグラム", color="gray", alpha=0.5)
        ax2.axhline(y=0, color="black", linewidth=0.5)
        ax2.set_ylabel("MACD")
    else:
        ax2.plot(df.index, df["Portfolio_Value"], label="戦略", color="blue", linewidth=1)
        ax2.plot(df.index, df["Buy_Hold_Value"], label="B&H", color="gray", linestyle="--")
        ax2.set_ylabel("資産（円）")
        ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f"¥{x:,.0f}"))
    ax2.legend(loc="upper left")
    ax2.grid(True, alpha=0.3)

    # 下段: 資産推移
    ax3 = axes[2]
    ax3.plot(df.index, df["Portfolio_Value"], label="戦略", color="blue", linewidth=1.5)
    ax3.plot(df.index, df["Buy_Hold_Value"], label="バイ&ホールド", color="gray", linewidth=1, linestyle="--")
    ax3.set_title("資産推移", fontsize=12)
    ax3.set_ylabel("資産（円）")
    ax3.set_xlabel("日付")
    ax3.legend(loc="upper left")
    ax3.grid(True, alpha=0.3)
    ax3.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f"¥{x:,.0f}"))

    plt.tight_layout()
    return fig


def main():
    # ヘッダー
    st.title("📈 日本株バックテスト - テクニカル戦略スクリーニング")

    # 使い方ガイド
    with st.expander("📖 **使い方ガイド** - クリックして開く", expanded=True):
        st.markdown("""
        ### このアプリでできること
        **テクニカル指標のパラメータを設定** → **全銘柄でバックテスト** → **条件に合う銘柄を発見**

        | ステップ | 内容 |
        |---------|------|
        | 1️⃣ 戦略を選択 | 移動平均/RSI/ボリンジャーバンド/MACD から選択 |
        | 2️⃣ パラメータ設定 | 各戦略のパラメータを調整 |
        | 3️⃣ フィルター設定 | 勝率・リターン・ドローダウンの条件を設定 |
        | 4️⃣ スクリーニング実行 | 全25銘柄をバックテストして条件に合うものを抽出 |
        | 5️⃣ 結果確認 | ランキング表示・個別チャート確認 |

        ---

        ### 戦略の説明
        | 戦略 | 買いシグナル | 売りシグナル |
        |-----|------------|------------|
        | **移動平均クロスオーバー** | 短期MAが長期MAを上抜け | 短期MAが長期MAを下抜け |
        | **RSI** | RSIが売られすぎ水準から回復 | RSIが買われすぎ水準から下落 |
        | **ボリンジャーバンド** | 株価が下限バンドにタッチ | 株価が上限バンドにタッチ |
        | **MACD** | MACDがシグナル線を上抜け | MACDがシグナル線を下抜け |
        """)

    st.divider()

    # === サイドバー: 戦略設定 ===
    with st.sidebar:
        st.header("⚙️ 戦略設定")

        # 戦略選択
        strategy_type = st.selectbox(
            "テクニカル戦略",
            ["移動平均クロスオーバー", "RSI", "ボリンジャーバンド", "MACD"]
        )

        st.divider()

        # パラメータ設定
        params = {}

        if strategy_type == "移動平均クロスオーバー":
            st.subheader("📊 移動平均パラメータ")
            params["ma_short"] = st.slider("短期MA（日）", 3, 30, 5)
            params["ma_long"] = st.slider("長期MA（日）", 10, 100, 25)

        elif strategy_type == "RSI":
            st.subheader("📊 RSIパラメータ")
            params["rsi_period"] = st.slider("RSI期間（日）", 7, 28, 14)
            params["oversold"] = st.slider("売られすぎ水準", 10, 40, 30)
            params["overbought"] = st.slider("買われすぎ水準", 60, 90, 70)

        elif strategy_type == "ボリンジャーバンド":
            st.subheader("📊 ボリンジャーバンドパラメータ")
            params["bb_period"] = st.slider("期間（日）", 10, 30, 20)
            params["bb_std"] = st.slider("標準偏差倍率", 1.0, 3.0, 2.0, 0.5)

        elif strategy_type == "MACD":
            st.subheader("📊 MACDパラメータ")
            params["fast_period"] = st.slider("短期EMA（日）", 5, 20, 12)
            params["slow_period"] = st.slider("長期EMA（日）", 15, 40, 26)
            params["signal_period"] = st.slider("シグナル期間（日）", 5, 15, 9)

        st.divider()

        # 期間選択
        st.subheader("📅 分析期間")
        period_label = st.selectbox(
            "期間",
            [label for _, label in PERIOD_OPTIONS.values()]
        )
        period = next(code for code, label in PERIOD_OPTIONS.values() if label == period_label)

        st.divider()

        # フィルター設定
        st.subheader("🔍 スクリーニング条件")
        min_win_rate = st.slider("最低勝率（%）", 0, 80, 40)
        min_return = st.slider("最低リターン（%）", -50, 50, 0)
        max_drawdown = st.slider("最大許容DD（%）", -80, 0, -30)
        min_trades = st.slider("最低取引回数", 1, 20, 3)

    # === メインエリア ===

    # セクター選択
    st.subheader("🏢 対象セクター")
    sector_options = list(SECTORS.keys())
    selected_sectors = st.multiselect(
        "分析するセクターを選択（複数選択可）",
        sector_options,
        default=sector_options
    )

    # 実行ボタン
    if st.button("🚀 スクリーニング実行", type="primary", use_container_width=True):

        # 対象銘柄を収集
        target_stocks = []
        for sector in selected_sectors:
            for ticker, name in SECTORS[sector].items():
                target_stocks.append({"ticker": ticker, "name": name, "sector": sector})

        if not target_stocks:
            st.warning("セクターを選択してください")
            return

        # プログレス表示
        progress_bar = st.progress(0)
        status_text = st.empty()

        all_results = []

        for i, stock in enumerate(target_stocks):
            status_text.text(f"分析中: {stock['name']} ({stock['ticker']})...")

            # データ取得
            df = fetch_stock_data(stock["ticker"], period)
            if df is None:
                progress_bar.progress((i + 1) / len(target_stocks))
                continue

            # 戦略適用
            if strategy_type == "移動平均クロスオーバー":
                df = apply_ma_crossover_strategy(df, params["ma_short"], params["ma_long"])
            elif strategy_type == "RSI":
                df = apply_rsi_strategy(df, params["rsi_period"], params["oversold"], params["overbought"])
            elif strategy_type == "ボリンジャーバンド":
                df = apply_bollinger_strategy(df, params["bb_period"], params["bb_std"])
            elif strategy_type == "MACD":
                df = apply_macd_strategy(df, params["fast_period"], params["slow_period"], params["signal_period"])

            # バックテスト
            df = run_backtest_with_strategy(df)

            # 分析
            result = analyze_backtest_results(df, stock["ticker"], stock["name"])
            if result:
                result["sector"] = stock["sector"]
                all_results.append(result)

            progress_bar.progress((i + 1) / len(target_stocks))

        status_text.empty()
        progress_bar.empty()

        if not all_results:
            st.error("分析結果がありません")
            return

        # フィルタリング
        filtered_results = [
            r for r in all_results
            if r["win_rate"] >= min_win_rate
            and r["total_return"] >= min_return
            and r["max_drawdown"] >= max_drawdown
            and r["num_trades"] >= min_trades
        ]

        # 結果表示
        st.header("📊 スクリーニング結果")

        col1, col2, col3 = st.columns(3)
        col1.metric("分析銘柄数", f"{len(all_results)}銘柄")
        col2.metric("条件適合銘柄", f"{len(filtered_results)}銘柄")
        col3.metric("適合率", f"{len(filtered_results)/len(all_results)*100:.1f}%")

        st.divider()

        if filtered_results:
            # ランキング表示
            st.subheader("🏆 条件適合銘柄ランキング（リターン順）")

            sorted_results = sorted(filtered_results, key=lambda x: x["total_return"], reverse=True)

            ranking_data = []
            for i, r in enumerate(sorted_results, 1):
                ranking_data.append({
                    "順位": i,
                    "銘柄": r["stock_name"],
                    "セクター": r["sector"],
                    "リターン": f"{r['total_return']:+.1f}%",
                    "勝率": f"{r['win_rate']:.1f}%",
                    "取引数": r["num_trades"],
                    "最大DD": f"{r['max_drawdown']:.1f}%",
                    "シャープ": f"{r['sharpe_ratio']:.2f}",
                    "損益": f"¥{r['profit_loss']:+,.0f}"
                })

            st.dataframe(pd.DataFrame(ranking_data), use_container_width=True, hide_index=True)

            # 個別チャート
            st.subheader("📈 個別チャート")

            # セッションステートに結果を保存
            st.session_state["results"] = sorted_results
            st.session_state["strategy_type"] = strategy_type
            st.session_state["params"] = params

            selected_stock = st.selectbox(
                "銘柄を選択してチャートを表示",
                [f"{r['stock_name']} ({r['ticker']}) - リターン: {r['total_return']:+.1f}%" for r in sorted_results]
            )

            # 選択された銘柄のチャートを表示
            selected_idx = [f"{r['stock_name']} ({r['ticker']}) - リターン: {r['total_return']:+.1f}%" for r in sorted_results].index(selected_stock)
            selected_result = sorted_results[selected_idx]

            # メトリクス表示
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("戦略リターン", f"{selected_result['total_return']:+.1f}%")
            col2.metric("勝率", f"{selected_result['win_rate']:.1f}%")
            col3.metric("最大DD", f"{selected_result['max_drawdown']:.1f}%")
            col4.metric("損益", f"¥{selected_result['profit_loss']:+,.0f}")

            # チャート描画
            fig = plot_chart(
                selected_result["df"],
                selected_result["ticker"],
                selected_result["stock_name"],
                strategy_type,
                params
            )
            st.pyplot(fig)
            plt.close()

        else:
            st.warning("条件に適合する銘柄がありませんでした。フィルター条件を緩めてください。")

            # 全結果を参考表示
            st.subheader("📋 全銘柄の結果（参考）")
            sorted_all = sorted(all_results, key=lambda x: x["total_return"], reverse=True)

            all_data = []
            for r in sorted_all:
                all_data.append({
                    "銘柄": r["stock_name"],
                    "セクター": r["sector"],
                    "リターン": f"{r['total_return']:+.1f}%",
                    "勝率": f"{r['win_rate']:.1f}%",
                    "取引数": r["num_trades"],
                    "最大DD": f"{r['max_drawdown']:.1f}%"
                })

            st.dataframe(pd.DataFrame(all_data), use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
