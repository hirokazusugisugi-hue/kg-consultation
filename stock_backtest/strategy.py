"""
戦略モジュール: 移動平均クロスオーバー戦略
"""

import pandas as pd
import numpy as np
from config import MA_SHORT, MA_LONG


def calculate_moving_averages(
    df: pd.DataFrame,
    short_window: int = MA_SHORT,
    long_window: int = MA_LONG
) -> pd.DataFrame:
    """
    移動平均を計算してDataFrameに追加

    Args:
        df: 株価データ（Close列が必要）
        short_window: 短期移動平均の期間
        long_window: 長期移動平均の期間

    Returns:
        MA_Short, MA_Long列が追加されたDataFrame
    """
    df = df.copy()
    df["MA_Short"] = df["Close"].rolling(window=short_window).mean()
    df["MA_Long"] = df["Close"].rolling(window=long_window).mean()
    return df


def generate_signals(df: pd.DataFrame) -> pd.DataFrame:
    """
    移動平均クロスオーバーに基づいて売買シグナルを生成

    ゴールデンクロス（短期MAが長期MAを上抜け）→ 買いシグナル (1)
    デッドクロス（短期MAが長期MAを下抜け）→ 売りシグナル (-1)

    Args:
        df: MA_Short, MA_Long列を含むDataFrame

    Returns:
        Signal, Position列が追加されたDataFrame
    """
    df = df.copy()

    # シグナル生成（0: ホールド, 1: 買い, -1: 売り）
    df["Signal"] = 0

    # 短期MAが長期MAを上回っている場合は1、下回っている場合は-1
    df["Position"] = np.where(df["MA_Short"] > df["MA_Long"], 1, -1)

    # クロスオーバーポイントを検出
    df["Signal"] = df["Position"].diff()

    # シグナルを正規化（2→1, -2→-1）
    df["Signal"] = df["Signal"].apply(lambda x: 1 if x > 0 else (-1 if x < 0 else 0))

    return df


def apply_strategy(
    df: pd.DataFrame,
    short_window: int = MA_SHORT,
    long_window: int = MA_LONG
) -> pd.DataFrame:
    """
    移動平均クロスオーバー戦略を適用

    Args:
        df: 株価データ
        short_window: 短期移動平均の期間
        long_window: 長期移動平均の期間

    Returns:
        戦略が適用されたDataFrame
    """
    df = calculate_moving_averages(df, short_window, long_window)
    df = generate_signals(df)
    return df
