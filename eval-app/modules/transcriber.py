"""
音声文字起こしモジュール（faster-whisper）

ローカル環境で高速・プライベートな日本語文字起こしを提供する。
Apple Silicon MPS 非対応のため CPU + int8 量子化で実行。

前提条件:
    pip install faster-whisper
    brew install ffmpeg
"""

import os
import tempfile
from pathlib import Path
from typing import Callable, Optional

from faster_whisper import WhisperModel

DEFAULT_MODEL_SIZE = os.environ.get("WHISPER_MODEL", "large-v3")

# シングルトンキャッシュ（モデルロードは30-90秒かかるため再利用）
_model_cache: dict = {}


def get_model(model_size: str = None) -> WhisperModel:
    """WhisperModel のシングルトンインスタンスを返す。"""
    model_size = model_size or DEFAULT_MODEL_SIZE

    if model_size not in _model_cache:
        _model_cache[model_size] = WhisperModel(
            model_size,
            device="cpu",
            compute_type="int8",
        )

    return _model_cache[model_size]


def transcribe_audio(
    audio_path: str,
    model_size: str = None,
    progress_callback: Optional[Callable[[str], None]] = None,
) -> str:
    """
    音声ファイルを文字起こしする。

    Args:
        audio_path: 音声ファイルパス（.wav, .mp3, .m4a 等）
        model_size: Whisper モデルサイズ（デフォルト: WHISPER_MODEL環境変数 or "large-v3"）
        progress_callback: 進捗通知用コールバック

    Returns:
        文字起こしテキスト
    """
    path = Path(audio_path)
    if not path.exists():
        raise FileNotFoundError(f"音声ファイルが見つかりません: {audio_path}")

    if progress_callback:
        progress_callback("Whisperモデルをロード中...")

    model = get_model(model_size)

    if progress_callback:
        progress_callback("文字起こしを実行中...（音声の長さに応じて数分かかります）")

    segments, info = model.transcribe(
        str(path),
        language="ja",
        beam_size=5,
        vad_filter=True,
        vad_parameters=dict(min_silence_duration_ms=500),
    )

    if progress_callback:
        progress_callback(
            f"言語: {info.language} (確率: {info.language_probability:.1%}), "
            f"音声長: {info.duration:.0f}秒"
        )

    lines = []
    for segment in segments:
        text = segment.text.strip()
        if text:
            lines.append(text)

    transcript = "\n".join(lines)

    if not transcript.strip():
        raise RuntimeError("文字起こし結果が空です。音声ファイルを確認してください。")

    if progress_callback:
        progress_callback(f"文字起こし完了: {len(transcript):,}文字")

    return transcript


def transcribe_uploaded_bytes(
    audio_bytes: bytes,
    file_extension: str = ".wav",
    model_size: str = None,
    progress_callback: Optional[Callable[[str], None]] = None,
) -> str:
    """
    Streamlit のアップロードバイト列から文字起こしする。

    Args:
        audio_bytes: 音声データのバイト列
        file_extension: ファイル拡張子
        model_size: Whisper モデルサイズ
        progress_callback: 進捗通知用コールバック

    Returns:
        文字起こしテキスト
    """
    with tempfile.NamedTemporaryFile(suffix=file_extension, delete=False) as tmp:
        tmp.write(audio_bytes)
        tmp_path = tmp.name

    try:
        return transcribe_audio(
            tmp_path,
            model_size=model_size,
            progress_callback=progress_callback,
        )
    finally:
        Path(tmp_path).unlink(missing_ok=True)
