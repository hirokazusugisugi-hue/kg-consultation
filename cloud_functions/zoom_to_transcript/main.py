"""
Zoom Recording to Transcript - Cloud Function (CF1)

GAS から HTTP リクエストを受け取り、Zoom 録画を文字起こしする。
Google Cloud Speech-to-Text API v2 を使用し、話者分離（Speaker Diarization）に対応。

環境変数:
  - SHARED_SECRET: GAS との共有シークレット（リクエスト認証用）
  - GCS_BUCKET: 音声ファイルの一時保存先 GCS バケット名
  - SPEECH_MODEL: Speech-to-Text モデル (default: "latest_long")

システム依存:
  - ffmpeg（pydub が音声変換に使用）
    ※ Cloud Functions Gen2 は Dockerfile カスタマイズ or buildpacks で対応
    ※ docker-compose.yml / Dockerfile に `RUN apt-get install -y ffmpeg` を追加

デプロイ:
  gcloud functions deploy zoom_to_transcript \
    --gen2 \
    --runtime python311 \
    --trigger-http \
    --allow-unauthenticated \
    --timeout 540 \
    --memory 1024MB \
    --set-env-vars SHARED_SECRET=xxx,GCS_BUCKET=xxx \
    --entry-point zoom_to_transcript
"""

import os
import json
import tempfile
import requests
import functions_framework
from google.cloud import speech_v2 as speech
from google.cloud import storage


SHARED_SECRET = os.environ.get("SHARED_SECRET", "")
GCS_BUCKET = os.environ.get("GCS_BUCKET", "")
SPEECH_MODEL = os.environ.get("SPEECH_MODEL", "latest_long")
GCP_PROJECT = os.environ.get("GCP_PROJECT", "")


def download_zoom_recording(download_url, zoom_token, dest_path):
    """Zoom 録画をダウンロード"""
    headers = {"Authorization": f"Bearer {zoom_token}"}
    resp = requests.get(download_url, headers=headers, stream=True, timeout=600)
    resp.raise_for_status()

    with open(dest_path, "wb") as f:
        for chunk in resp.iter_content(chunk_size=8192):
            f.write(chunk)

    file_size = os.path.getsize(dest_path)
    print(f"Downloaded: {file_size / (1024*1024):.1f} MB -> {dest_path}")
    return file_size


def extract_audio(video_path, audio_path):
    """MP4 から音声を抽出して FLAC に変換（pydub 使用）"""
    from pydub import AudioSegment

    audio = AudioSegment.from_file(video_path, format="mp4")
    # モノラル、16kHz に変換（Speech-to-Text 推奨）
    audio = audio.set_channels(1).set_frame_rate(16000)
    audio.export(audio_path, format="flac")

    file_size = os.path.getsize(audio_path)
    duration_sec = len(audio) / 1000
    print(f"Audio extracted: {file_size / (1024*1024):.1f} MB, {duration_sec:.0f}s")
    return duration_sec


def upload_to_gcs(bucket_name, source_path, dest_blob_name):
    """GCS にファイルをアップロード"""
    client = storage.Client()
    bucket = client.bucket(bucket_name)
    blob = bucket.blob(dest_blob_name)
    blob.upload_from_filename(source_path)
    gcs_uri = f"gs://{bucket_name}/{dest_blob_name}"
    print(f"Uploaded to GCS: {gcs_uri}")
    return gcs_uri


def delete_from_gcs(bucket_name, blob_name):
    """GCS からファイルを削除"""
    try:
        client = storage.Client()
        bucket = client.bucket(bucket_name)
        blob = bucket.blob(blob_name)
        blob.delete()
        print(f"Deleted from GCS: {blob_name}")
    except Exception as e:
        print(f"GCS cleanup warning: {e}")


def transcribe_audio(gcs_uri, project_id, duration_sec):
    """
    Speech-to-Text API v2 で文字起こし（話者分離対応）
    長時間音声は Long Running Operation で処理
    """
    client = speech.SpeechClient()

    # 話者分離設定（2-6名想定）
    diarization_config = speech.SpeakerDiarizationConfig(
        min_speaker_count=2,
        max_speaker_count=6,
    )

    config = speech.RecognitionConfig(
        auto_decoding_config=speech.AutoDetectDecodingConfig(),
        language_codes=["ja-JP"],
        model=SPEECH_MODEL,
        features=speech.RecognitionFeatures(
            enable_automatic_punctuation=True,
            diarization_config=diarization_config,
            enable_word_time_offsets=True,
        ),
    )

    # GCS URI から認識リクエスト
    file_metadata = speech.BatchRecognizeFileMetadata(uri=gcs_uri)

    request = speech.BatchRecognizeRequest(
        recognizer=f"projects/{project_id}/locations/global/recognizers/_",
        config=config,
        files=[file_metadata],
        recognition_output_config=speech.RecognitionOutputConfig(
            inline_response_config=speech.InlineOutputConfig(),
        ),
    )

    print(f"Starting transcription: {gcs_uri} ({duration_sec:.0f}s)")
    operation = client.batch_recognize(request=request)

    # 長時間処理を待つ（タイムアウト: 最大9分）
    response = operation.result(timeout=520)

    return response


def format_transcript(response, gcs_uri):
    """
    Speech-to-Text レスポンスを読みやすいテキストに整形
    話者ごとにラベル付け
    """
    lines = []
    current_speaker = None

    file_results = response.results.get(gcs_uri)
    if not file_results:
        # キーがフルURIでない場合のフォールバック
        for key, val in response.results.items():
            file_results = val
            break

    if not file_results or not file_results.transcript:
        return "（文字起こし結果が空です）"

    result = file_results.transcript

    for res in result.results:
        if not res.alternatives:
            continue

        alt = res.alternatives[0]
        transcript_text = alt.transcript.strip()
        if not transcript_text:
            continue

        # 話者情報の取得
        speaker_tag = None
        if alt.words:
            speaker_tag = alt.words[0].speaker_label

        if speaker_tag and speaker_tag != current_speaker:
            current_speaker = speaker_tag
            lines.append(f"\n【話者{speaker_tag}】")

        lines.append(transcript_text)

    return "\n".join(lines).strip()


@functions_framework.http
def zoom_to_transcript(request):
    """
    Cloud Function エントリポイント

    期待するリクエストボディ（JSON）:
    {
      "secret": "共有シークレット",
      "zoom_token": "Zoom アクセストークン",
      "download_url": "Zoom 録画ダウンロード URL",
      "application_id": "申込ID",
      "meeting_topic": "ミーティングトピック"
    }

    レスポンス:
    {
      "success": true,
      "application_id": "申込ID",
      "transcript": "文字起こし結果テキスト",
      "duration_sec": 5400,
      "speaker_count": 4
    }
    """
    # CORS preflight
    if request.method == "OPTIONS":
        return ("", 204, {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "POST",
            "Access-Control-Allow-Headers": "Content-Type",
            "Access-Control-Max-Age": "3600",
        })

    if request.method != "POST":
        return json.dumps({"success": False, "error": "POST only"}), 405

    try:
        data = request.get_json(silent=True)
        if not data:
            return json.dumps({"success": False, "error": "Invalid JSON"}), 400

        # 認証チェック
        if not SHARED_SECRET or data.get("secret") != SHARED_SECRET:
            return json.dumps({"success": False, "error": "Unauthorized"}), 401

        # 必須パラメータ
        zoom_token = data.get("zoom_token")
        download_url = data.get("download_url")
        application_id = data.get("application_id", "unknown")
        meeting_topic = data.get("meeting_topic", "")

        if not zoom_token or not download_url:
            return json.dumps({
                "success": False,
                "error": "zoom_token and download_url are required",
            }), 400

        if not GCS_BUCKET:
            return json.dumps({
                "success": False,
                "error": "GCS_BUCKET not configured",
            }), 500

        project_id = GCP_PROJECT
        if not project_id:
            # メタデータサーバーからプロジェクトIDを取得
            try:
                resp = requests.get(
                    "http://metadata.google.internal/computeMetadata/v1/project/project-id",
                    headers={"Metadata-Flavor": "Google"},
                    timeout=5,
                )
                project_id = resp.text
            except Exception:
                return json.dumps({
                    "success": False,
                    "error": "GCP_PROJECT not configured",
                }), 500

        # 一時ファイル
        tmp_dir = tempfile.mkdtemp()
        video_path = os.path.join(tmp_dir, f"{application_id}.mp4")
        audio_path = os.path.join(tmp_dir, f"{application_id}.flac")
        gcs_blob_name = f"transcripts/{application_id}.flac"

        try:
            # 1. Zoom 録画ダウンロード
            download_zoom_recording(download_url, zoom_token, video_path)

            # 2. 音声抽出
            duration_sec = extract_audio(video_path, audio_path)

            # 動画ファイルは不要なので削除
            os.remove(video_path)

            # 3. GCS にアップロード
            gcs_uri = upload_to_gcs(GCS_BUCKET, audio_path, gcs_blob_name)

            # ローカル音声ファイル削除
            os.remove(audio_path)

            # 4. Speech-to-Text で文字起こし
            response = transcribe_audio(gcs_uri, project_id, duration_sec)

            # 5. テキスト整形
            transcript = format_transcript(response, gcs_uri)

            # 話者数カウント
            speaker_set = set()
            for line in transcript.split("\n"):
                if line.startswith("【話者"):
                    speaker_set.add(line)
            speaker_count = len(speaker_set)

            print(
                f"Transcription complete: {application_id}, "
                f"{duration_sec:.0f}s, {speaker_count} speakers, "
                f"{len(transcript)} chars"
            )

            return json.dumps({
                "success": True,
                "application_id": application_id,
                "transcript": transcript,
                "duration_sec": int(duration_sec),
                "speaker_count": speaker_count,
            }), 200

        finally:
            # クリーンアップ
            for f in [video_path, audio_path]:
                if os.path.exists(f):
                    os.remove(f)
            try:
                os.rmdir(tmp_dir)
            except OSError:
                pass

            # GCS クリーンアップ
            delete_from_gcs(GCS_BUCKET, gcs_blob_name)

    except requests.exceptions.HTTPError as e:
        print(f"Zoom download error: {e}")
        return json.dumps({
            "success": False,
            "error": f"Zoom download failed: {e.response.status_code}",
        }), 502

    except Exception as e:
        print(f"Error: {e}")
        return json.dumps({
            "success": False,
            "error": str(e),
        }), 500
