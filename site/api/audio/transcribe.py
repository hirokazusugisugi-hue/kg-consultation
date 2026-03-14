#!/usr/bin/env python3
"""
音声文字起こしワーカー

Google Cloud Speech-to-Text V2 API (chirp_3, Dynamic Batch) を使用して
音声ファイルを文字起こしし、結果をGAS webhookにPOSTする。

Usage:
    python3 transcribe.py <audio_file_path> <row_number> <gas_webhook_url> <api_token>

前提条件:
    - GCPプロジェクトでSpeech-to-Text API有効化
    - サービスアカウントJSON鍵: ~/.gcp/speech-to-text-key.json
    - GCSバケット作成済み
    - pip install --user google-cloud-speech google-cloud-storage
"""

import sys
import os
import json
import time
import urllib.request
import urllib.error

# GCP設定
GCS_BUCKET = 'kg-consultation-audio'  # GCSバケット名（要変更）
GCP_PROJECT = 'your-gcp-project-id'   # GCPプロジェクトID（要変更）
GCP_LOCATION = 'asia-northeast1'      # リージョン
RECOGNIZER = 'projects/{project}/locations/{location}/recognizers/_'
SERVICE_ACCOUNT_KEY = os.path.expanduser('~/.gcp/speech-to-text-key.json')

# 環境変数からオーバーライド可能
if os.environ.get('GCS_BUCKET'):
    GCS_BUCKET = os.environ['GCS_BUCKET']
if os.environ.get('GCP_PROJECT'):
    GCP_PROJECT = os.environ['GCP_PROJECT']
if os.environ.get('GOOGLE_APPLICATION_CREDENTIALS'):
    SERVICE_ACCOUNT_KEY = os.environ['GOOGLE_APPLICATION_CREDENTIALS']


def upload_to_gcs(local_path, gcs_path):
    """音声ファイルをGCSにアップロード"""
    from google.cloud import storage
    client = storage.Client()
    bucket = client.bucket(GCS_BUCKET)
    blob = bucket.blob(gcs_path)
    blob.upload_from_filename(local_path)
    return f'gs://{GCS_BUCKET}/{gcs_path}'


def delete_from_gcs(gcs_path):
    """GCSの一時ファイルを削除"""
    from google.cloud import storage
    client = storage.Client()
    bucket = client.bucket(GCS_BUCKET)
    blob = bucket.blob(gcs_path)
    if blob.exists():
        blob.delete()


def transcribe_audio(gcs_uri):
    """Speech-to-Text V2 APIで文字起こし（Dynamic Batch）"""
    from google.cloud.speech_v2 import SpeechClient
    from google.cloud.speech_v2.types import cloud_speech

    client = SpeechClient()

    config = cloud_speech.RecognitionConfig(
        auto_decoding_config=cloud_speech.AutoDetectDecodingConfig(),
        language_codes=['ja-JP'],
        model='chirp_2',
        features=cloud_speech.RecognitionFeatures(
            enable_automatic_punctuation=True,
        ),
    )

    file_metadata = cloud_speech.BatchRecognizeFileMetadata(uri=gcs_uri)

    request = cloud_speech.BatchRecognizeRequest(
        recognizer=RECOGNIZER.format(project=GCP_PROJECT, location=GCP_LOCATION),
        config=config,
        files=[file_metadata],
        recognition_output_config=cloud_speech.RecognitionOutputConfig(
            inline_response_config=cloud_speech.InlineOutputConfig(),
        ),
    )

    operation = client.batch_recognize(request=request)
    print(f'文字起こし処理を開始しました。Operation: {operation.operation.name}')

    # 完了まで待機
    response = operation.result(timeout=3600)  # 最大1時間

    # 結果テキストを結合
    transcript_parts = []
    for file_result in response.results.values():
        for result in file_result.transcript.results:
            for alt in result.alternatives:
                transcript_parts.append(alt.transcript)

    return '\n'.join(transcript_parts)


def send_to_gas(webhook_url, row, transcript_text, api_token, status='completed'):
    """結果をGAS webhookにPOST"""
    data = json.dumps({
        'action': 'transcribe-callback',
        'row': str(row),
        'transcript': transcript_text,
        'status': status,
        'token': api_token
    }).encode('utf-8')

    req = urllib.request.Request(
        webhook_url,
        data=data,
        headers={
            'Content-Type': 'application/json; charset=utf-8'
        },
        method='POST'
    )

    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            result = json.loads(resp.read().decode('utf-8'))
            print(f'GAS webhook応答: {result}')
            return result
    except urllib.error.HTTPError as e:
        # GASはリダイレクトすることがあるので、GET で再試行
        if e.code in (301, 302):
            redirect_url = e.headers.get('Location', webhook_url)
            req2 = urllib.request.Request(redirect_url)
            with urllib.request.urlopen(req2, timeout=30) as resp:
                result = json.loads(resp.read().decode('utf-8'))
                print(f'GAS webhook応答（リダイレクト後）: {result}')
                return result
        raise


def main():
    if len(sys.argv) < 5:
        print('Usage: python3 transcribe.py <audio_file_path> <row_number> <gas_webhook_url> <api_token>')
        sys.exit(1)

    audio_path = sys.argv[1]
    row = int(sys.argv[2])
    webhook_url = sys.argv[3]
    api_token = sys.argv[4]

    # サービスアカウント鍵を設定
    if os.path.exists(SERVICE_ACCOUNT_KEY):
        os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = SERVICE_ACCOUNT_KEY

    print(f'音声ファイル: {audio_path}')
    print(f'行番号: {row}')

    if not os.path.exists(audio_path):
        print(f'エラー: ファイルが見つかりません: {audio_path}')
        send_to_gas(webhook_url, row, '', api_token, status='error')
        sys.exit(1)

    gcs_path = f'temp/{row}_{int(time.time())}_{os.path.basename(audio_path)}'

    try:
        # 1. GCSにアップロード
        print('GCSにアップロード中...')
        gcs_uri = upload_to_gcs(audio_path, gcs_path)
        print(f'GCS URI: {gcs_uri}')

        # 2. Speech-to-Text API呼出
        print('文字起こし処理中...')
        transcript = transcribe_audio(gcs_uri)
        print(f'文字起こし完了: {len(transcript)} 文字')

        # 3. GASにコールバック
        print('GAS webhookに送信中...')
        send_to_gas(webhook_url, row, transcript, api_token, status='completed')

        print('全処理が完了しました')

    except Exception as e:
        print(f'エラー: {e}')
        # エラーをGASに通知
        try:
            send_to_gas(webhook_url, row, str(e), api_token, status='error')
        except Exception as e2:
            print(f'GAS通知にも失敗: {e2}')

    finally:
        # 4. GCS一時ファイル削除
        try:
            print('GCS一時ファイルを削除中...')
            delete_from_gcs(gcs_path)
        except Exception:
            pass


if __name__ == '__main__':
    main()
