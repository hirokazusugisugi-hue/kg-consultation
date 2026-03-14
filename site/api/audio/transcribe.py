#!/usr/bin/env python3
"""
音声文字起こしワーカー（REST API版）

Google Cloud Speech-to-Text V2 REST API を使用して
音声ファイルを文字起こしし、結果をGAS webhookにPOSTする。
gRPCライブラリ不要（さくら共有ホスティング対応）。

Usage:
    python3 transcribe.py <audio_file_path> <row_number> <gas_webhook_url> <api_token>

前提条件:
    - GCPプロジェクトでSpeech-to-Text API有効化
    - サービスアカウントJSON鍵: ~/.gcp/speech-to-text-key.json
    - GCSバケット作成済み
    - pip install --user google-auth
"""

import sys
import os
import json
import time
import base64
import urllib.request
import urllib.error
import urllib.parse
import mimetypes

# GCP設定
GCS_BUCKET = 'kg-consultation-audio'
GCP_PROJECT = 'kg-consultation-audio'
GCP_LOCATION = 'global'  # batchRecognize は global のみ対応
SERVICE_ACCOUNT_KEY = os.path.expanduser('~/.gcp/speech-to-text-key.json')

# 環境変数からオーバーライド可能
if os.environ.get('GCS_BUCKET'):
    GCS_BUCKET = os.environ['GCS_BUCKET']
if os.environ.get('GCP_PROJECT'):
    GCP_PROJECT = os.environ['GCP_PROJECT']
if os.environ.get('GOOGLE_APPLICATION_CREDENTIALS'):
    SERVICE_ACCOUNT_KEY = os.environ['GOOGLE_APPLICATION_CREDENTIALS']

SCOPES = [
    'https://www.googleapis.com/auth/cloud-platform',
]


def get_access_token():
    """サービスアカウント鍵からアクセストークンを取得"""
    from google.auth.transport.requests import Request
    from google.oauth2 import service_account

    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_KEY, scopes=SCOPES
    )
    credentials.refresh(Request())
    return credentials.token


def upload_to_gcs(local_path, gcs_path, token):
    """音声ファイルをGCSにアップロード（REST API）"""
    # ファイルサイズに応じてアップロード方式を選択
    file_size = os.path.getsize(local_path)
    content_type = mimetypes.guess_type(local_path)[0] or 'application/octet-stream'

    url = 'https://storage.googleapis.com/upload/storage/v1/b/{}/o?uploadType=media&name={}'.format(
        urllib.parse.quote(GCS_BUCKET, safe=''),
        urllib.parse.quote(gcs_path, safe='')
    )

    with open(local_path, 'rb') as f:
        data = f.read()

    req = urllib.request.Request(url, data=data, method='POST')
    req.add_header('Authorization', 'Bearer ' + token)
    req.add_header('Content-Type', content_type)
    req.add_header('Content-Length', str(len(data)))

    with urllib.request.urlopen(req, timeout=600) as resp:
        result = json.loads(resp.read().decode('utf-8'))
        print('GCSアップロード完了: gs://{}/{}'.format(GCS_BUCKET, gcs_path))
        return 'gs://{}/{}'.format(GCS_BUCKET, gcs_path)


def delete_from_gcs(gcs_path, token):
    """GCSの一時ファイルを削除（REST API）"""
    url = 'https://storage.googleapis.com/storage/v1/b/{}/o/{}'.format(
        urllib.parse.quote(GCS_BUCKET, safe=''),
        urllib.parse.quote(gcs_path, safe='')
    )

    req = urllib.request.Request(url, method='DELETE')
    req.add_header('Authorization', 'Bearer ' + token)

    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            pass
    except urllib.error.HTTPError as e:
        if e.code != 404:
            raise


def transcribe_audio(gcs_uri, token):
    """Speech-to-Text V2 REST API で文字起こし（Batch Recognize）"""

    recognizer = 'projects/{}/locations/{}/recognizers/_'.format(GCP_PROJECT, GCP_LOCATION)
    url = 'https://speech.googleapis.com/v2/{}:batchRecognize'.format(
        urllib.parse.quote(recognizer, safe='/:@')
    )

    body = {
        'config': {
            'autoDecodingConfig': {},
            'languageCodes': ['ja-JP'],
            'model': 'long',
            'features': {
                'enableAutomaticPunctuation': True
            }
        },
        'files': [
            {'uri': gcs_uri}
        ],
        'recognitionOutputConfig': {
            'inlineResponseConfig': {}
        }
    }

    data = json.dumps(body).encode('utf-8')
    req = urllib.request.Request(url, data=data, method='POST')
    req.add_header('Authorization', 'Bearer ' + token)
    req.add_header('Content-Type', 'application/json')

    with urllib.request.urlopen(req, timeout=60) as resp:
        result = json.loads(resp.read().decode('utf-8'))

    operation_name = result.get('name', '')
    print('文字起こし処理開始: {}'.format(operation_name))

    # オペレーション完了をポーリング
    poll_url = 'https://speech.googleapis.com/v2/{}'.format(
        urllib.parse.quote(operation_name, safe='/:@')
    )

    for attempt in range(360):  # 最大1時間（10秒 x 360）
        time.sleep(10)

        poll_req = urllib.request.Request(poll_url, method='GET')
        poll_req.add_header('Authorization', 'Bearer ' + token)

        with urllib.request.urlopen(poll_req, timeout=30) as resp:
            op_result = json.loads(resp.read().decode('utf-8'))

        if op_result.get('done'):
            if 'error' in op_result:
                raise Exception('Speech API error: {}'.format(op_result['error']))

            response = op_result.get('response', {})
            results_map = response.get('results', {})

            transcript_parts = []
            for file_key, file_result in results_map.items():
                transcript_obj = file_result.get('transcript', {})
                for result in transcript_obj.get('results', []):
                    for alt in result.get('alternatives', []):
                        text = alt.get('transcript', '')
                        if text:
                            transcript_parts.append(text)

            return '\n'.join(transcript_parts)

        if attempt % 6 == 0:
            print('文字起こし処理中... ({}秒経過)'.format(attempt * 10))

    raise Exception('文字起こし処理がタイムアウトしました（1時間）')


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
            print('GAS webhook応答: {}'.format(result))
            return result
    except urllib.error.HTTPError as e:
        if e.code in (301, 302):
            redirect_url = e.headers.get('Location', webhook_url)
            req2 = urllib.request.Request(redirect_url)
            with urllib.request.urlopen(req2, timeout=30) as resp:
                result = json.loads(resp.read().decode('utf-8'))
                print('GAS webhook応答（リダイレクト後）: {}'.format(result))
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
    else:
        print('エラー: サービスアカウント鍵が見つかりません: {}'.format(SERVICE_ACCOUNT_KEY))
        send_to_gas(webhook_url, row, 'サービスアカウント鍵が未配置です', api_token, status='error')
        sys.exit(1)

    print('音声ファイル: {}'.format(audio_path))
    print('行番号: {}'.format(row))

    if not os.path.exists(audio_path):
        print('エラー: ファイルが見つかりません: {}'.format(audio_path))
        send_to_gas(webhook_url, row, '', api_token, status='error')
        sys.exit(1)

    gcs_path = 'temp/{}_{}'.format(row, os.path.basename(audio_path))
    access_token = None

    try:
        # 0. アクセストークン取得
        print('GCP認証中...')
        access_token = get_access_token()

        # 1. GCSにアップロード
        print('GCSにアップロード中...')
        gcs_uri = upload_to_gcs(audio_path, gcs_path, access_token)

        # 2. Speech-to-Text API呼出
        print('文字起こし処理中...')
        transcript = transcribe_audio(gcs_uri, access_token)
        print('文字起こし完了: {} 文字'.format(len(transcript)))

        # 3. GASにコールバック
        print('GAS webhookに送信中...')
        send_to_gas(webhook_url, row, transcript, api_token, status='completed')

        print('全処理が完了しました')

    except Exception as e:
        print('エラー: {}'.format(e))
        try:
            send_to_gas(webhook_url, row, str(e), api_token, status='error')
        except Exception as e2:
            print('GAS通知にも失敗: {}'.format(e2))

    finally:
        # 4. GCS一時ファイル削除
        if access_token:
            try:
                print('GCS一時ファイルを削除中...')
                delete_from_gcs(gcs_path, access_token)
            except Exception:
                pass


if __name__ == '__main__':
    main()
