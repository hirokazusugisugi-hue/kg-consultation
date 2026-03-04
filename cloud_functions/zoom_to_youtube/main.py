"""
Zoom Recording to YouTube Private Upload - Cloud Function

GAS から HTTP リクエストを受け取り、Zoom 録画を YouTube に非公開アップロードする。

環境変数（Secret Manager）:
  - YOUTUBE_CLIENT_ID: YouTube OAuth クライアントID
  - YOUTUBE_CLIENT_SECRET: YouTube OAuth クライアントシークレット
  - YOUTUBE_REFRESH_TOKEN: YouTube OAuth リフレッシュトークン
  - SHARED_SECRET: GAS との共有シークレット（リクエスト認証用）
"""

import os
import json
import tempfile
import requests
import functions_framework
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2.credentials import Credentials


SHARED_SECRET = os.environ.get("SHARED_SECRET", "")
YOUTUBE_CLIENT_ID = os.environ.get("YOUTUBE_CLIENT_ID", "")
YOUTUBE_CLIENT_SECRET = os.environ.get("YOUTUBE_CLIENT_SECRET", "")
YOUTUBE_REFRESH_TOKEN = os.environ.get("YOUTUBE_REFRESH_TOKEN", "")
TOKEN_URI = "https://oauth2.googleapis.com/token"


def get_youtube_service():
    """YouTube Data API v3 サービスを構築"""
    credentials = Credentials(
        token=None,
        refresh_token=YOUTUBE_REFRESH_TOKEN,
        client_id=YOUTUBE_CLIENT_ID,
        client_secret=YOUTUBE_CLIENT_SECRET,
        token_uri=TOKEN_URI,
    )
    return build("youtube", "v3", credentials=credentials)


def download_zoom_recording(download_url, zoom_token, dest_path):
    """Zoom 録画をダウンロード"""
    headers = {"Authorization": f"Bearer {zoom_token}"}
    resp = requests.get(download_url, headers=headers, stream=True, timeout=300)
    resp.raise_for_status()

    with open(dest_path, "wb") as f:
        for chunk in resp.iter_content(chunk_size=8192):
            f.write(chunk)

    file_size = os.path.getsize(dest_path)
    print(f"Downloaded: {file_size / (1024*1024):.1f} MB -> {dest_path}")
    return file_size


def upload_to_youtube(service, file_path, title, description):
    """YouTube に非公開動画としてアップロード"""
    body = {
        "snippet": {
            "title": title,
            "description": description,
            "categoryId": "27",  # Education
        },
        "status": {
            "privacyStatus": "private",
            "selfDeclaredMadeForKids": False,
        },
    }

    media = MediaFileUpload(
        file_path,
        mimetype="video/mp4",
        resumable=True,
        chunksize=10 * 1024 * 1024,  # 10MB chunks
    )

    request = service.videos().insert(
        part="snippet,status",
        body=body,
        media_body=media,
    )

    response = None
    while response is None:
        status, response = request.next_chunk()
        if status:
            print(f"Upload progress: {int(status.progress() * 100)}%")

    video_id = response["id"]
    print(f"Upload complete: videoId={video_id}")
    return video_id


@functions_framework.http
def zoom_to_youtube(request):
    """
    Cloud Function エントリポイント

    期待するリクエストボディ（JSON）:
    {
      "secret": "共有シークレット",
      "zoom_token": "Zoom アクセストークン",
      "download_url": "Zoom 録画ダウンロード URL",
      "title": "動画タイトル",
      "description": "動画説明文",
      "file_type": "MP4"  (オプション、デフォルト MP4)
    }

    レスポンス:
    {
      "success": true,
      "video_id": "YouTube動画ID",
      "youtube_url": "https://www.youtube.com/watch?v=...",
      "studio_url": "https://studio.youtube.com/video/.../sharing"
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
        title = data.get("title", "経営相談録画")
        description = data.get("description", "")

        if not zoom_token or not download_url:
            return json.dumps({
                "success": False,
                "error": "zoom_token and download_url are required",
            }), 400

        # YouTube 認証情報チェック
        if not YOUTUBE_REFRESH_TOKEN:
            return json.dumps({
                "success": False,
                "error": "YouTube credentials not configured",
            }), 500

        # 一時ファイルにダウンロード
        with tempfile.NamedTemporaryFile(
            suffix=".mp4", delete=False, dir="/tmp"
        ) as tmp:
            tmp_path = tmp.name

        try:
            download_zoom_recording(download_url, zoom_token, tmp_path)

            # YouTube にアップロード
            service = get_youtube_service()
            video_id = upload_to_youtube(service, tmp_path, title, description)

            return json.dumps({
                "success": True,
                "video_id": video_id,
                "youtube_url": f"https://www.youtube.com/watch?v={video_id}",
                "studio_url": f"https://studio.youtube.com/video/{video_id}/sharing",
            }), 200

        finally:
            # 一時ファイル削除
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
                print(f"Cleaned up: {tmp_path}")

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
