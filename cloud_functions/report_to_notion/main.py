"""
Report to Notion - Cloud Function (CF3)

Google Docs の報告書ドラフトを PDF 化し、Notion Database にページを作成する。
オプション機能 — CONFIG.NOTION.ENABLED = true の場合のみ呼び出される。

環境変数:
  - SHARED_SECRET: GAS との共有シークレット
  - NOTION_API_KEY: Notion Integration トークン
  - NOTION_DATABASE_ID: 相談案件データベース ID

デプロイ:
  gcloud functions deploy report_to_notion \
    --gen2 \
    --runtime python311 \
    --trigger-http \
    --allow-unauthenticated \
    --timeout 120 \
    --memory 256MB \
    --set-env-vars SHARED_SECRET=xxx,NOTION_API_KEY=xxx,NOTION_DATABASE_ID=xxx \
    --entry-point report_to_notion
"""

import os
import json
import tempfile
import requests
import functions_framework
from googleapiclient.discovery import build
import google.auth

_creds = None
def _get_creds():
    global _creds
    if _creds is None:
        _creds, _ = google.auth.default()
    return _creds

SHARED_SECRET = os.environ.get("SHARED_SECRET", "")
NOTION_API_KEY = os.environ.get("NOTION_API_KEY", "")
NOTION_DATABASE_ID = os.environ.get("NOTION_DATABASE_ID", "")
NOTION_API_URL = "https://api.notion.com/v1"
NOTION_VERSION = "2022-06-28"


def export_doc_as_pdf(doc_id):
    """Google Docs を PDF としてエクスポート"""
    drive_service = build("drive", "v3", credentials=_get_creds())
    request = drive_service.files().export_media(
        fileId=doc_id, mimeType="application/pdf"
    )
    pdf_content = request.execute()
    print(f"Exported PDF: {len(pdf_content)} bytes")
    return pdf_content


def upload_pdf_to_drive(pdf_content, filename, folder_id=None):
    """PDF を Drive にアップロード"""
    drive_service = build("drive", "v3", credentials=_get_creds())
    from googleapiclient.http import MediaInMemoryUpload

    media = MediaInMemoryUpload(pdf_content, mimetype="application/pdf")
    file_metadata = {"name": filename, "mimeType": "application/pdf"}
    if folder_id:
        file_metadata["parents"] = [folder_id]

    file = drive_service.files().create(
        body=file_metadata, media_body=media, fields="id, webViewLink"
    ).execute()

    file_id = file["id"]
    file_url = file.get("webViewLink", f"https://drive.google.com/file/d/{file_id}")

    # 共有設定：リンクを知っている人が閲覧可能
    drive_service.permissions().create(
        fileId=file_id,
        body={"type": "anyone", "role": "reader"},
    ).execute()

    print(f"PDF uploaded to Drive: {file_id}")
    return file_id, file_url


def create_notion_page(consultation_info, doc_url, pdf_url, report_summary):
    """Notion Database にページを作成"""
    headers = {
        "Authorization": f"Bearer {NOTION_API_KEY}",
        "Content-Type": "application/json",
        "Notion-Version": NOTION_VERSION,
    }

    # Notionページプロパティ
    properties = {
        "Name": {
            "title": [
                {
                    "text": {
                        "content": f"{consultation_info.get('company', '')}様 - {consultation_info.get('application_id', '')}"
                    }
                }
            ]
        },
        "申込ID": {
            "rich_text": [
                {"text": {"content": consultation_info.get("application_id", "")}}
            ]
        },
        "企業名": {
            "rich_text": [
                {"text": {"content": consultation_info.get("company", "")}}
            ]
        },
        "業種": {
            "rich_text": [
                {"text": {"content": consultation_info.get("industry", "")}}
            ]
        },
        "テーマ": {
            "rich_text": [
                {"text": {"content": consultation_info.get("theme", "")}}
            ]
        },
        "リーダー": {
            "rich_text": [
                {"text": {"content": consultation_info.get("leader", "")}}
            ]
        },
        "ステータス": {"select": {"name": "ドラフト完了"}},
    }

    # 相談日をDate型で設定（あれば）
    confirmed_date = consultation_info.get("confirmed_date", "")
    if confirmed_date:
        # yyyy/MM/dd形式をISO形式に変換
        try:
            parts = confirmed_date.replace("/", "-").split(" ")
            date_iso = parts[0]
            properties["相談日"] = {"date": {"start": date_iso}}
        except (ValueError, IndexError):
            pass

    # ページ本文（レポートサマリ + リンク）
    children = [
        {
            "object": "block",
            "type": "heading_2",
            "heading_2": {
                "rich_text": [{"text": {"content": "診断報告書"}}]
            },
        },
        {
            "object": "block",
            "type": "paragraph",
            "paragraph": {
                "rich_text": [
                    {
                        "text": {
                            "content": "Google Docs（編集用）: ",
                        }
                    },
                    {
                        "text": {
                            "content": doc_url,
                            "link": {"url": doc_url},
                        }
                    },
                ]
            },
        },
    ]

    if pdf_url:
        children.append({
            "object": "block",
            "type": "paragraph",
            "paragraph": {
                "rich_text": [
                    {"text": {"content": "PDF（配信用）: "}},
                    {
                        "text": {
                            "content": pdf_url,
                            "link": {"url": pdf_url},
                        }
                    },
                ]
            },
        })

    # レポートサマリを分割して追加（Notion APIの2000文字制限対応）
    if report_summary:
        children.append({
            "object": "block",
            "type": "divider",
            "divider": {},
        })
        children.append({
            "object": "block",
            "type": "heading_3",
            "heading_3": {
                "rich_text": [{"text": {"content": "レポート要約"}}]
            },
        })

        # 2000文字ずつ分割
        chunk_size = 1900
        for i in range(0, len(report_summary), chunk_size):
            chunk = report_summary[i : i + chunk_size]
            children.append({
                "object": "block",
                "type": "paragraph",
                "paragraph": {
                    "rich_text": [{"text": {"content": chunk}}]
                },
            })

    body = {
        "parent": {"database_id": NOTION_DATABASE_ID},
        "properties": properties,
        "children": children[:100],  # Notion API制限: 最大100ブロック
    }

    resp = requests.post(
        f"{NOTION_API_URL}/pages",
        headers=headers,
        json=body,
        timeout=30,
    )

    if resp.status_code != 200:
        print(f"Notion API error: {resp.status_code} {resp.text}")
        resp.raise_for_status()

    page = resp.json()
    page_id = page["id"]
    page_url = page.get("url", "")
    print(f"Notion page created: {page_id}")
    return page_id, page_url


@functions_framework.http
def report_to_notion(request):
    """
    Cloud Function エントリポイント

    期待するリクエストボディ（JSON）:
    {
      "secret": "共有シークレット",
      "doc_id": "Google Docs ID",
      "doc_url": "Google Docs URL",
      "application_id": "申込ID",
      "company": "企業名",
      "industry": "業種",
      "theme": "相談テーマ",
      "leader": "リーダー名",
      "confirmed_date": "相談日時",
      "report_summary": "レポートサマリテキスト（先頭1000文字程度）",
      "drive_folder_id": "PDF保存先フォルダID（オプション）"
    }

    レスポンス:
    {
      "success": true,
      "notion_page_id": "Notion Page ID",
      "notion_page_url": "Notion Page URL",
      "pdf_file_id": "Drive PDF File ID",
      "pdf_file_url": "Drive PDF URL"
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

        doc_id = data.get("doc_id")
        doc_url = data.get("doc_url", "")
        application_id = data.get("application_id", "")
        if not doc_id:
            return json.dumps({
                "success": False,
                "error": "doc_id is required",
            }), 400

        consultation_info = {
            "application_id": application_id,
            "company": data.get("company", ""),
            "industry": data.get("industry", ""),
            "theme": data.get("theme", ""),
            "leader": data.get("leader", ""),
            "confirmed_date": data.get("confirmed_date", ""),
        }
        report_summary = data.get("report_summary", "")
        drive_folder_id = data.get("drive_folder_id", "")

        result = {"success": True, "application_id": application_id}

        # 1. Google Docs → PDF エクスポート
        pdf_content = export_doc_as_pdf(doc_id)
        pdf_filename = f"{application_id}_診断報告書.pdf"
        pdf_file_id, pdf_file_url = upload_pdf_to_drive(
            pdf_content, pdf_filename, drive_folder_id or None
        )
        result["pdf_file_id"] = pdf_file_id
        result["pdf_file_url"] = pdf_file_url

        # 2. Notion ページ作成
        if NOTION_API_KEY and NOTION_DATABASE_ID:
            page_id, page_url = create_notion_page(
                consultation_info, doc_url, pdf_file_url, report_summary
            )
            result["notion_page_id"] = page_id
            result["notion_page_url"] = page_url
        else:
            print("Notion credentials not configured, skipping")
            result["notion_page_id"] = None
            result["notion_page_url"] = None

        return json.dumps(result), 200

    except Exception as e:
        print(f"Error: {e}")
        return json.dumps({
            "success": False,
            "error": str(e),
        }), 500
