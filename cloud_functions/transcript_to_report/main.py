"""
Transcript to Report - Cloud Function (CF2)

文字起こしテキストから Claude API で構造化レポートを生成し、
Google Docs に保存する。

環境変数:
  - SHARED_SECRET: GAS との共有シークレット
  - ANTHROPIC_API_KEY: Claude API キー
  - GOOGLE_DOCS_FOLDER_ID: レポート保存先 Drive フォルダ ID
  - CLAUDE_MODEL: 使用モデル (default: "claude-sonnet-4-20250514")

デプロイ:
  gcloud functions deploy transcript_to_report \
    --gen2 \
    --runtime python311 \
    --trigger-http \
    --allow-unauthenticated \
    --timeout 300 \
    --memory 512MB \
    --set-env-vars SHARED_SECRET=xxx,ANTHROPIC_API_KEY=xxx,GOOGLE_DOCS_FOLDER_ID=xxx \
    --entry-point transcript_to_report
"""

import os
import json
import functions_framework
import anthropic
from googleapiclient.discovery import build
import google.auth


SHARED_SECRET = os.environ.get("SHARED_SECRET", "")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
CLAUDE_MODEL = os.environ.get("CLAUDE_MODEL", "claude-sonnet-4-20250514")
GOOGLE_DOCS_FOLDER_ID = os.environ.get("GOOGLE_DOCS_FOLDER_ID", "")

# レポート生成プロンプト
REPORT_SYSTEM_PROMPT = """あなたは中小企業経営の専門家であり、関西学院大学 中小企業経営診断研究会の診断報告書を作成するアシスタントです。

経営相談の文字起こしテキストから、以下の構造で診断報告書のドラフトを作成してください。

## 報告書の構成

### 1. 相談概要
- 相談日時、相談企業、業種、相談テーマを簡潔にまとめる
- 参加者（診断士・オブザーバー）を記載

### 2. 現状分析
- 相談内容から読み取れる企業の現状を整理
- 強み・機会・課題を箇条書きで記載

### 3. 課題整理
- 主要な経営課題を優先順位付きで列挙
- 各課題の背景・影響を簡潔に説明

### 4. 提言・アドバイス
- 相談で出たアドバイスを整理・構造化
- 短期（3ヶ月以内）・中期（3-12ヶ月）に分類

### 5. アクションプラン
- 具体的な次のステップを3-5項目で提案
- 各項目に期限目安と担当を示唆

### 6. 参考情報
- 相談中に言及された参考リソース、制度、ツール等を整理

## 注意事項
- 文字起こしの誤変換を適宜修正して記載
- 話者分離がある場合、診断士側の発言とクライアント側の発言を区別して分析
- 専門用語は平易な表現を併記
- 機密情報（個人名以外の具体的数値等）はそのまま記載可
- Markdown形式で出力（Google Docsに変換される前提）
"""


def generate_report_with_claude(transcript, consultation_info):
    """Claude API でレポートドラフトを生成"""
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    user_message = f"""以下の経営相談の文字起こしから、診断報告書ドラフトを作成してください。

## 相談情報
- 申込ID: {consultation_info.get('application_id', '')}
- 相談日時: {consultation_info.get('confirmed_date', '')}
- 相談企業: {consultation_info.get('company', '')}
- 業種: {consultation_info.get('industry', '')}
- 相談テーマ: {consultation_info.get('theme', '')}
- 相談者: {consultation_info.get('name', '')}
- リーダー: {consultation_info.get('leader', '')}

## 文字起こし

{transcript}
"""

    response = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=8192,
        system=REPORT_SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_message}],
    )

    report_text = response.content[0].text
    token_usage = {
        "input_tokens": response.usage.input_tokens,
        "output_tokens": response.usage.output_tokens,
    }

    print(
        f"Claude report generated: {len(report_text)} chars, "
        f"tokens: {token_usage['input_tokens']}in/{token_usage['output_tokens']}out"
    )
    return report_text, token_usage


def create_google_doc(title, content, folder_id=None):
    """Google Docs にレポートを作成"""
    # Cloud Functions の Application Default Credentials を使用
    creds, _ = google.auth.default()
    docs_service = build("docs", "v1", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)

    # 新規ドキュメント作成
    doc = docs_service.documents().create(body={"title": title}).execute()
    doc_id = doc["documentId"]
    print(f"Created Google Doc: {doc_id}")

    # Markdown をドキュメントに挿入
    # シンプルなテキスト挿入（Markdown のままリーダーが編集する前提）
    requests_body = []

    # 本文を挿入
    requests_body.append({
        "insertText": {
            "location": {"index": 1},
            "text": content,
        }
    })

    if requests_body:
        docs_service.documents().batchUpdate(
            documentId=doc_id,
            body={"requests": requests_body},
        ).execute()

    # フォルダに移動
    if folder_id:
        try:
            # 現在の親フォルダを取得
            file = drive_service.files().get(
                fileId=doc_id, fields="parents"
            ).execute()
            previous_parents = ",".join(file.get("parents", []))

            drive_service.files().update(
                fileId=doc_id,
                addParents=folder_id,
                removeParents=previous_parents,
                fields="id, parents",
            ).execute()
            print(f"Moved to folder: {folder_id}")
        except Exception as e:
            print(f"Folder move warning: {e}")

    doc_url = f"https://docs.google.com/document/d/{doc_id}/edit"
    return doc_id, doc_url


@functions_framework.http
def transcript_to_report(request):
    """
    Cloud Function エントリポイント

    期待するリクエストボディ（JSON）:
    {
      "secret": "共有シークレット",
      "transcript": "文字起こしテキスト",
      "application_id": "申込ID",
      "company": "企業名",
      "industry": "業種",
      "theme": "相談テーマ",
      "name": "相談者名",
      "leader": "リーダー名",
      "confirmed_date": "相談日時"
    }

    レスポンス:
    {
      "success": true,
      "application_id": "申込ID",
      "doc_id": "Google Docs ID",
      "doc_url": "Google Docs URL",
      "report_text": "レポートテキスト（Markdown）",
      "token_usage": { "input_tokens": 1234, "output_tokens": 5678 }
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

        transcript = data.get("transcript", "")
        if not transcript:
            return json.dumps({
                "success": False,
                "error": "transcript is required",
            }), 400

        if not ANTHROPIC_API_KEY:
            return json.dumps({
                "success": False,
                "error": "ANTHROPIC_API_KEY not configured",
            }), 500

        application_id = data.get("application_id", "unknown")
        consultation_info = {
            "application_id": application_id,
            "company": data.get("company", ""),
            "industry": data.get("industry", ""),
            "theme": data.get("theme", ""),
            "name": data.get("name", ""),
            "leader": data.get("leader", ""),
            "confirmed_date": data.get("confirmed_date", ""),
        }

        # 1. Claude API でレポート生成
        report_text, token_usage = generate_report_with_claude(
            transcript, consultation_info
        )

        # 2. Google Docs に保存
        doc_title = (
            f"【診断報告書ドラフト】{consultation_info['company']}様 - "
            f"{application_id}"
        )
        folder_id = GOOGLE_DOCS_FOLDER_ID or None
        doc_id, doc_url = create_google_doc(doc_title, report_text, folder_id)

        print(
            f"Report created: {application_id} -> {doc_id}"
        )

        return json.dumps({
            "success": True,
            "application_id": application_id,
            "doc_id": doc_id,
            "doc_url": doc_url,
            "report_text": report_text,
            "token_usage": token_usage,
        }), 200

    except anthropic.APIError as e:
        print(f"Claude API error: {e}")
        return json.dumps({
            "success": False,
            "error": f"Claude API error: {str(e)}",
        }), 502

    except Exception as e:
        print(f"Error: {e}")
        return json.dumps({
            "success": False,
            "error": str(e),
        }), 500
