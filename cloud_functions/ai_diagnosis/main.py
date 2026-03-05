"""
AI Business Diagnosis - Cloud Function (Phase 4)

経営相談の文字起こしテキストまたはテキスト入力から、
6軸スコアリング + 診断レポートを生成する。

6軸評価:
  1. 経営戦略 (Strategy)
  2. 財務 (Finance)
  3. 組織・人材 (Organization)
  4. マーケティング (Marketing)
  5. 業務プロセス (Operations)
  6. IT・DX (Digital)

各軸 1-5点（5段階）、総合 0-100点

環境変数:
  - SHARED_SECRET: GAS との共有シークレット
  - ANTHROPIC_API_KEY: Claude API キー
  - CLAUDE_MODEL: 使用モデル (default: "claude-sonnet-4-20250514")

デプロイ:
  gcloud functions deploy ai_diagnosis \
    --gen2 \
    --runtime python311 \
    --trigger-http \
    --allow-unauthenticated \
    --timeout 120 \
    --memory 512MB \
    --set-env-vars SHARED_SECRET=xxx,ANTHROPIC_API_KEY=xxx \
    --entry-point ai_diagnosis
"""

import os
import json
import functions_framework
import anthropic


SHARED_SECRET = os.environ.get("SHARED_SECRET", "")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
CLAUDE_MODEL = os.environ.get("CLAUDE_MODEL", "claude-sonnet-4-20250514")


DIAGNOSIS_SYSTEM_PROMPT = """あなたは中小企業の経営診断の専門家です。
関西学院大学 中小企業経営診断研究会のAI診断アシスタントとして、
企業の経営状態を6軸で評価してください。

## 評価軸（各1-5点の5段階）

1. **経営戦略 (strategy)**: ビジョン・方向性の明確さ、競争優位性、中長期計画
2. **財務 (finance)**: 収益性、資金繰り、財務管理、投資計画
3. **組織・人材 (organization)**: 組織体制、人材育成、採用・定着、リーダーシップ
4. **マーケティング (marketing)**: 顧客理解、販路、ブランディング、プロモーション
5. **業務プロセス (operations)**: 業務効率化、品質管理、サプライチェーン、生産性
6. **IT・DX (digital)**: IT活用度、デジタル化、データ活用、イノベーション

## 評価基準

- 5点: 非常に優れている（業界トップレベル）
- 4点: 優れている（平均以上、改善の余地あり）
- 3点: 普通（基本的な取り組みはあるが課題あり）
- 2点: やや不足（改善が必要）
- 1点: 大幅な改善が必要（基盤が不十分）

## 出力形式

必ず以下のJSON形式で出力してください。JSON以外のテキストは出力しないでください。

```json
{
  "scores": {
    "strategy": 3,
    "finance": 2,
    "organization": 4,
    "marketing": 3,
    "operations": 2,
    "digital": 1
  },
  "total_score": 62,
  "summary": "企業全体の診断サマリー（100-200文字）",
  "strengths": ["強み1", "強み2", "強み3"],
  "challenges": [
    {
      "area": "課題領域",
      "description": "課題の説明",
      "priority": "高/中/低"
    }
  ],
  "recommendations": [
    {
      "title": "提言タイトル",
      "description": "具体的な提言内容",
      "timeline": "短期（3ヶ月以内）/中期（3-12ヶ月）/長期（1年以上）",
      "expected_impact": "高/中/低"
    }
  ],
  "axis_comments": {
    "strategy": "経営戦略の詳細コメント",
    "finance": "財務の詳細コメント",
    "organization": "組織・人材の詳細コメント",
    "marketing": "マーケティングの詳細コメント",
    "operations": "業務プロセスの詳細コメント",
    "digital": "IT・DXの詳細コメント"
  }
}
```

## 注意事項

- total_score は各軸スコアの加重平均ではなく、企業全体の総合評価（0-100点）
  - 目安: 全軸平均3.0 = 50点、全軸平均5.0 = 100点
  - 計算式の目安: (各軸合計 / 30) × 100 を基に、全体的な印象で調整
- challenges は最低3つ、recommendations は最低3つ出力
- 情報が不足している軸は3点（普通）をデフォルトとし、その旨をaxis_commentsに記載
- 文字起こしの誤変換を考慮して文脈から判断
- 温度感: やや辛口（改善の余地を明確に指摘）
"""


def run_diagnosis(text, company_info):
    """Claude API で経営診断を実行"""
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    user_message = f"""以下の情報から経営診断を行い、JSON形式で結果を出力してください。

## 企業情報
- 企業名: {company_info.get('company', '不明')}
- 業種: {company_info.get('industry', '不明')}
- 相談テーマ: {company_info.get('theme', '')}

## 診断対象テキスト

{text}
"""

    response = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=4096,
        temperature=0.3,  # 一貫性のため低温度
        system=DIAGNOSIS_SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_message}],
    )

    raw_text = response.content[0].text.strip()
    token_usage = {
        "input_tokens": response.usage.input_tokens,
        "output_tokens": response.usage.output_tokens,
    }

    # JSON抽出（コードブロックに囲まれている場合の対応）
    if "```json" in raw_text:
        raw_text = raw_text.split("```json")[1].split("```")[0].strip()
    elif "```" in raw_text:
        raw_text = raw_text.split("```")[1].split("```")[0].strip()

    try:
        diagnosis = json.loads(raw_text)
    except json.JSONDecodeError as e:
        print(f"JSON parse error: {e}\nRaw text: {raw_text[:500]}")
        raise ValueError(f"診断結果のJSON解析に失敗しました: {e}")

    # バリデーション
    scores = diagnosis.get("scores", {})
    required_axes = ["strategy", "finance", "organization", "marketing", "operations", "digital"]
    for axis in required_axes:
        if axis not in scores:
            scores[axis] = 3  # デフォルト
        scores[axis] = max(1, min(5, int(scores[axis])))

    diagnosis["scores"] = scores
    diagnosis["total_score"] = max(0, min(100, int(diagnosis.get("total_score", 50))))

    print(
        f"Diagnosis complete: total={diagnosis['total_score']}, "
        f"scores={scores}, "
        f"tokens: {token_usage['input_tokens']}in/{token_usage['output_tokens']}out"
    )

    return diagnosis, token_usage


@functions_framework.http
def ai_diagnosis(request):
    """
    Cloud Function エントリポイント

    期待するリクエストボディ（JSON）:
    {
      "secret": "共有シークレット",
      "text": "診断対象テキスト（文字起こし or 直接入力）",
      "company": "企業名",
      "industry": "業種",
      "theme": "相談テーマ",
      "application_id": "申込ID（オプション）"
    }

    レスポンス:
    {
      "success": true,
      "application_id": "申込ID",
      "diagnosis": {
        "scores": { "strategy": 3, ... },
        "total_score": 62,
        "summary": "...",
        "strengths": [...],
        "challenges": [...],
        "recommendations": [...],
        "axis_comments": { "strategy": "...", ... }
      },
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

        text = data.get("text", "")
        if not text:
            return json.dumps({
                "success": False,
                "error": "text is required",
            }), 400

        if not ANTHROPIC_API_KEY:
            return json.dumps({
                "success": False,
                "error": "ANTHROPIC_API_KEY not configured",
            }), 500

        company_info = {
            "company": data.get("company", ""),
            "industry": data.get("industry", ""),
            "theme": data.get("theme", ""),
        }

        diagnosis, token_usage = run_diagnosis(text, company_info)

        return json.dumps({
            "success": True,
            "application_id": data.get("application_id", ""),
            "diagnosis": diagnosis,
            "token_usage": token_usage,
        }, ensure_ascii=False), 200

    except ValueError as e:
        return json.dumps({
            "success": False,
            "error": str(e),
        }), 422

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
