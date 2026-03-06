"""
コンサルタント評価 Cloud Function

ICMCI CMC・Schein理論・SERVQUAL・MITIに基づく22項目の学術的評価。
Claude API × 6回分割呼出による精密評価を実行。

Deploy:
    gcloud functions deploy consultation_evaluation \
        --gen2 --runtime python311 --trigger-http \
        --timeout 600 --memory 1024MB \
        --set-env-vars SHARED_SECRET=xxx,ANTHROPIC_API_KEY=xxx
"""

import json
import os
import traceback
from pathlib import Path

import anthropic
import functions_framework

# ── Config ──
SHARED_SECRET = os.environ.get("SHARED_SECRET", "")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
MODEL = "claude-sonnet-4-20250514"
MAX_TRANSCRIPT_CHARS = 120000  # Claude context limit safety margin

# ── Prompt loading ──
PROMPTS_DIR = Path(__file__).parent / "prompts"

def load_prompt(filename: str) -> str:
    """Load a prompt template from the prompts directory."""
    path = PROMPTS_DIR / filename
    if path.exists():
        return path.read_text(encoding="utf-8")
    return ""

SYSTEM_PROMPT = load_prompt("system.txt")

CALL_PROMPTS = [
    ("c1", "call1_problem.txt"),
    ("c2", "call2_solution.txt"),
    ("c3", "call3_communication.txt"),
    ("c4", "call4_time.txt"),
    ("c5", "call5_logic.txt"),
    ("c6", "call6_ethics.txt"),
]

# ── Category metadata ──
CATEGORY_MAX_SCORES = {
    "c1": 20,  # 4 items x 5 points
    "c2": 25,  # 5 items x 5 points
    "c3": 20,  # 4 items x 5 points
    "c4": 15,  # 3 items x 5 points
    "c5": 15,  # 3 items x 5 points
    "c6": 15,  # 3 items x 5 points
}  # Total max: 110 points raw, scaled to 90

TOTAL_RAW_MAX = sum(CATEGORY_MAX_SCORES.values())  # 110
AI_SCALED_MAX = 90


def scale_to_90(raw_total: int) -> int:
    """Scale raw score (0-110) to AI score (0-90)."""
    if raw_total <= 0:
        return 0
    return round(raw_total * AI_SCALED_MAX / TOTAL_RAW_MAX)


def scale_category(raw_score: int, category: str) -> float:
    """Scale individual category score proportionally to 90-point total."""
    max_raw = CATEGORY_MAX_SCORES.get(category, 15)
    if max_raw == 0:
        return 0
    return round(raw_score * (max_raw / TOTAL_RAW_MAX) * AI_SCALED_MAX, 1)


# ── Claude API call ──

def call_claude(client: anthropic.Anthropic, prompt: str, transcript: str) -> dict:
    """Call Claude API with a specific evaluation prompt."""
    user_content = prompt.replace("{transcript}", transcript)

    response = client.messages.create(
        model=MODEL,
        max_tokens=4096,
        temperature=0,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_content}],
    )

    text = response.content[0].text.strip()

    # Extract JSON from response (handle markdown code blocks)
    if "```json" in text:
        text = text.split("```json")[1].split("```")[0].strip()
    elif "```" in text:
        text = text.split("```")[1].split("```")[0].strip()

    return json.loads(text)


# ── Main evaluation pipeline ──

def evaluate_transcript(transcript: str, metadata: dict) -> dict:
    """Run the full 6-call evaluation pipeline."""
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    # Truncate transcript if too long
    if len(transcript) > MAX_TRANSCRIPT_CHARS:
        transcript = transcript[:MAX_TRANSCRIPT_CHARS] + "\n\n[...テキストが長いため省略されました...]"

    results = {}
    all_item_scores = {}
    all_evidence = {}
    all_ng_words = []
    raw_total = 0
    category_scores = {}

    for cat_key, prompt_file in CALL_PROMPTS:
        prompt_template = load_prompt(prompt_file)
        if not prompt_template:
            print(f"Warning: prompt file {prompt_file} not found, skipping")
            continue

        try:
            result = call_claude(client, prompt_template, transcript)
            results[cat_key] = result

            # Collect item scores and evidence
            items = result.get("items", {})
            subtotal = 0
            for item_num, item_data in items.items():
                score = int(item_data.get("score", 3))
                score = max(1, min(5, score))  # Clamp to 1-5
                all_item_scores[item_num] = score
                all_evidence[item_num] = {
                    "evidence": item_data.get("evidence", ""),
                    "reasoning": item_data.get("reasoning", ""),
                }
                subtotal += score

            raw_total += subtotal
            category_scores[cat_key] = subtotal

            # Collect NG words
            ng = result.get("ng_words", [])
            if isinstance(ng, list):
                all_ng_words.extend(ng)

        except Exception as e:
            print(f"Error in {cat_key}: {e}")
            traceback.print_exc()
            # Default scores for failed category
            category_scores[cat_key] = 0

    # Scale scores
    ai_total = scale_to_90(raw_total)
    scaled_categories = {}
    for cat_key in ["c1", "c2", "c3", "c4", "c5", "c6"]:
        scaled_categories[cat_key] = scale_category(
            category_scores.get(cat_key, 0), cat_key
        )

    return {
        "success": True,
        "evaluation_id": metadata.get("evaluation_id", ""),
        "ai_total": ai_total,
        "raw_total": raw_total,
        "categories": scaled_categories,
        "item_scores": all_item_scores,
        "evidence": all_evidence,
        "ng_words": all_ng_words,
    }


# ── HTTP entry point ──

@functions_framework.http
def consultation_evaluation(request):
    """HTTP Cloud Function entry point."""
    # CORS
    if request.method == "OPTIONS":
        headers = {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "POST, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type",
            "Access-Control-Max-Age": "3600",
        }
        return ("", 204, headers)

    headers = {"Access-Control-Allow-Origin": "*", "Content-Type": "application/json"}

    try:
        data = request.get_json(silent=True) or {}

        # Auth
        secret = data.get("secret", "")
        if SHARED_SECRET and secret != SHARED_SECRET:
            return (json.dumps({"success": False, "error": "Unauthorized"}), 403, headers)

        # Validate
        transcript = data.get("transcript", "")
        if not transcript:
            return (
                json.dumps({"success": False, "error": "transcript is required"}),
                400,
                headers,
            )

        if not ANTHROPIC_API_KEY:
            return (
                json.dumps({"success": False, "error": "ANTHROPIC_API_KEY not configured"}),
                500,
                headers,
            )

        metadata = {
            "evaluation_id": data.get("evaluation_id", ""),
            "application_id": data.get("application_id", ""),
            "consultant_name": data.get("consultant_name", ""),
            "company_name": data.get("company_name", ""),
            "industry": data.get("industry", ""),
            "theme": data.get("theme", ""),
        }

        print(f"Starting evaluation: {metadata.get('evaluation_id', 'unknown')}")

        result = evaluate_transcript(transcript, metadata)

        print(
            f"Evaluation complete: {metadata.get('evaluation_id', 'unknown')}, "
            f"AI total: {result.get('ai_total', 0)}/90"
        )

        return (json.dumps(result, ensure_ascii=False), 200, headers)

    except Exception as e:
        print(f"Error: {e}")
        traceback.print_exc()
        return (
            json.dumps({"success": False, "error": str(e)}),
            500,
            headers,
        )
