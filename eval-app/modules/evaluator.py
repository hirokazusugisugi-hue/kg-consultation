"""
評価実行エンジン（共通モジュール）

ローカルモード: Claude APIを直接呼び出し
CFモード: Cloud Function経由で呼び出し
"""

import json
import os
from pathlib import Path
from typing import Optional

import anthropic
import requests
from dotenv import load_dotenv

from config.settings import CATEGORIES, RAW_MAX, AI_MAX

load_dotenv()

PROMPTS_DIR = Path(__file__).parent.parent.parent / "cloud_functions" / "consultation_evaluation" / "prompts"
MODEL = "claude-sonnet-4-20250514"
MAX_TRANSCRIPT_CHARS = 120000


def load_prompt(filename: str) -> str:
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


def scale_to_90(raw_total: int) -> int:
    if raw_total <= 0:
        return 0
    return round(raw_total * AI_MAX / RAW_MAX)


def _call_claude(client: anthropic.Anthropic, prompt: str, transcript: str) -> dict:
    user_content = prompt.replace("{transcript}", transcript)
    response = client.messages.create(
        model=MODEL,
        max_tokens=4096,
        temperature=0,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_content}],
    )
    text = response.content[0].text.strip()
    if "```json" in text:
        text = text.split("```json")[1].split("```")[0].strip()
    elif "```" in text:
        text = text.split("```")[1].split("```")[0].strip()
    return json.loads(text)


def evaluate_local(transcript: str, progress_callback=None) -> dict:
    """ローカルモード: Claude APIを直接呼び出して評価"""
    if len(transcript) > MAX_TRANSCRIPT_CHARS:
        transcript = transcript[:MAX_TRANSCRIPT_CHARS] + "\n\n[...テキストが長いため省略されました...]"

    client = anthropic.Anthropic()

    all_scores = {}
    all_evidence = {}
    all_ng = []
    raw_total = 0
    cat_scores = {}

    for i, (cat_key, prompt_file) in enumerate(CALL_PROMPTS):
        prompt = load_prompt(prompt_file)
        if not prompt:
            continue

        cat_name = CATEGORIES[cat_key]["name"]
        if progress_callback:
            progress_callback(i, len(CALL_PROMPTS), f"{cat_name} を評価中...")

        result = _call_claude(client, prompt, transcript)

        subtotal = 0
        for num, data in result.get("items", {}).items():
            score = max(1, min(5, int(data.get("score", 3))))
            all_scores[num] = score
            all_evidence[num] = {
                "evidence": data.get("evidence", ""),
                "reasoning": data.get("reasoning", ""),
            }
            subtotal += score

        raw_total += subtotal
        cat_scores[cat_key] = subtotal
        all_ng.extend(result.get("ng_words", []))

    if progress_callback:
        progress_callback(len(CALL_PROMPTS), len(CALL_PROMPTS), "評価完了")

    ai_total = scale_to_90(raw_total)

    return {
        "ai_total": ai_total,
        "raw_total": raw_total,
        "category_scores": cat_scores,
        "item_scores": all_scores,
        "evidence": all_evidence,
        "ng_words": all_ng,
    }


def evaluate_via_cf(
    transcript: str,
    cf_url: str,
    cf_secret: str,
    metadata: Optional[dict] = None,
) -> dict:
    """CFモード: Cloud Function経由で評価"""
    if not metadata:
        metadata = {}

    payload = {
        "secret": cf_secret,
        "transcript": transcript,
        "evaluation_id": metadata.get("evaluation_id", ""),
        "application_id": metadata.get("application_id", ""),
        "consultant_name": metadata.get("consultant_name", ""),
        "company_name": metadata.get("company_name", ""),
        "industry": metadata.get("industry", ""),
        "theme": metadata.get("theme", ""),
    }

    resp = requests.post(cf_url, json=payload, timeout=660)
    resp.raise_for_status()
    result = resp.json()

    if not result.get("success"):
        raise RuntimeError(result.get("error", "CF returned success=false"))

    return {
        "ai_total": result.get("ai_total", 0),
        "raw_total": result.get("raw_total", 0),
        "category_scores": result.get("categories", {}),
        "item_scores": result.get("item_scores", {}),
        "evidence": result.get("evidence", {}),
        "ng_words": result.get("ng_words", []),
    }
