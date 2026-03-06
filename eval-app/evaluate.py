"""
コンサルタント評価 ローカル実行スクリプト（CLI）

Usage:
    python evaluate.py transcript.txt [--output result.json]
"""

import argparse
import json
from pathlib import Path

from config.settings import CATEGORIES, ITEM_NAMES
from modules.evaluator import evaluate_local


def main():
    parser = argparse.ArgumentParser(description="Consultant Evaluation (CLI)")
    parser.add_argument("transcript", help="Path to transcript file")
    parser.add_argument("--output", "-o", help="Output JSON path")
    args = parser.parse_args()

    transcript = Path(args.transcript).read_text(encoding="utf-8")
    print(f"Transcript loaded: {len(transcript):,} chars")

    def on_progress(current, total, message):
        print(f"  [{current}/{total}] {message}")

    result = evaluate_local(transcript, progress_callback=on_progress)

    print(f"\n{'='*50}")
    print(f"AI Total Score: {result['ai_total']}/90")
    print(f"Raw Total: {result['raw_total']}/110")
    print(f"{'='*50}")

    for cat_key, cat_info in CATEGORIES.items():
        print(f"  {cat_info['name']}: {result['category_scores'].get(cat_key, 0)}/{cat_info['max_raw']}")

    print(f"\nItem Scores:")
    for num in sorted(result["item_scores"], key=lambda x: int(x)):
        print(f"  No.{num} {ITEM_NAMES.get(int(num), '')}: {result['item_scores'][num]}/5")

    if result["ng_words"]:
        print(f"\nNG Words: {len(result['ng_words'])} detected")

    if args.output:
        Path(args.output).write_text(
            json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(f"\nResults saved to {args.output}")


if __name__ == "__main__":
    main()
