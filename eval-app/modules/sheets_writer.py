"""
Google Sheets自動保存モジュール

評価結果を研究用スプレッドシート（4シート構成）に自動保存する。
gspread + google-authを使用したサービスアカウント認証。

環境変数:
    GOOGLE_SHEETS_CREDENTIALS_PATH: サービスアカウントJSONキーのパス
    EVAL_SPREADSHEET_ID: 既存スプレッドシートID（未設定時は自動作成）
    EVAL_SPREADSHEET_SHARE_EMAIL: 管理者メール（自動共有用、任意）
"""

import json
import logging
import os
from datetime import datetime
from pathlib import Path

logger = logging.getLogger(__name__)

# シート名定義
SHEET_EVAL_DATA = "評価データ"
SHEET_EVIDENCE = "エビデンス"
SHEET_NG_WORDS = "NG語句"
SHEET_TRANSCRIPT = "文字起こし"

# 評価データシートのヘッダー（40列）
EVAL_DATA_HEADERS = [
    "eval_id", "evaluated_at",
    # メタデータ (C-J)
    "consultant_name", "company_name", "industry", "theme",
    "consultation_type", "input_type", "transcript_chars", "transcription_time_sec",
    # スコアサマリー (K-L)
    "ai_total", "raw_total",
    # カテゴリ小計 (M-R)
    "c1_問題の本質把握", "c2_解決策・方針提示", "c3_コミュニケーション品質",
    "c4_時間管理・進行力", "c5_論理的展開力", "c6_倫理・自律性支援",
    # 22項目個別スコア (S-AN)
    "item_01_傾聴・受容", "item_02_質問力・深掘り", "item_03_本質課題の特定",
    "item_04_情報収集の網羅性", "item_05_解決策の具体性", "item_06_複数選択肢の提示",
    "item_07_実現可能性への配慮", "item_08_専門知識の適切な活用", "item_09_リスク・留意点の説明",
    "item_10_わかりやすさ", "item_11_信頼関係の構築", "item_12_要約・確認",
    "item_13_適切な言葉遣い", "item_14_時間配分の適切さ", "item_15_議論の進行管理",
    "item_16_ネクストステップの明確化", "item_17_論理構成の明確さ", "item_18_根拠に基づく説明",
    "item_19_構造化・可視化", "item_20_クライアントの自律性尊重", "item_21_守秘義務・倫理意識",
    "item_22_継続的改善の促進",
    # その他 (AO-AP)
    "ng_word_count", "json_filename",
]

EVIDENCE_HEADERS = [
    "eval_id", "item_number", "item_name", "category", "score", "evidence", "reasoning",
]

NG_WORDS_HEADERS = [
    "eval_id", "evaluated_at", "consultant_name", "ng_category", "ng_text", "ng_context",
]

TRANSCRIPT_HEADERS = [
    "eval_id", "evaluated_at", "consultant_name", "transcript",
]

# 文字起こしシートの文字数上限
TRANSCRIPT_CHAR_LIMIT = 50000

# ITEM_NAMESのローカルコピー（循環importを避ける）
_ITEM_NAMES = {
    1: "傾聴・受容", 2: "質問力・深掘り", 3: "本質課題の特定", 4: "情報収集の網羅性",
    5: "解決策の具体性", 6: "複数選択肢の提示", 7: "実現可能性への配慮",
    8: "専門知識の適切な活用", 9: "リスク・留意点の説明", 10: "わかりやすさ",
    11: "信頼関係の構築", 12: "要約・確認", 13: "適切な言葉遣い",
    14: "時間配分の適切さ", 15: "議論の進行管理", 16: "ネクストステップの明確化",
    17: "論理構成の明確さ", 18: "根拠に基づく説明", 19: "構造化・可視化",
    20: "クライアントの自律性尊重", 21: "守秘義務・倫理意識", 22: "継続的改善の促進",
}

# カテゴリとアイテムの対応
_ITEM_TO_CATEGORY = {}
_CATEGORY_NAMES = {
    "c1": "問題の本質把握", "c2": "解決策・方針提示", "c3": "コミュニケーション品質",
    "c4": "時間管理・進行力", "c5": "論理的展開力", "c6": "倫理・自律性支援",
}
_CATEGORY_ITEMS = {
    "c1": [1, 2, 3, 4], "c2": [5, 6, 7, 8, 9], "c3": [10, 11, 12, 13],
    "c4": [14, 15, 16], "c5": [17, 18, 19], "c6": [20, 21, 22],
}
for _ck, _items in _CATEGORY_ITEMS.items():
    for _item in _items:
        _ITEM_TO_CATEGORY[_item] = _ck


class SheetsWriter:
    """Google Sheets自動保存クラス"""

    def __init__(self, credentials_path: str, spreadsheet_id: str = None,
                 share_email: str = None):
        """
        Args:
            credentials_path: サービスアカウントJSONキーのパス
            spreadsheet_id: 既存スプレッドシートID（Noneなら自動作成）
            share_email: 自動共有先メールアドレス（任意）
        """
        import gspread
        from google.oauth2.service_account import Credentials

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = Credentials.from_service_account_file(credentials_path, scopes=scopes)
        self.gc = gspread.authorize(creds)
        self.share_email = share_email
        self.spreadsheet = None
        self.spreadsheet_id = spreadsheet_id

        if spreadsheet_id:
            self.spreadsheet = self.gc.open_by_key(spreadsheet_id)
        else:
            self._create_spreadsheet()

    def _create_spreadsheet(self):
        """新規スプレッドシートを作成し、4シート構成にする"""
        title = f"コンサルタント評価データ_{datetime.now().strftime('%Y%m%d')}"
        self.spreadsheet = self.gc.create(title)
        self.spreadsheet_id = self.spreadsheet.id
        logger.info(f"スプレッドシート作成: {self.spreadsheet_id}")
        print(f"\n{'='*60}")
        print(f"新規スプレッドシートを作成しました")
        print(f"ID: {self.spreadsheet_id}")
        print(f"URL: https://docs.google.com/spreadsheets/d/{self.spreadsheet_id}")
        print(f".envに以下を設定してください:")
        print(f"EVAL_SPREADSHEET_ID={self.spreadsheet_id}")
        print(f"{'='*60}\n")

        # Sheet1をリネームして「評価データ」に
        sheet1 = self.spreadsheet.sheet1
        sheet1.update_title(SHEET_EVAL_DATA)
        sheet1.append_row(EVAL_DATA_HEADERS)

        # 残り3シート作成
        ws_evidence = self.spreadsheet.add_worksheet(SHEET_EVIDENCE, rows=1000, cols=7)
        ws_evidence.append_row(EVIDENCE_HEADERS)

        ws_ng = self.spreadsheet.add_worksheet(SHEET_NG_WORDS, rows=1000, cols=6)
        ws_ng.append_row(NG_WORDS_HEADERS)

        ws_transcript = self.spreadsheet.add_worksheet(SHEET_TRANSCRIPT, rows=1000, cols=4)
        ws_transcript.append_row(TRANSCRIPT_HEADERS)

        # ヘッダー行を太字にする（各シート）
        for ws in [sheet1, ws_evidence, ws_ng, ws_transcript]:
            ws.format("1", {"textFormat": {"bold": True}})

        # 管理者メールに共有
        if self.share_email:
            self.spreadsheet.share(self.share_email, perm_type="user", role="writer")
            logger.info(f"スプレッドシートを共有: {self.share_email}")

    def _ensure_sheet(self, sheet_name: str, headers: list):
        """シートが存在しなければ作成する"""
        try:
            return self.spreadsheet.worksheet(sheet_name)
        except Exception:
            ws = self.spreadsheet.add_worksheet(sheet_name, rows=1000, cols=len(headers))
            ws.append_row(headers)
            ws.format("1", {"textFormat": {"bold": True}})
            return ws

    def save_result(self, result: dict, transcript: str, json_filename: str):
        """評価結果を4シートに一括保存する

        Args:
            result: 評価結果dict（ai_total, raw_total, category_scores, item_scores, evidence, ng_words, metadata）
            transcript: 文字起こしテキスト
            json_filename: ローカルJSONファイル名
        """
        meta = result.get("metadata", {})
        eval_id = meta.get("evaluated_at", datetime.now().isoformat())
        evaluated_at = meta.get("evaluated_at", "")
        consultant_name = meta.get("consultant_name", "")

        # Sheet 1: 評価データ
        self._save_eval_data(result, meta, eval_id, json_filename)

        # Sheet 2: エビデンス
        self._save_evidence(result, eval_id)

        # Sheet 3: NG語句
        self._save_ng_words(result, eval_id, evaluated_at, consultant_name)

        # Sheet 4: 文字起こし
        self._save_transcript(eval_id, evaluated_at, consultant_name, transcript)

        logger.info(f"Google Sheets保存完了: eval_id={eval_id}")

    def _save_eval_data(self, result: dict, meta: dict, eval_id: str, json_filename: str):
        """Sheet 1: 評価データ（1行）"""
        ws = self._ensure_sheet(SHEET_EVAL_DATA, EVAL_DATA_HEADERS)

        item_scores = result.get("item_scores", {})
        category_scores = result.get("category_scores", {})

        row = [
            eval_id,
            meta.get("evaluated_at", ""),
            # メタデータ
            meta.get("consultant_name", ""),
            meta.get("company_name", ""),
            meta.get("industry", ""),
            meta.get("theme", ""),
            meta.get("consultation_type", ""),
            meta.get("input_type", ""),
            meta.get("transcript_chars", ""),
            meta.get("transcription_time_sec", ""),
            # スコアサマリー
            result.get("ai_total", 0),
            result.get("raw_total", 0),
            # カテゴリ小計
            category_scores.get("c1", 0),
            category_scores.get("c2", 0),
            category_scores.get("c3", 0),
            category_scores.get("c4", 0),
            category_scores.get("c5", 0),
            category_scores.get("c6", 0),
        ]

        # 22項目個別スコア
        for i in range(1, 23):
            score = item_scores.get(str(i), item_scores.get(i, 0))
            row.append(score)

        # NG語句数 + ファイル名
        row.append(len(result.get("ng_words", [])))
        row.append(json_filename)

        ws.append_row(row, value_input_option="USER_ENTERED")

    def _save_evidence(self, result: dict, eval_id: str):
        """Sheet 2: エビデンス（22行）"""
        ws = self._ensure_sheet(SHEET_EVIDENCE, EVIDENCE_HEADERS)

        item_scores = result.get("item_scores", {})
        evidence = result.get("evidence", {})

        rows = []
        for i in range(1, 23):
            num_str = str(i)
            score = item_scores.get(num_str, item_scores.get(i, 0))
            ev = evidence.get(num_str, evidence.get(str(i), {}))
            if not isinstance(ev, dict):
                ev = {}

            cat_key = _ITEM_TO_CATEGORY.get(i, "")
            cat_name = _CATEGORY_NAMES.get(cat_key, "")

            rows.append([
                eval_id,
                i,
                _ITEM_NAMES.get(i, ""),
                cat_name,
                score,
                ev.get("evidence", ""),
                ev.get("reasoning", ""),
            ])

        ws.append_rows(rows, value_input_option="USER_ENTERED")

    def _save_ng_words(self, result: dict, eval_id: str, evaluated_at: str,
                       consultant_name: str):
        """Sheet 3: NG語句（可変行数）"""
        ng_words = result.get("ng_words", [])
        if not ng_words:
            return

        ws = self._ensure_sheet(SHEET_NG_WORDS, NG_WORDS_HEADERS)

        rows = []
        for ng in ng_words:
            rows.append([
                eval_id,
                evaluated_at,
                consultant_name,
                ng.get("category", ""),
                ng.get("text", ""),
                ng.get("context", ""),
            ])

        ws.append_rows(rows, value_input_option="USER_ENTERED")

    def _save_transcript(self, eval_id: str, evaluated_at: str, consultant_name: str,
                         transcript: str):
        """Sheet 4: 文字起こし（1行、50,000文字上限）"""
        ws = self._ensure_sheet(SHEET_TRANSCRIPT, TRANSCRIPT_HEADERS)

        # 上限超過時は切り詰め
        if len(transcript) > TRANSCRIPT_CHAR_LIMIT:
            truncated = transcript[:TRANSCRIPT_CHAR_LIMIT]
            truncated += f"\n\n[...以降省略（全{len(transcript):,}文字中{TRANSCRIPT_CHAR_LIMIT:,}文字まで記録。全文はローカルJSONを参照）]"
        else:
            truncated = transcript

        ws.append_row(
            [eval_id, evaluated_at, consultant_name, truncated],
            value_input_option="USER_ENTERED",
        )

    def check_connection(self) -> dict:
        """接続状態チェック

        Returns:
            dict: {"connected": bool, "spreadsheet_title": str, "spreadsheet_url": str, "sheet_count": int}
        """
        try:
            title = self.spreadsheet.title
            url = f"https://docs.google.com/spreadsheets/d/{self.spreadsheet_id}"
            sheets = self.spreadsheet.worksheets()
            return {
                "connected": True,
                "spreadsheet_title": title,
                "spreadsheet_url": url,
                "sheet_count": len(sheets),
                "sheet_names": [s.title for s in sheets],
            }
        except Exception as e:
            return {
                "connected": False,
                "error": str(e),
            }


def get_sheets_writer() -> "SheetsWriter | None":
    """環境変数からSheetsWriterインスタンスを生成する。

    未設定時はNoneを返す（Google Sheets連携が無効な状態）。
    """
    credentials_path = os.environ.get("GOOGLE_SHEETS_CREDENTIALS_PATH", "")
    if not credentials_path:
        return None

    # eval-appディレクトリからの相対パスを解決
    if not os.path.isabs(credentials_path):
        base_dir = Path(__file__).parent.parent
        credentials_path = str(base_dir / credentials_path)

    if not os.path.exists(credentials_path):
        logger.warning(f"認証ファイルが見つかりません: {credentials_path}")
        return None

    spreadsheet_id = os.environ.get("EVAL_SPREADSHEET_ID", "") or None
    share_email = os.environ.get("EVAL_SPREADSHEET_SHARE_EMAIL", "") or None

    try:
        return SheetsWriter(
            credentials_path=credentials_path,
            spreadsheet_id=spreadsheet_id,
            share_email=share_email,
        )
    except Exception as e:
        logger.error(f"SheetsWriter初期化エラー: {e}")
        return None
