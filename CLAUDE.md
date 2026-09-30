# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

すべての UI・メール・ドキュメントは日本語です。

## Project Overview

**KG無料経営相談システム** — 関西学院大学 中小企業経営診断研究会（IBA）が運営する無料経営相談の
予約・運用自動化システム。公開サイトは https://iba-consulting.jp （さくらレンタルサーバー）。

中小企業経営者からの相談申込を受け付け、NDA同意 → 日程調整 → 担当者自動選定 → Zoom/対面での相談 →
録画・文字起こし・報告書 → アンケート・評価 まで一気通貫で扱う。運用の中核データは Google スプレッドシート
「予約管理」（所有者 `kgibaconsultant@gmail.com`）。

> 注意：`stock_backtest/`, `calculator.py`, `todo_app.py`, `hello_world.py` はこのシステムとは無関係な
> 過去の練習用コードです（本システムの構成要素ではありません）。

## Architecture

### 1. 静的サイト（さくらレンタルサーバー / iba-consulting.jp）
HTML/CSS/JS を SCP で配置。接続情報はメモリ `reference_sakura_server` を参照。

- `index.html` → `index15.html`（相談予約LP）へリダイレクト。`news.html`（お知らせ）
- `site/` — 組織紹介サイト・コラム・スタッフポータル（`site/portal.html`）
- `docs/` — マニュアル（`manual_consultee.html` / `manual_staff.html` / `manual_admin.html`）、`system_overview.html`、`CHANGELOG.yml`
- `consent_updated.pdf`（相談者NDA）, `observer_nda.pdf`（オブザーバーNDA） — GASから参照
- GA4 導入済み（測定ID: G-GELJ4GB33D）

### 2. GAS バックエンド（`gas/`） — 中核
Google Apps Script の Web App（clasp プロジェクト）。LP や各画面からの API を `doGet`/`doPost` で処理。

- clasp 設定: `gas/.clasp.json`（scriptId `1GCCYhyzkuSzJuOfapiUWzgHHoaq39pAIV7k1zo6M-GBiPvdX78dEB0Uu`）
- 本番 exec URL は固定デプロイID `AKfycbzR7l1lyRF9dNZ0qqIov8LZwxDvkkyT4NNo2LSJKbQR_i46iqLfSRg4EuqQRflP76elAg`
- 設定は `config.gs` の `CONFIG`（`SPREADSHEET_ID`, 管理者メール, Zoom/YouTube/文字起こしCF, フォームポーリング日程 等）。
  シークレットは GAS の **スクリプト プロパティ**（例: `ADMIN_API_TOKEN`, `ZOOM_*`, `*_CF_SECRET`）。

主なファイル（`.gs`）:
- `main.gs` — `doGet`/`doPost` のルーティング（`action` パラメータで分岐、124+アクション）
- `config.gs` — CONFIG・列定義（`COLUMNS` 等）・ステータス定数
- `spreadsheet.gs` — 予約管理シートの読み書き
- `form_polling.gs` — 月次の日程アンケート（オプトアウト方式）・Googleフォーム生成・回答集計
- `schedule.gs` / `members.gs` / `leader.gs` — 日程・メンバー・担当者/リーダー自動選定
- `zoom.gs` / `transcript.gs` / `report.gs` / `podcast.gs` / `articles.gs` / `venue.gs` / `survey.gs` / `evaluation.gs`
- `portal.gs` — スタッフポータル（マジックリンク＋セッショントークン認証）
- `triggers.gs` — 時限/onEdit/onFormSubmit トリガーの設定（`setupAllTriggers`）
- `templates.gs` — メール本文テンプレート

### 3. Cloud Functions（`cloud_functions/`）
GCP プロジェクト `kg-consultation`（asia-northeast1）:
- `zoom-to-youtube` — Zoom録画を YouTube に限定公開アップロード
- `zoom-to-transcript` — Zoom録画を文字起こし（一時ファイルは GCS `kg-consultation-transcript-temp`、1日で自動削除）

### 中核スプレッドシート「予約管理」（`1tD6-0WVXsud8_APnR4IxlI7yR085I47WoTjiFZ8oFBc`）
主なシート: `予約管理`（申込・ステータス）, `メンバーマスタ`, `日程設定`, `回答集計`, `お知らせ`,
`アンケート`, `会場マスタ`, `オブザーバーNDA`, `リーダー履歴`, `レポート管理`, `コンサルタント評価`,
`Podcast`, `記事管理`。列定義は `config.gs` の各 `*_COLUMNS` を正とする。

予約ステータス遷移: `仮予約 → 同意済 → 書類受領 → 確定 → 完了`（または `キャンセル`）。

## Web App API とセキュリティ

`doGet` は `action` で分岐。おおまかに:
- **公開アクション**（LP/サイト/ポータルが使用）: `news` `members` `schedule` `articles` `article` `podcasts`
  `venues` `survey` `nda` `nda-submit` `observer` `status` など、および `portal-*`
- **`portal-*`** は独自のセッショントークン（`verifyPortalLogin`）で認証
- **管理/デバッグ/保守アクション** は `ADMIN_API_TOKEN` 必須（`main.gs` の `ADMIN_ONLY_ACTIONS` に列挙、
  未設定時は fail-closed で 403）。例: `read-sheet` `row-info` `write-cell` `retry-zoom`
  `list-triggers` `cleanup-form-triggers` `setup-*` `migrate-*` 等。
  呼び出し時は `&token=<ADMIN_API_TOKEN>` を付与する。

管理画面リンクが使う `*-toggle` / `*-delete` や作成系 `*-add` は互換のため現状ゲート対象外（今後トークン化検討）。

## Commands

### GAS のデプロイ（`gas/` で実行）
clasp は認証更新で `~/.clasprc.json` に書き込むため、**サンドボックス環境では解除して実行**する。
```bash
cd gas
npx clasp status            # push対象の確認
npx clasp push -f           # ローカル .gs → プロジェクトHEAD
# 本番URLを変えずに反映するには既存デプロイIDを更新する:
npx clasp deploy -i AKfycbzR7l1lyRF9dNZ0qqIov8LZwxDvkkyT4NNo2LSJKbQR_i46iqLfSRg4EuqQRflP76elAg -d "変更内容"
```
デプロイ前に `npx clasp pull` を別ディレクトリで取得し、ローカルとの差分（＝これから本番に増える変更）を
確認すると安全（リモートHEADにしか無い編集を上書きしないため）。

### 予約データの確認
運用者以外は、Google Drive でシートを閲覧アカウントに共有 → 直接参照。
または `?action=read-sheet&sheet=予約管理&token=<ADMIN_API_TOKEN>`（管理トークン必須）。

## 運用上の重要な注意

- **新しい OAuth スコープを追加したら、GASエディタから対象関数を手動実行して承認が必要**
  （承認しないと時限トリガーが失敗する）。詳細はメモリ `feedback_gas_gmail_auth`。
- **フォーム送信トリガーの累積に注意**：月次ポーリングは毎月新しい Googleフォームを作るため、
  フォーム送信トリガー（`processFormResponse` / `processConfirmationResponse`）が溜まりやすい。
  GASのトリガー上限（1ユーザー×1スクリプトで20個）に達すると `runFirstPolling` が例外で失敗し、
  日程アンケートの自動送信が止まる。`setupFormSubmitTrigger` は同ハンドラの旧トリガーを毎回全削除する
  実装になっている。累積解消の保守用に `?action=cleanup-form-triggers&token=...` あり。
- 月次スケジュール: `runFirstPolling`（毎月25日=翌々月分アンケート）→ `runReminderPolling`（15日=確認）→
  `finalizeSchedule`（21日=日程確定）。設定は `CONFIG.FORM_POLLING`。
- 日程アンケートは「参加不可日のみ申告（回答不要なら全枠参加として自動登録）」方式。
  そのため `回答集計` の「○」は必ずしも本人の明示回答ではない（未回答＝デフォルト○）点に注意。

## Key Dependencies / 外部連携

- Google Apps Script + clasp（`@google/clasp`）
- Google スプレッドシート / Google フォーム / Gmail（`GmailApp`）/ Google Drive
- Zoom Server-to-Server OAuth（録画取得・ミーティング作成）
- Google Cloud Functions（Zoom→YouTube / Zoom→文字起こし）
- さくらレンタルサーバー（静的サイト配信、SSH/SCP）
- GA4（アクセス解析）
