# メディアサイト パイロット版 要件定義

---

## 概要

既存PRサイト（`site/`）の `media/` ディレクトリに、
お知らせ・コラム・Podcastを閲覧できる**独立メディアサイト**を構築する。

パイロット版は**静的HTML + GAS APIの構成**で、
まずGitHub上で動作確認できる状態を目指す。
さくらサーバーへのデプロイは既存のGitHub Actionsで自動反映される。

### スコープ

| 含む | 含まない |
|------|---------|
| メディアサイトのフロントエンド（HTML/CSS/JS） | note.com連携 |
| GAS APIからのデータ取得・表示 | さくらサーバーへの音声アップロード |
| コラム一覧・詳細（Markdown→HTML変換） | 画像アップロードAPI（PHP） |
| Podcastページ（外部配信リンク + QRコード） | GAS管理画面の改修 |
| お知らせ一覧 | ステータスフロー改定（別途実装） |
| サンプルデータでの表示確認 | |

### 対象外（後続フェーズ）

- ステータスフロー改定（予約確定・最終完了）→ 別要件として実装
- 画像アップロードAPI（さくらPHP）→ 画像はURL直指定で対応
- note.com RSS連携 → スコープ外
- GAS側の記事管理画面拡張 → 既存の管理画面をそのまま利用

---

## 1. URL構成

```
site/media/
├── index.html          — トップページ（最新ニュース + コラム + Podcast）
├── news.html           — お知らせ一覧
├── columns.html        — コラム一覧（カテゴリフィルタ付き）
├── column.html?id=A001 — コラム詳細
├── podcast.html        — Podcast一覧（外部配信リンク + QRコード）
├── css/
│   └── media.css       — メディアサイト専用CSS
└── js/
    └── media.js        — API連携・ページ描画・Markdown変換
```

---

## 2. ページ設計

### 2-1. トップページ（`index.html`）

```
┌──────────────────────────────────────────────────┐
│ [ヘッダー] PRサイト共通                            │
│  ホーム / 研究会について / メディア / 相談申込       │
├──────────────────────────────────────────────────┤
│                                                  │
│  MEDIA & INSIGHTS                                │
│  お知らせ・コラム・Podcast                         │
│                                                  │
├──────────────────────────────────────────────────┤
│                                                  │
│  ── NEWS ──                                      │
│  2026.04.01  無料経営相談の本格運用を開始しました  │
│  2026.03.20  コラム「経営課題...」を公開しました   │
│  2026.03.15  ...                                 │
│                        [お知らせ一覧へ →]          │
│                                                  │
├──────────────────────────────────────────────────┤
│                                                  │
│  ── COLUMN ──                                    │
│  ┌──────────┐ ┌──────────┐ ┌──────────┐         │
│  │ サムネイル │ │ サムネイル │ │ サムネイル │         │
│  │ カテゴリ  │ │ カテゴリ  │ │ カテゴリ  │         │
│  │ タイトル  │ │ タイトル  │ │ タイトル  │         │
│  │ 日付     │ │ 日付     │ │ 日付     │         │
│  └──────────┘ └──────────┘ └──────────┘         │
│                        [コラム一覧へ →]            │
│                                                  │
├──────────────────────────────────────────────────┤
│                                                  │
│  ── PODCAST ──                                   │
│  🎧 最新エピソード                                │
│  「第3回：資金繰り改善の5つのポイント」            │
│  [Spotify] [Apple Podcast] [YouTube]             │
│                        [Podcast一覧へ →]          │
│                                                  │
├──────────────────────────────────────────────────┤
│ [CTA] 無料経営相談のお申し込み                     │
├──────────────────────────────────────────────────┤
│ [フッター] PRサイト共通                            │
└──────────────────────────────────────────────────┘
```

### 2-2. お知らせ一覧（`news.html`）

```
┌──────────────────────────────────────────────────┐
│ NEWS                                             │
│ お知らせ                                         │
│                                                  │
│ 2026.04.01  無料経営相談の本格運用を開始しました    │
│ ─────────────────────────────────                │
│ 2026.03.20  コラム「経営課題...」を公開しました     │
│ ─────────────────────────────────                │
│ 2026.03.15  ...                                  │
│                                                  │
│ ← GAS API（getLatestNews）から取得               │
│   表示件数: 全件（ページネーション付き）            │
└──────────────────────────────────────────────────┘
```

### 2-3. コラム一覧（`columns.html`）

```
┌──────────────────────────────────────────────────┐
│ COLUMN                                           │
│ コラム・記事                                      │
│                                                  │
│ [全て] [経営戦略] [財務・会計] [コラム] [DX・IT]   │
│ ← カテゴリフィルタ（GAS APIのカテゴリ一覧から生成） │
│                                                  │
│ ┌───────────────────────────────────────────┐    │
│ │ [サムネイル]  タイトル                     │    │
│ │              経営戦略 | 著者名 | 2026.04.01│    │
│ │              要約テキスト...               │    │
│ └───────────────────────────────────────────┘    │
│ ┌───────────────────────────────────────────┐    │
│ │ [サムネイル]  タイトル                     │    │
│ │              ...                          │    │
│ └───────────────────────────────────────────┘    │
│                                                  │
│ [1] [2] [3] ... ← ページネーション               │
│ ← GAS API（getArticles）から取得                 │
└──────────────────────────────────────────────────┘
```

### 2-4. コラム詳細（`column.html?id=A001`）

```
┌──────────────────────────────────────────────────┐
│ ← コラム一覧に戻る                                │
│                                                  │
│ 経営戦略                                         │
│                                                  │
│ 中小企業の経営課題と診断士の役割                    │
│ ═══════════════════════════════                   │
│ 中小企業経営診断研究会  |  2026.04.01              │
│                                                  │
│ [サムネイル画像]                                  │
│                                                  │
│ ## はじめに                          ← Markdown  │
│ 本文テキスト...                       → HTML変換  │
│                                                  │
│ ![写真](https://...)                             │
│                                                  │
│ 続き本文...                                      │
│                                                  │
│ ┌─ Podcastでも聴けます ─────────────────────┐    │
│ │ [Spotify] [Apple Podcast] [YouTube]       │    │
│ │                          [QRコード]       │    │
│ └───────────────────────────────────────────┘    │
│ ↑ 記事にPodcast連携がある場合のみ表示             │
│                                                  │
│ [← 前の記事]              [次の記事 →]            │
│                                                  │
│ ← GAS API（getArticleById）から取得              │
│   Markdown→HTML変換はフロントエンド（marked.js）   │
└──────────────────────────────────────────────────┘
```

### 2-5. Podcast一覧（`podcast.html`）

```
┌──────────────────────────────────────────────────┐
│ PODCAST                                          │
│ 経営のヒント                                      │
│                                                  │
│ ┌───────────────────────────────────────────┐    │
│ │ [カバー画像]                               │    │
│ │                                            │    │
│ │ 第3回「資金繰り改善の5つのポイント」         │    │
│ │ 2026.04.01                                 │    │
│ │                                            │    │
│ │ 説明テキスト...                             │    │
│ │                                            │    │
│ │ 配信プラットフォーム:                        │    │
│ │ [🟢 Spotify] [🍎 Apple] [▶ YouTube]       │    │
│ │                                            │    │
│ │ ┌──────┐                                   │    │
│ │ │ QR   │ ← 主要配信リンクのQRコード         │    │
│ │ │ コード│                                   │    │
│ │ └──────┘                                   │    │
│ │                                            │    │
│ │ 📎 関連コラム: 「資金繰り改善の...」         │    │
│ └───────────────────────────────────────────┘    │
│                                                  │
│ ── 過去のエピソード ──                            │
│ 第2回「DX推進のポイント」     2026.03.15          │
│ 第1回「経営課題と診断士」     2026.03.01          │
│                                                  │
│ ← GAS API（getPodcasts）から取得                 │
│   QRコードは api.qrserver.com で動的生成          │
└──────────────────────────────────────────────────┘
```

---

## 3. データソース

### 3-1. 既存のGAS APIを活用

パイロット版ではGAS側の**コード変更は最小限**。既存APIをそのまま利用する。

| データ | GAS API | 既存関数 | 備考 |
|--------|---------|----------|------|
| お知らせ | `?action=news` | `getLatestNews()` | 既存のまま利用 |
| コラム一覧 | `?action=articles` | `getArticles()` | 既存のまま利用 |
| コラム詳細 | `?action=article&id=A001` | `getArticleById()` | 既存のまま利用 |
| カテゴリ一覧 | `?action=article-categories` | `getArticleCategories()` | 既存のまま利用 |
| Podcast一覧 | `?action=podcasts` | **新規作成** | Podcastシート用 |
| Podcast詳細 | `?action=podcast&id=P001` | **新規作成** | Podcastシート用 |

### 3-2. GAS側の追加（最小限）

#### Podcastシート（新規）

スプレッドシートに「Podcast」シートを新設:

| 列 | 項目 | 説明 |
|----|------|------|
| A | エピソードID | `P001`, `P002`, ... |
| B | タイトル | エピソード名 |
| C | 説明 | エピソードの説明文 |
| D | 公開日 | `yyyy.MM.dd` |
| E | 状態 | `draft` / `published` |
| F | Spotify URL | 外部配信リンク |
| G | Apple Podcast URL | 外部配信リンク |
| H | YouTube URL | 外部配信リンク |
| I | サムネイルURL | カバー画像のURL |
| J | 関連記事ID | 記事管理シートの記事IDと紐付け（任意） |
| K | 作成日時 | 自動記録 |

※ 音声ファイルのアップロード・ホスティングは行わない。
  外部配信プラットフォーム（Spotify, Apple Podcast, YouTube等）のリンクのみ。

#### 記事 × Podcast連携

記事管理シートの既存K列（音声ファイルID、現在未使用）を転用:

| 列 | 現状 | パイロット版 |
|----|------|-------------|
| K | 音声ファイルID（未使用） | **PodcastエピソードID**（例: `P001`） |

コラム詳細ページでK列にエピソードIDがある場合、Podcastセクションを表示。

#### 新規GASファイル: `gas/podcast.gs`

```
機能:
- setupPodcastSheet()     — シートのセットアップ
- getPodcasts()           — 一覧取得（published のみ）
- getPodcastById(id)      — 詳細取得
- addPodcast(params)      — エピソード追加
- generatePodcastAdminPage() — 管理画面
```

#### `gas/main.gs` への追加

```
新規アクション:
- podcasts     → getPodcasts()
- podcast      → getPodcastById(id)
- podcast-admin → generatePodcastAdminPage()
```

---

## 4. フロントエンド技術仕様

### 4-1. 依存ライブラリ

| ライブラリ | 用途 | CDN |
|-----------|------|-----|
| **marked.js** | Markdown→HTML変換 | `https://cdn.jsdelivr.net/npm/marked/marked.min.js` |
| — | その他の外部ライブラリは不使用 | — |

Font Awesome、Google Fonts は既存PRサイトと同じCDNを使用。

### 4-2. media.js の構成

```javascript
// 定数
const API_URL = '（GAS Web App URL）';

// API呼び出し
async function fetchAPI(action, params) { ... }

// ページ描画
function renderTopPage() { ... }        // index.html 用
function renderNewsPage() { ... }       // news.html 用
function renderColumnsPage() { ... }    // columns.html 用
function renderColumnDetail(id) { ... } // column.html 用
function renderPodcastPage() { ... }    // podcast.html 用

// Markdown変換
function renderMarkdown(md) { ... }     // marked.js ラッパー

// QRコード生成
function getQRCodeURL(url) { ... }      // api.qrserver.com を利用

// ユーティリティ
function formatDate(str) { ... }
function escapeHTML(str) { ... }
function truncate(str, len) { ... }
function getParam(name) { ... }         // URLパラメータ取得
```

### 4-3. デザインシステム

既存 `site/css/style.css` を読み込んだ上で、`media.css` で追加スタイルを定義:

```
media.css で追加するスタイル:
- .media-hero          — ページヒーローセクション
- .news-list           — ニュース一覧
- .news-item           — ニュース1件
- .column-grid         — コラムカードのグリッド
- .column-card         — コラムカード
- .column-card-thumb   — サムネイル画像
- .column-card-meta    — カテゴリ・著者・日付
- .column-detail       — コラム詳細の本文エリア
- .column-detail img   — 本文内画像のレスポンシブ対応
- .category-filter     — カテゴリフィルタバー
- .category-btn        — カテゴリボタン
- .podcast-card        — Podcastエピソードカード
- .podcast-platforms   — 配信プラットフォームリンク
- .podcast-qr          — QRコード表示エリア
- .pagination          — ページネーション
- .loading-spinner     — ローディング表示
- .empty-state         — データなし時の表示
```

### 4-4. QRコード

```javascript
function getQRCodeURL(targetUrl) {
    return 'https://api.qrserver.com/v1/create-qr-code/?size=200x200&data='
         + encodeURIComponent(targetUrl);
}
```

表示は `<img>` タグで動的生成。保存不要。

---

## 5. ナビゲーション

### 5-1. メディアサイト内ナビ

共通ヘッダーにメディアサイトのナビゲーションを追加:

```html
<nav>
    <ul class="nav-list">
        <li><a href="../../index15.html">ホーム</a></li>
        <li><a href="../about.html">研究会について</a></li>
        <li><a href="index.html">メディア</a></li>
        <li class="nav-cta"><a href="../../index15.html#contact">相談申込</a></li>
    </ul>
</nav>
```

メディアサイト内のサブナビ（ページ上部）:

```html
<div class="media-subnav">
    <a href="news.html">お知らせ</a>
    <a href="columns.html">コラム</a>
    <a href="podcast.html">Podcast</a>
</div>
```

### 5-2. 既存サイトへの影響（Phase 1では最小限）

Phase 1（パイロット版）では既存ページのナビは**変更しない**。
メディアサイトからPRサイト・LPへのリンクは設置する（片方向）。

---

## 6. サンプルデータ

パイロット版の表示確認用に、GAS管理画面から以下を登録:

### お知らせ（既存シートに追加）
既存の `getLatestNews()` で取得可能なデータがあれば、そのまま利用。

### コラム（既存の記事管理シートのサンプル記事）
`setupArticlesSheet()` で投入済みのサンプル2件:
- A001: 「中小企業の経営課題と診断士の役割」（経営戦略）
- A002: 「資金繰り改善のための5つのポイント」（財務・会計）

### Podcast（新規シートにサンプル投入）
`setupPodcastSheet()` で以下を投入:
- P001: 「第1回：経営課題と診断士の役割」
  - Spotify / Apple Podcast / YouTube のダミーURL
  - 関連記事: A001

---

## 7. 実装ファイル一覧

### 新規作成

| ファイル | 内容 |
|----------|------|
| `site/media/index.html` | メディアサイト トップページ |
| `site/media/news.html` | お知らせ一覧ページ |
| `site/media/columns.html` | コラム一覧ページ |
| `site/media/column.html` | コラム詳細ページ |
| `site/media/podcast.html` | Podcast一覧ページ |
| `site/media/css/media.css` | メディアサイト専用スタイル |
| `site/media/js/media.js` | API連携・ページ描画・Markdown変換 |
| `gas/podcast.gs` | Podcastシート管理（CRUD + 管理画面） |

### 修正（最小限）

| ファイル | 変更内容 |
|----------|----------|
| `gas/main.gs` | `podcasts`, `podcast`, `podcast-admin` アクション追加 |

### 変更不要

| ファイル | 理由 |
|----------|------|
| `gas/articles.gs` | 既存API（getArticles, getArticleById）をそのまま利用 |
| `gas/news.gs` | 既存API（getLatestNews）をそのまま利用 |
| `gas/config.gs` | パイロット版では変更不要 |
| `site/css/style.css` | メディアサイトから読み込むのみ |
| `site/js/main.js` | メディアサイトから読み込むのみ |
| `.github/workflows/deploy.yml` | `site/` 配下は既にデプロイ対象 |
| 既存のLP・PRサイト | Phase 1では一切変更しない |

---

## 8. デプロイ

既存の GitHub Actions（`.github/workflows/deploy.yml`）で `site/` 以下は
自動的にさくらサーバーにデプロイされるため、**追加設定は不要**。

```
git push → GitHub Actions → rsync → iba-consulting.jp/site/media/
```

パイロット版の確認:
1. ローカルでHTMLファイルを直接開いて構造確認（APIは CORS でブロックされるため表示のみ）
2. さくらサーバーにデプロイ後、`https://iba-consulting.jp/site/media/` で動作確認

---

## 9. 検証ポイント

1. `site/media/index.html` — トップページが表示され、3セクション（ニュース・コラム・Podcast）が見える
2. `site/media/news.html` — GAS APIからお知らせ一覧を取得・表示
3. `site/media/columns.html` — GAS APIからコラム一覧を取得・カテゴリフィルタが動作
4. `site/media/column.html?id=A001` — コラム詳細がMarkdown→HTMLで表示
5. `site/media/podcast.html` — Podcast一覧が表示、外部配信リンクが正しくリンク
6. QRコードが正しい配信URLで表示される
7. PRサイトと同じデザイン（ヘッダー・フッター・配色・フォント）
8. スマートフォンでのレスポンシブ表示
9. メディアサイトからLP（相談申込）への導線が機能
