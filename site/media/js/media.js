/**
 * KG中小企業経営診断研究会 メディアサイト — JS
 */

const MEDIA_API_URL = 'https://script.google.com/macros/s/AKfycbzR7l1lyRF9dNZ0qqIov8LZwxDvkkyT4NNo2LSJKbQR_i46iqLfSRg4EuqQRflP76elAg/exec';
const LP_URL = 'https://iba-consulting.jp/index15.html';

/**
 * GAS APIからデータ取得
 */
async function mediaFetch(action, params = {}) {
    const url = new URL(MEDIA_API_URL);
    url.searchParams.set('action', action);
    Object.entries(params).forEach(([k, v]) => {
        if (v !== null && v !== undefined && v !== '') url.searchParams.set(k, v);
    });
    try {
        const resp = await fetch(url.toString());
        return await resp.json();
    } catch (err) {
        console.error('API Error:', err);
        return { success: false, error: err.message };
    }
}

/**
 * QRコード画像URLを生成
 */
function getQRCodeURL(targetUrl) {
    return 'https://api.qrserver.com/v1/create-qr-code/?size=200x200&data=' + encodeURIComponent(targetUrl);
}

/**
 * URLパラメータ取得
 */
function getParam(name) {
    const params = new URLSearchParams(window.location.search);
    return params.get(name);
}

/**
 * HTMLエスケープ
 */
function esc(str) {
    if (!str) return '';
    return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/**
 * テキスト切り詰め
 */
function truncate(str, len) {
    if (!str) return '';
    return str.length > len ? str.substring(0, len) + '...' : str;
}

/**
 * Markdown→HTML（marked.js利用）
 */
function renderMarkdown(md) {
    if (!md) return '';
    if (typeof marked !== 'undefined') {
        return marked.parse(md);
    }
    // fallback: 最低限の変換
    return md
        .replace(/^### (.+)$/gm, '<h3>$1</h3>')
        .replace(/^## (.+)$/gm, '<h2>$1</h2>')
        .replace(/^# (.+)$/gm, '<h1>$1</h1>')
        .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
        .replace(/!\[([^\]]*)\]\(([^)]+)\)/g, '<img src="$2" alt="$1">')
        .replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2">$1</a>')
        .replace(/^> (.+)$/gm, '<blockquote>$1</blockquote>')
        .replace(/^- (.+)$/gm, '<li>$1</li>')
        .replace(/\n\n/g, '</p><p>')
        .replace(/^/, '<p>').replace(/$/, '</p>');
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// トップページ
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

async function renderTopPage() {
    const content = document.getElementById('media-content');
    if (!content) return;

    content.innerHTML = '<div class="media-loading"><i class="fas fa-spinner fa-spin"></i> 読み込み中...</div>';

    const [newsRes, articlesRes, podcastsRes] = await Promise.all([
        mediaFetch('news'),
        mediaFetch('articles', { limit: '3' }),
        mediaFetch('podcasts', { limit: '1' })
    ]);

    let html = '';

    // ニュースセクション
    html += '<div class="media-section">';
    html += '<div class="media-section-header"><h2>News</h2><a href="news.html">お知らせ一覧 &rarr;</a></div>';
    if (newsRes.success && newsRes.news && newsRes.news.length > 0) {
        html += '<ul class="news-list">';
        newsRes.news.slice(0, 5).forEach(n => {
            html += `<li class="news-item"><span class="news-date">${esc(n.date)}</span><span class="news-text">${esc(n.content)}</span></li>`;
        });
        html += '</ul>';
    } else {
        html += '<p class="empty-state">お知らせはまだありません</p>';
    }
    html += '</div>';

    // コラムセクション
    html += '<div class="media-section">';
    html += '<div class="media-section-header"><h2>Column</h2><a href="columns.html">コラム一覧 &rarr;</a></div>';
    if (articlesRes.success && articlesRes.articles && articlesRes.articles.length > 0) {
        html += '<div class="column-grid">';
        articlesRes.articles.forEach(a => {
            const thumb = a.thumbnail
                ? `<img src="${esc(a.thumbnail)}" alt="${esc(a.title)}">`
                : '<i class="fas fa-file-alt"></i>';
            html += `<a class="column-card" href="column.html?id=${esc(a.id)}">
                <div class="column-card-thumb">${thumb}</div>
                <div class="column-card-body">
                    <div class="column-card-category">${esc(a.category)}</div>
                    <div class="column-card-title">${esc(a.title)}</div>
                    <div class="column-card-meta">
                        <span>${esc(a.author)}</span>
                        <span>${esc(a.publishDate)}</span>
                    </div>
                </div>
            </a>`;
        });
        html += '</div>';
    } else {
        html += '<p class="empty-state">コラムはまだありません</p>';
    }
    html += '</div>';

    // Podcastセクション
    html += '<div class="media-section">';
    html += '<div class="media-section-header"><h2>Podcast</h2><a href="podcast.html">Podcast一覧 &rarr;</a></div>';
    if (podcastsRes.success && podcastsRes.podcasts && podcastsRes.podcasts.length > 0) {
        const p = podcastsRes.podcasts[0];
        html += '<div class="podcast-card">';
        html += '<div class="podcast-card-header">';
        html += `<div class="podcast-card-thumb">${p.thumbnail ? '<img src="' + esc(p.thumbnail) + '" alt="">' : '<i class="fas fa-podcast"></i>'}</div>`;
        html += '<div class="podcast-card-info">';
        html += `<h3>${esc(p.title)}</h3>`;
        html += `<div class="date">${esc(p.publishDate)}</div>`;
        html += '</div></div>';
        if (p.description) {
            html += `<div class="podcast-card-description">${esc(p.description)}</div>`;
        }
        html += renderPlatformLinks(p);
        html += '</div>';
    } else {
        html += '<p class="empty-state">Podcastはまだありません</p>';
    }
    html += '</div>';

    content.innerHTML = html;
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// お知らせ一覧
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

async function renderNewsPage() {
    const content = document.getElementById('media-content');
    if (!content) return;

    content.innerHTML = '<div class="media-loading"><i class="fas fa-spinner fa-spin"></i> 読み込み中...</div>';

    const res = await mediaFetch('news');

    let html = '';
    if (res.success && res.news && res.news.length > 0) {
        html += '<ul class="news-list">';
        res.news.forEach(n => {
            html += `<li class="news-item"><span class="news-date">${esc(n.date)}</span><span class="news-text">${esc(n.content)}</span></li>`;
        });
        html += '</ul>';
    } else {
        html += '<div class="empty-state"><i class="fas fa-bell-slash"></i>お知らせはまだありません</div>';
    }

    content.innerHTML = html;
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// コラム一覧
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

let currentCategory = null;
let currentPage = 0;
const PAGE_SIZE = 10;

async function renderColumnsPage() {
    const content = document.getElementById('media-content');
    if (!content) return;

    content.innerHTML = '<div class="media-loading"><i class="fas fa-spinner fa-spin"></i> 読み込み中...</div>';

    // カテゴリフィルタ
    const catRes = await mediaFetch('article-categories');
    let filterHtml = '<div class="filter-bar">';
    filterHtml += `<button class="filter-btn ${!currentCategory ? 'active' : ''}" onclick="filterCategory(null)">全て</button>`;
    if (catRes.success && catRes.categories) {
        catRes.categories.forEach(c => {
            const isActive = currentCategory === c.name ? 'active' : '';
            filterHtml += `<button class="filter-btn ${isActive}" onclick="filterCategory('${esc(c.name)}')">${esc(c.name)} (${c.count})</button>`;
        });
    }
    filterHtml += '</div>';

    // 記事取得
    const res = await mediaFetch('articles', {
        category: currentCategory,
        limit: String(PAGE_SIZE),
        offset: String(currentPage * PAGE_SIZE)
    });

    let listHtml = '';
    if (res.success && res.articles && res.articles.length > 0) {
        res.articles.forEach(a => {
            const thumb = a.thumbnail
                ? `<img src="${esc(a.thumbnail)}" alt="${esc(a.title)}">`
                : '<i class="fas fa-file-alt"></i>';
            listHtml += `<div class="column-list-item" onclick="location.href='column.html?id=${esc(a.id)}'">
                <div class="column-list-thumb">${thumb}</div>
                <div class="column-list-body">
                    <div class="column-card-category">${esc(a.category)}</div>
                    <div class="column-card-title">${esc(a.title)}</div>
                    <div class="column-card-meta">
                        <span>${esc(a.author)}</span>
                        <span>${esc(a.publishDate)}</span>
                    </div>
                    ${a.summary ? `<div class="column-card-summary">${esc(a.summary)}</div>` : ''}
                </div>
            </div>`;
        });

        // ページネーション
        const totalPages = Math.ceil(res.total / PAGE_SIZE);
        if (totalPages > 1) {
            listHtml += '<div class="pagination">';
            for (let i = 0; i < totalPages; i++) {
                if (i === currentPage) {
                    listHtml += `<span class="current">${i + 1}</span>`;
                } else {
                    listHtml += `<a href="#" onclick="goPage(${i}); return false;">${i + 1}</a>`;
                }
            }
            listHtml += '</div>';
        }
    } else {
        listHtml = '<div class="empty-state"><i class="fas fa-folder-open"></i>記事はまだありません</div>';
    }

    content.innerHTML = filterHtml + listHtml;
}

function filterCategory(cat) {
    currentCategory = cat;
    currentPage = 0;
    renderColumnsPage();
}

function goPage(page) {
    currentPage = page;
    renderColumnsPage();
    window.scrollTo(0, 0);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// コラム詳細
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

async function renderColumnDetail() {
    const content = document.getElementById('media-content');
    if (!content) return;

    const id = getParam('id');
    if (!id) {
        content.innerHTML = '<div class="empty-state"><i class="fas fa-exclamation-triangle"></i>記事IDが指定されていません</div>';
        return;
    }

    content.innerHTML = '<div class="media-loading"><i class="fas fa-spinner fa-spin"></i> 読み込み中...</div>';

    const res = await mediaFetch('article', { id: id });

    if (!res.success || !res.article) {
        content.innerHTML = '<div class="empty-state"><i class="fas fa-exclamation-triangle"></i>記事が見つかりません</div>';
        return;
    }

    const a = res.article;
    let html = '';

    // パンくず
    html += `<div class="breadcrumb">
        <a href="columns.html">コラム一覧</a>
        <span>/</span>
        <span>${esc(a.title)}</span>
    </div>`;

    // 記事本文
    html += '<article class="article-content">';
    html += `<div class="article-meta">
        <span class="tag">${esc(a.category)}</span>
        <span>${esc(a.author)}</span>
        <span>${esc(a.publishDate)}</span>
    </div>`;
    html += `<h1>${esc(a.title)}</h1>`;

    if (a.thumbnail) {
        html += `<img src="${esc(a.thumbnail)}" alt="${esc(a.title)}" style="border-radius: 12px; margin-bottom: 2rem;">`;
    }

    // Markdown本文
    html += renderMarkdown(a.body);

    // Podcast連携セクション（記事のaudioFileId = PodcastエピソードID）
    if (a.audioFileId) {
        html += await renderColumnPodcast(a.audioFileId);
    }

    html += '</article>';

    // 前後の記事ナビ
    html += `<div class="column-nav">
        <a href="columns.html">&larr; コラム一覧に戻る</a>
    </div>`;

    content.innerHTML = html;

    // ページタイトル更新
    document.title = a.title + ' | メディア | 関西学院大学 中小企業経営診断研究会';
}

async function renderColumnPodcast(episodeId) {
    const res = await mediaFetch('podcast', { id: episodeId });
    if (!res.success || !res.podcast) return '';

    const p = res.podcast;
    let html = '<div class="column-podcast-section">';
    html += '<h4><i class="fas fa-podcast" style="margin-right:0.5rem; color:var(--kgu-gold);"></i>Podcastでも聴けます</h4>';
    html += renderPlatformLinks(p);
    html += '</div>';
    return html;
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// Podcast一覧
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

async function renderPodcastPage() {
    const content = document.getElementById('media-content');
    if (!content) return;

    content.innerHTML = '<div class="media-loading"><i class="fas fa-spinner fa-spin"></i> 読み込み中...</div>';

    const res = await mediaFetch('podcasts', { limit: '50' });

    if (!res.success || !res.podcasts || res.podcasts.length === 0) {
        content.innerHTML = '<div class="empty-state"><i class="fas fa-podcast"></i>Podcastエピソードはまだありません</div>';
        return;
    }

    let html = '';
    const latest = res.podcasts[0];
    const rest = res.podcasts.slice(1);

    // 最新エピソード（カード表示）
    html += '<div class="podcast-card">';
    html += '<div class="podcast-card-header">';
    html += `<div class="podcast-card-thumb">${latest.thumbnail ? '<img src="' + esc(latest.thumbnail) + '" alt="">' : '<i class="fas fa-podcast"></i>'}</div>`;
    html += '<div class="podcast-card-info">';
    html += `<h3>${esc(latest.title)}</h3>`;
    html += `<div class="date">${esc(latest.publishDate)}</div>`;
    html += '</div></div>';
    if (latest.description) {
        html += `<div class="podcast-card-description">${esc(latest.description)}</div>`;
    }
    html += renderPlatformLinks(latest);

    // QRコード
    const qrUrl = latest.spotifyUrl || latest.appleUrl || latest.youtubeUrl;
    if (qrUrl) {
        html += '<div style="display:flex; gap:1.5rem; align-items:start; margin-top:1rem;">';
        html += `<div class="podcast-qr"><img src="${getQRCodeURL(qrUrl)}" alt="QRコード"><div class="podcast-qr-label">QRコードで聴く</div></div>`;
        html += '</div>';
    }

    // 関連コラム
    if (latest.relatedArticleId) {
        html += `<div class="podcast-related"><i class="fas fa-link" style="margin-right:0.5rem;"></i>関連コラム: <a href="column.html?id=${esc(latest.relatedArticleId)}">${esc(latest.relatedArticleId)}</a></div>`;
    }
    html += '</div>';

    // 過去のエピソード
    if (rest.length > 0) {
        html += '<h3 style="margin: 2rem 0 1rem; font-size: 0.85rem; letter-spacing: 0.15em; color: var(--text-muted); text-transform: uppercase;">過去のエピソード</h3>';
        rest.forEach(p => {
            const platforms = [];
            if (p.spotifyUrl) platforms.push('Spotify');
            if (p.appleUrl) platforms.push('Apple');
            if (p.youtubeUrl) platforms.push('YouTube');
            html += `<div class="podcast-list-item">
                <div>
                    <div class="podcast-list-title">${esc(p.title)}</div>
                    ${platforms.length > 0 ? '<div style="font-size:0.75rem; color:var(--text-muted); margin-top:0.2rem;">' + platforms.join(' / ') + '</div>' : ''}
                </div>
                <div class="podcast-list-date">${esc(p.publishDate)}</div>
            </div>`;
        });
    }

    content.innerHTML = html;
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 共通: プラットフォームリンク描画
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function renderPlatformLinks(p) {
    let html = '<div class="podcast-platforms">';
    if (p.spotifyUrl) {
        html += `<a href="${esc(p.spotifyUrl)}" target="_blank" rel="noopener" class="podcast-platform-btn spotify"><i class="fab fa-spotify"></i> Spotify</a>`;
    }
    if (p.appleUrl) {
        html += `<a href="${esc(p.appleUrl)}" target="_blank" rel="noopener" class="podcast-platform-btn apple"><i class="fab fa-apple"></i> Apple Podcast</a>`;
    }
    if (p.youtubeUrl) {
        html += `<a href="${esc(p.youtubeUrl)}" target="_blank" rel="noopener" class="podcast-platform-btn youtube"><i class="fab fa-youtube"></i> YouTube</a>`;
    }
    if (!p.spotifyUrl && !p.appleUrl && !p.youtubeUrl) {
        html += '<span style="font-size:0.85rem; color:var(--text-muted);">配信リンクは準備中です</span>';
    }
    html += '</div>';
    return html;
}
