/**
 * スタッフポータル SPA — portal.js
 * ハッシュベースルーティング + セッション管理
 */

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// State
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const PORTAL = {
    sessionId: null,
    session: null,       // { name, email, role }
    currentPage: null,
    shiftYear: new Date().getFullYear(),
    shiftMonth: new Date().getMonth() + 1,
    caseFilter: 'all'
};

const ROLE_LABELS = { admin: '管理者', leader: 'リーダー', member: 'メンバー' };
const STATUS_MAP = {
    '仮予約': 'tentative', 'NDA同意済': 'tentative', '書類受領': 'tentative',
    '確定': 'confirmed', '完了': 'completed', 'キャンセル': 'cancelled'
};

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// Session
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function saveSession(sessionId) {
    PORTAL.sessionId = sessionId;
    localStorage.setItem('portal_session', sessionId);
}

function loadSession() {
    return localStorage.getItem('portal_session');
}

function clearSession() {
    PORTAL.sessionId = null;
    PORTAL.session = null;
    localStorage.removeItem('portal_session');
}

async function portalFetch(action, params = {}) {
    if (PORTAL.sessionId) params.sessionId = PORTAL.sessionId;
    return fetchAPI(action, params);
}

async function validateCurrentSession() {
    const sid = loadSession();
    if (!sid) return false;

    PORTAL.sessionId = sid;
    const res = await portalFetch('portal-session');
    if (res.success && res.session) {
        PORTAL.session = res.session;
        return true;
    }
    clearSession();
    return false;
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// Router
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function getRoute() {
    const hash = window.location.hash || '#/login';
    const [path, queryStr] = hash.substring(1).split('?');
    const params = {};
    if (queryStr) {
        queryStr.split('&').forEach(p => {
            const [k, v] = p.split('=');
            params[decodeURIComponent(k)] = decodeURIComponent(v || '');
        });
    }
    return { path, params };
}

async function router() {
    const { path, params } = getRoute();
    const main = document.getElementById('portalMain');
    const nav = document.getElementById('portalNav');
    const headerRight = document.getElementById('portalHeaderRight');

    // Handle magic link verification redirect
    if (path === '/verified') {
        if (params.sessionId) {
            saveSession(params.sessionId);
            // Validate and get session data
            const valid = await validateCurrentSession();
            if (valid) {
                window.location.hash = '#/dashboard';
                return;
            }
        }
        window.location.hash = '#/login';
        return;
    }

    // Check auth for protected pages
    if (path !== '/login') {
        const valid = await validateCurrentSession();
        if (!valid) {
            window.location.hash = '#/login';
            return;
        }
        // Show portal chrome
        nav.style.display = '';
        headerRight.style.display = '';
        document.getElementById('portalUserName').textContent = PORTAL.session.name;
        document.getElementById('portalRoleBadge').textContent = ROLE_LABELS[PORTAL.session.role] || PORTAL.session.role;

        // Update active nav
        document.querySelectorAll('.portal-nav-item').forEach(a => {
            a.classList.toggle('active', a.dataset.page === path.replace('/', ''));
        });
    } else {
        nav.style.display = 'none';
        headerRight.style.display = 'none';
    }

    // Route
    PORTAL.currentPage = path;
    switch (path) {
        case '/login': renderLogin(main); break;
        case '/dashboard': await renderDashboard(main); break;
        case '/shifts': await renderShifts(main); break;
        case '/cases': await renderCases(main); break;
        case '/evaluations': await renderEvaluations(main); break;
        case '/evaluation-detail': await renderEvaluationDetail(main, params); break;
        case '/news': await renderNews(main); break;
        case '/profile': await renderProfile(main); break;
        default: window.location.hash = '#/dashboard';
    }
}

window.addEventListener('hashchange', router);

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// Login
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function renderLogin(container) {
    container.innerHTML = `
        <div class="login-page">
            <div class="login-card">
                <div class="login-icon"><i class="fas fa-user-shield"></i></div>
                <h2>スタッフポータル</h2>
                <p>メンバー登録済みのメールアドレスを入力してください。<br>ログイン用リンクをお送りします。</p>
                <form id="loginForm" onsubmit="handleLogin(event)">
                    <input type="email" class="login-input" id="loginEmail" placeholder="メールアドレス" required>
                    <button type="submit" class="login-btn" id="loginBtn">ログインリンクを送信</button>
                </form>
                <div class="login-msg" id="loginMsg"></div>
            </div>
        </div>
    `;
}

async function handleLogin(e) {
    e.preventDefault();
    const email = document.getElementById('loginEmail').value.trim();
    const btn = document.getElementById('loginBtn');
    const msg = document.getElementById('loginMsg');

    if (!email) return;

    btn.disabled = true;
    btn.textContent = '送信中...';
    msg.className = 'login-msg';
    msg.style.display = 'none';

    const res = await fetchAPI('portal-login', { email });

    if (res.success) {
        msg.className = 'login-msg success';
        msg.innerHTML = '<i class="fas fa-check-circle"></i> ' + escapeHTML(res.message);
    } else {
        msg.className = 'login-msg error';
        msg.innerHTML = '<i class="fas fa-exclamation-circle"></i> ' + escapeHTML(res.message || 'エラーが発生しました');
        btn.disabled = false;
        btn.textContent = 'ログインリンクを送信';
    }
}

async function portalLogout() {
    if (PORTAL.sessionId) {
        await portalFetch('portal-logout');
    }
    clearSession();
    window.location.hash = '#/login';
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// Dashboard
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

async function renderDashboard(container) {
    container.innerHTML = '<div class="portal-container"><div class="loading" style="padding:3rem 0;text-align:center"><i class="fas fa-spinner fa-spin"></i> ダッシュボードを読み込み中...</div></div>';

    const res = await portalFetch('portal-dashboard');
    if (!res.success) {
        container.innerHTML = '<div class="portal-container"><p class="empty-state"><i class="fas fa-exclamation-circle"></i>データの取得に失敗しました</p></div>';
        return;
    }

    const d = res.data;
    let html = '<div class="portal-container">';
    html += '<p class="dash-welcome">おかえりなさい、<strong>' + escapeHTML(d.member.name) + '</strong> さん</p>';

    // Upcoming consultations
    html += '<div class="portal-card"><h3><i class="fas fa-calendar-check"></i> 直近の相談予定</h3>';
    if (d.upcomingConsultations && d.upcomingConsultations.length > 0) {
        d.upcomingConsultations.forEach(c => {
            html += '<div class="dash-consultation">' +
                '<div class="dash-consultation-date">' + escapeHTML(c.date.split(' ')[0].split('/').slice(1).join('/')) + '</div>' +
                '<div class="dash-consultation-info"><strong>' + escapeHTML(c.company) + '</strong>' +
                '<small>' + escapeHTML(c.method || '') + ' | リーダー: ' + escapeHTML(c.leader || '未定') + ' | ' + escapeHTML(c.theme || '') + '</small></div></div>';
        });
    } else {
        html += '<div class="empty-state"><i class="far fa-calendar"></i>直近の相談予定はありません</div>';
    }
    html += '</div>';

    // Pending tasks
    if (d.pendingTasks && d.pendingTasks.length > 0) {
        html += '<div class="portal-card"><h3><i class="fas fa-tasks"></i> 未対応タスク</h3>';
        d.pendingTasks.forEach(t => {
            const badge = t.status === '期限超過' ? 'overdue' : '';
            html += '<div class="dash-task"><span class="dash-task-badge ' + badge + '">' + escapeHTML(t.label) + '</span>' +
                '<span>' + escapeHTML(t.company) + '</span>' +
                (t.deadline ? '<small style="color:var(--text-muted)">期限: ' + escapeHTML(t.deadline) + '</small>' : '') +
                '</div>';
        });
        html += '</div>';
    }

    // Recent news
    html += '<div class="portal-card"><h3><i class="fas fa-newspaper"></i> 最新のお知らせ</h3>';
    if (d.recentNews && d.recentNews.length > 0) {
        d.recentNews.forEach(n => {
            html += '<div class="dash-news-item"><strong>' + escapeHTML(n.title) + '</strong><small>' + escapeHTML(n.date || '') + ' | ' + escapeHTML(n.category || '') + '</small></div>';
        });
    } else {
        html += '<div class="empty-state"><i class="far fa-newspaper"></i>お知らせはありません</div>';
    }
    html += '</div>';

    html += '</div>';
    container.innerHTML = html;
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// Shifts
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

async function renderShifts(container) {
    container.innerHTML = '<div class="portal-container"><div class="loading" style="padding:3rem 0;text-align:center"><i class="fas fa-spinner fa-spin"></i> シフトを読み込み中...</div></div>';

    const monthStr = PORTAL.shiftYear + '/' + String(PORTAL.shiftMonth).padStart(2, '0');
    const res = await portalFetch('portal-shifts', { month: monthStr });

    let html = '<div class="portal-container">';
    html += '<h2 class="portal-page-title"><i class="fas fa-calendar-alt"></i> シフト管理</h2>';

    // Month navigation
    html += '<div class="shift-month-nav">' +
        '<button onclick="changeMonth(-1)"><i class="fas fa-chevron-left"></i></button>' +
        '<span>' + PORTAL.shiftYear + '年' + PORTAL.shiftMonth + '月</span>' +
        '<button onclick="changeMonth(1)"><i class="fas fa-chevron-right"></i></button></div>';

    if (!res.success || !res.shifts || res.shifts.length === 0) {
        html += '<div class="empty-state"><i class="far fa-calendar-times"></i>この月の日程はありません</div>';
    } else {
        html += '<div class="shift-list">';
        res.shifts.forEach(s => {
            const cls = s.participating ? 'participating' : (s.bookingStatus === '予約済み' ? 'booked' : '');
            const dateParts = s.date.split('/');
            html += '<div class="shift-item ' + cls + '">' +
                '<div class="shift-date"><div class="shift-date-num">' + parseInt(dateParts[2]) + '</div><div class="shift-date-day">' + escapeHTML(s.dayOfWeek) + '</div></div>' +
                '<div class="shift-info"><strong>' + escapeHTML(s.time) + ' ' + escapeHTML(s.method) + '</strong>' +
                '<div class="shift-meta">予約: ' + escapeHTML(s.bookingStatus || '未定') + '</div></div>' +
                '<div class="shift-score">配置<br><strong>' + s.score + '</strong>pt</div>' +
                '<button class="shift-toggle-btn ' + (s.participating ? 'leave' : 'join') + '" onclick="toggleShift(' + s.row + ',' + !s.participating + ',this)" ' +
                (s.bookable === '不可' ? 'disabled' : '') + '>' +
                (s.participating ? '取消' : '参加') + '</button></div>';
        });
        html += '</div>';
    }

    html += '</div>';
    container.innerHTML = html;
}

function changeMonth(delta) {
    PORTAL.shiftMonth += delta;
    if (PORTAL.shiftMonth > 12) { PORTAL.shiftMonth = 1; PORTAL.shiftYear++; }
    if (PORTAL.shiftMonth < 1) { PORTAL.shiftMonth = 12; PORTAL.shiftYear--; }
    renderShifts(document.getElementById('portalMain'));
}

async function toggleShift(row, join, el) {
    const btn = el || document.activeElement;
    btn.disabled = true;
    btn.textContent = '処理中...';

    const res = await portalFetch('portal-shift-toggle', { row: String(row), join: String(join) });

    if (res.success) {
        renderShifts(document.getElementById('portalMain'));
    } else {
        alert(res.message || 'エラーが発生しました');
        btn.disabled = false;
        btn.textContent = join ? '参加' : '取消';
    }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// Cases
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

async function renderCases(container) {
    container.innerHTML = '<div class="portal-container"><div class="loading" style="padding:3rem 0;text-align:center"><i class="fas fa-spinner fa-spin"></i> 案件を読み込み中...</div></div>';

    const res = await portalFetch('portal-cases');

    let html = '<div class="portal-container">';
    html += '<h2 class="portal-page-title"><i class="fas fa-briefcase"></i> 案件管理</h2>';

    if (!res.success || !res.cases || res.cases.length === 0) {
        html += '<div class="empty-state"><i class="fas fa-inbox"></i>案件がありません</div>';
        html += '</div>';
        container.innerHTML = html;
        return;
    }

    // Filters
    const statuses = ['all', '確定', '仮予約', '完了', 'キャンセル'];
    html += '<div class="case-filters">';
    statuses.forEach(s => {
        const label = s === 'all' ? 'すべて' : s;
        html += '<button class="case-filter-btn ' + (PORTAL.caseFilter === s ? 'active' : '') + '" onclick="filterCases(\'' + s + '\')">' + label + '</button>';
    });
    html += '</div>';

    // Case list
    const filtered = PORTAL.caseFilter === 'all' ? res.cases : res.cases.filter(c => c.status === PORTAL.caseFilter);
    html += '<div class="case-list">';
    if (filtered.length === 0) {
        html += '<div class="empty-state">該当する案件はありません</div>';
    } else {
        filtered.forEach(c => {
            const statusCls = STATUS_MAP[c.status] || 'tentative';
            html += '<div class="case-item"><div class="case-item-header"><h4>' + escapeHTML(c.company || '企業名不明') + '</h4>' +
                '<span class="case-status ' + statusCls + '">' + escapeHTML(c.status) + '</span></div>' +
                '<div class="case-meta">' +
                '<span><i class="far fa-calendar"></i> ' + escapeHTML(c.confirmedDate || '日程未定') + '</span>' +
                '<span><i class="fas fa-user"></i> ' + escapeHTML(c.leader || '未定') + '</span>' +
                '<span><i class="fas fa-tag"></i> ' + escapeHTML(c.theme || '') + '</span>' +
                '<span><i class="fas fa-desktop"></i> ' + escapeHTML(c.method || '') + '</span>' +
                (c.reportStatus ? '<span><i class="fas fa-file-alt"></i> ' + escapeHTML(c.reportStatus) + '</span>' : '') +
                '</div></div>';
        });
    }
    html += '</div>';

    html += '<p style="margin-top:1rem;font-size:0.8rem;color:var(--text-muted)">合計: ' + res.total + '件</p>';
    html += '</div>';
    container.innerHTML = html;

    // Store cases for filtering
    window._portalCases = res;
}

function filterCases(status) {
    PORTAL.caseFilter = status;
    if (window._portalCases) {
        renderCasesFromData(window._portalCases);
    } else {
        renderCases(document.getElementById('portalMain'));
    }
}

function renderCasesFromData(res) {
    // Re-render with stored data
    const container = document.getElementById('portalMain');
    let html = '<div class="portal-container">';
    html += '<h2 class="portal-page-title"><i class="fas fa-briefcase"></i> 案件管理</h2>';

    const statuses = ['all', '確定', '仮予約', '完了', 'キャンセル'];
    html += '<div class="case-filters">';
    statuses.forEach(s => {
        const label = s === 'all' ? 'すべて' : s;
        html += '<button class="case-filter-btn ' + (PORTAL.caseFilter === s ? 'active' : '') + '" onclick="filterCases(\'' + s + '\')">' + label + '</button>';
    });
    html += '</div>';

    const filtered = PORTAL.caseFilter === 'all' ? res.cases : res.cases.filter(c => c.status === PORTAL.caseFilter);
    html += '<div class="case-list">';
    if (filtered.length === 0) {
        html += '<div class="empty-state">該当する案件はありません</div>';
    } else {
        filtered.forEach(c => {
            const statusCls = STATUS_MAP[c.status] || 'tentative';
            html += '<div class="case-item"><div class="case-item-header"><h4>' + escapeHTML(c.company || '企業名不明') + '</h4>' +
                '<span class="case-status ' + statusCls + '">' + escapeHTML(c.status) + '</span></div>' +
                '<div class="case-meta">' +
                '<span><i class="far fa-calendar"></i> ' + escapeHTML(c.confirmedDate || '日程未定') + '</span>' +
                '<span><i class="fas fa-user"></i> ' + escapeHTML(c.leader || '未定') + '</span>' +
                '<span><i class="fas fa-tag"></i> ' + escapeHTML(c.theme || '') + '</span>' +
                '<span><i class="fas fa-desktop"></i> ' + escapeHTML(c.method || '') + '</span>' +
                (c.reportStatus ? '<span><i class="fas fa-file-alt"></i> ' + escapeHTML(c.reportStatus) + '</span>' : '') +
                '</div></div>';
        });
    }
    html += '</div>';
    html += '<p style="margin-top:1rem;font-size:0.8rem;color:var(--text-muted)">合計: ' + res.total + '件</p>';
    html += '</div>';
    container.innerHTML = html;
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// News
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

async function renderNews(container) {
    container.innerHTML = '<div class="portal-container"><div class="loading" style="padding:3rem 0;text-align:center"><i class="fas fa-spinner fa-spin"></i> お知らせを読み込み中...</div></div>';

    const res = await fetchAPI('news');

    let html = '<div class="portal-container">';
    html += '<h2 class="portal-page-title"><i class="fas fa-newspaper"></i> お知らせ</h2>';

    // News posting form (leader/admin only)
    if (PORTAL.session && (PORTAL.session.role === 'admin' || PORTAL.session.role === 'leader')) {
        html += '<div class="portal-card" id="newsFormCard">' +
            '<h3><i class="fas fa-pen"></i> 新規投稿</h3>' +
            '<form onsubmit="submitNews(event)">' +
            '<div class="news-form-group"><label>カテゴリ</label>' +
            '<select class="news-form-select" id="newsCategory"><option>お知らせ</option><option>活動報告</option><option>イベント</option><option>重要</option></select></div>' +
            '<div class="news-form-group"><label>タイトル</label><input class="news-form-input" id="newsTitle" placeholder="タイトルを入力" required></div>' +
            '<div class="news-form-group"><label>本文</label><textarea class="news-form-textarea" id="newsBody" placeholder="本文を入力" required></textarea></div>' +
            '<button type="submit" class="news-submit-btn" id="newsSubmitBtn"><i class="fas fa-paper-plane" style="margin-right:0.3rem"></i>投稿</button>' +
            '</form></div>';
    }

    // News list
    html += '<div class="portal-card"><h3><i class="fas fa-list"></i> お知らせ一覧</h3>';
    if (res.success && res.news && res.news.length > 0) {
        res.news.forEach(n => {
            html += '<div class="dash-news-item"><strong>' + escapeHTML(n.title) + '</strong>' +
                '<small>' + escapeHTML(n.date || '') + ' | ' + escapeHTML(n.category || '') +
                (n.author ? ' | ' + escapeHTML(n.author) : '') + '</small>' +
                (n.body ? '<p style="font-size:0.8rem;color:var(--text-light);margin-top:0.3rem;line-height:1.6">' + escapeHTML(n.body).substring(0, 200) + '</p>' : '') +
                '</div>';
        });
    } else {
        html += '<div class="empty-state"><i class="far fa-newspaper"></i>お知らせはありません</div>';
    }
    html += '</div>';

    html += '</div>';
    container.innerHTML = html;
}

async function submitNews(e) {
    e.preventDefault();
    const btn = document.getElementById('newsSubmitBtn');
    const title = document.getElementById('newsTitle').value.trim();
    const body = document.getElementById('newsBody').value.trim();
    const category = document.getElementById('newsCategory').value;

    if (!title || !body) return;

    btn.disabled = true;
    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 投稿中...';

    const res = await portalFetch('portal-news-post', { title, body, category });

    if (res.success) {
        alert('ニュースを投稿しました');
        renderNews(document.getElementById('portalMain'));
    } else {
        alert(res.message || '投稿に失敗しました');
        btn.disabled = false;
        btn.innerHTML = '<i class="fas fa-paper-plane" style="margin-right:0.3rem"></i>投稿';
    }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// Profile
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

async function renderProfile(container) {
    container.innerHTML = '<div class="portal-container"><div class="loading" style="padding:3rem 0;text-align:center"><i class="fas fa-spinner fa-spin"></i> プロフィールを読み込み中...</div></div>';

    const res = await portalFetch('portal-profile');

    let html = '<div class="portal-container">';
    html += '<h2 class="portal-page-title"><i class="fas fa-user"></i> プロフィール</h2>';

    if (!res.success || !res.profile) {
        html += '<div class="empty-state"><i class="fas fa-exclamation-circle"></i>プロフィールを取得できませんでした</div>';
    } else {
        const p = res.profile;
        html += '<div class="portal-card"><h3><i class="fas fa-id-card"></i> 基本情報</h3>';
        const fields = [
            ['氏名', p.name],
            ['メール', p.email],
            ['期', p.term],
            ['資格', p.cert],
            ['区分', p.type],
            ['ロール', ROLE_LABELS[p.role] || p.role],
            ['電話', p.phone],
            ['得意業種', p.specialties],
            ['得意テーマ', p.themes],
            ['肩書き', p.titles]
        ];
        fields.forEach(([label, value]) => {
            if (value) {
                html += '<div class="profile-field"><div class="profile-label">' + label + '</div><div class="profile-value">' + escapeHTML(String(value)) + '</div></div>';
            }
        });
        html += '</div>';
    }

    // Logout button
    html += '<div style="margin-top:2rem;text-align:center">' +
        '<button onclick="portalLogout()" class="login-btn" style="max-width:300px;background:#dc3545">' +
        '<i class="fas fa-sign-out-alt" style="margin-right:0.5rem"></i>ログアウト</button></div>';

    html += '</div>';
    container.innerHTML = html;
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// Evaluations
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const EVAL_CATEGORY_LABELS = {
    c1: '問題把握', c2: '解決策', c3: 'コミュニケーション',
    c4: '時間管理', c5: '論理的展開', c6: '倫理・自律性'
};

const EVAL_STATUS_MAP = {
    '完了': 'complete', 'AI完了': 'ai-complete', 'AI評価中': 'running',
    '人間評価中': 'ai-complete', 'エラー': 'error', '未評価': ''
};

function getEvalScoreColor(score) {
    if (score >= 80) return '#28a745';
    if (score >= 60) return '#ffc107';
    return '#dc3545';
}

async function renderEvaluations(container) {
    container.innerHTML = '<div class="portal-container"><div class="loading" style="padding:3rem 0;text-align:center"><i class="fas fa-spinner fa-spin"></i> 評価データを読み込み中...</div></div>';

    const [statsRes, listRes] = await Promise.all([
        portalFetch('evaluation-stats'),
        portalFetch('portal-evaluations', { limit: '50' })
    ]);

    let html = '<div class="portal-container">';
    html += '<h2 class="portal-page-title"><i class="fas fa-chart-bar"></i> コンサルタント評価</h2>';

    // Summary cards
    if (statsRes.success) {
        html += '<div class="eval-summary-grid">' +
            '<div class="eval-summary-card"><div class="eval-summary-num">' + (statsRes.totalEvaluations || 0) + '</div><div class="eval-summary-label">総評価数</div></div>' +
            '<div class="eval-summary-card"><div class="eval-summary-num" style="color:' + getEvalScoreColor(statsRes.averageScore || 0) + '">' + (statsRes.averageScore || 0) + '<small>/100</small></div><div class="eval-summary-label">平均スコア</div></div>';

        if (statsRes.categoryAverages) {
            const cats = Object.keys(EVAL_CATEGORY_LABELS);
            const topCat = cats.reduce((a, b) => (statsRes.categoryAverages[a] || 0) >= (statsRes.categoryAverages[b] || 0) ? a : b);
            html += '<div class="eval-summary-card"><div class="eval-summary-num" style="font-size:1.2rem">' + EVAL_CATEGORY_LABELS[topCat] + '</div><div class="eval-summary-label">最高カテゴリ</div></div>';
        }
        html += '</div>';
    }

    // Admin: run evaluation button
    if (PORTAL.session && PORTAL.session.role === 'admin') {
        html += '<div class="portal-card"><h3><i class="fas fa-play-circle"></i> 評価を実行</h3>' +
            '<p style="font-size:0.8rem;color:var(--text-muted);margin-bottom:0.8rem">予約管理シートの行番号を指定して評価を実行します（文字起こし済みが必要）</p>' +
            '<div style="display:flex;gap:0.5rem;align-items:center">' +
            '<input type="number" id="evalRunRow" min="2" placeholder="行番号" style="padding:8px 12px;border:2px solid var(--border-light);border-radius:8px;width:120px;font-size:0.85rem">' +
            '<button class="eval-submit-btn" id="evalRunBtn" onclick="runPortalEvaluation()">実行</button>' +
            '</div><div id="evalRunMsg" style="margin-top:0.5rem;font-size:0.8rem"></div></div>';
    }

    // Filter
    html += '<div class="case-filters" style="margin-top:1rem">';
    const statuses = ['all', '完了', 'AI完了', '人間評価中', 'エラー'];
    statuses.forEach(s => {
        const label = s === 'all' ? 'すべて' : s;
        const active = (PORTAL.evalFilter || 'all') === s ? 'active' : '';
        html += '<button class="case-filter-btn ' + active + '" onclick="filterEvals(\'' + s + '\')">' + label + '</button>';
    });
    html += '</div>';

    // Evaluation list
    if (listRes.success && listRes.results && listRes.results.length > 0) {
        const filtered = (PORTAL.evalFilter && PORTAL.evalFilter !== 'all')
            ? listRes.results.filter(r => r.status === PORTAL.evalFilter)
            : listRes.results;

        if (filtered.length === 0) {
            html += '<div class="empty-state" style="margin-top:1rem">該当する評価はありません</div>';
        } else {
            filtered.forEach(r => {
                const statusCls = EVAL_STATUS_MAP[r.status] || '';
                html += '<div class="eval-card status-' + statusCls + '" onclick="window.location.hash=\'#/evaluation-detail?id=' + r.evalId + '\'">' +
                    '<div class="eval-card-header"><h4>' + escapeHTML(r.consultantName || '未設定') + '</h4>' +
                    '<span class="eval-score-badge">' + (r.totalScore || 0) + '</span></div>' +
                    '<div class="eval-card-meta">' +
                    '<span><i class="far fa-calendar"></i> ' + escapeHTML(r.evalDate || '') + '</span>' +
                    '<span><i class="fas fa-building"></i> ' + escapeHTML(r.companyName || '') + '</span>' +
                    '<span class="eval-status-badge ' + statusCls + '">' + escapeHTML(r.status || '') + '</span>' +
                    '</div></div>';
            });
        }
        html += '<p style="margin-top:1rem;font-size:0.8rem;color:var(--text-muted)">合計: ' + listRes.total + '件</p>';
    } else {
        html += '<div class="empty-state" style="margin-top:2rem"><i class="fas fa-chart-bar"></i>評価データがまだありません</div>';
    }

    html += '</div>';
    container.innerHTML = html;
    window._evalData = listRes;
}

function filterEvals(status) {
    PORTAL.evalFilter = status;
    renderEvaluations(document.getElementById('portalMain'));
}

async function runPortalEvaluation() {
    const row = document.getElementById('evalRunRow').value;
    const btn = document.getElementById('evalRunBtn');
    const msg = document.getElementById('evalRunMsg');
    if (!row) { msg.innerHTML = '<span style="color:#dc3545">行番号を入力してください</span>'; return; }

    btn.disabled = true;
    btn.textContent = '実行中...';
    msg.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 評価を実行中（数分かかる場合があります）...';

    const res = await portalFetch('portal-run-evaluation', { row });
    if (res.success) {
        msg.innerHTML = '<span style="color:#28a745"><i class="fas fa-check-circle"></i> ' + escapeHTML(res.message || '完了') + '</span>';
        setTimeout(() => renderEvaluations(document.getElementById('portalMain')), 2000);
    } else {
        msg.innerHTML = '<span style="color:#dc3545"><i class="fas fa-exclamation-circle"></i> ' + escapeHTML(res.message || res.error || 'エラー') + '</span>';
    }
    btn.disabled = false;
    btn.textContent = '実行';
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// Evaluation Detail
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const EVAL_ITEM_NAMES = {
    '1': '傾聴・受容', '2': '質問力・深掘り', '3': '本質課題の特定', '4': '情報収集の網羅性',
    '5': '解決策の具体性', '6': '複数選択肢の提示', '7': '実現可能性への配慮', '8': '専門知識の活用', '9': 'リスク説明',
    '10': 'わかりやすさ', '11': '信頼関係構築', '12': '要約・確認', '13': '言葉遣い',
    '14': '時間配分', '15': '議論の進行管理', '16': 'ネクストステップ',
    '17': '論理構成', '18': '根拠に基づく説明', '19': '構造化・可視化',
    '20': '自律性尊重', '21': '守秘義務・倫理', '22': '継続的改善'
};

const ITEM_TO_CATEGORY = {
    '1': 'c1', '2': 'c1', '3': 'c1', '4': 'c1',
    '5': 'c2', '6': 'c2', '7': 'c2', '8': 'c2', '9': 'c2',
    '10': 'c3', '11': 'c3', '12': 'c3', '13': 'c3',
    '14': 'c4', '15': 'c4', '16': 'c4',
    '17': 'c5', '18': 'c5', '19': 'c5',
    '20': 'c6', '21': 'c6', '22': 'c6'
};

async function renderEvaluationDetail(container, params) {
    container.innerHTML = '<div class="portal-container"><div class="loading" style="padding:3rem 0;text-align:center"><i class="fas fa-spinner fa-spin"></i> 評価詳細を読み込み中...</div></div>';

    const evalId = params.id;
    if (!evalId) { window.location.hash = '#/evaluations'; return; }

    const res = await portalFetch('portal-evaluation-detail', { id: evalId });
    if (!res.success || !res.evaluation) {
        container.innerHTML = '<div class="portal-container"><div class="empty-state"><i class="fas fa-exclamation-circle"></i>評価データが見つかりません</div><a href="#/evaluations" style="display:block;text-align:center;margin-top:1rem;font-size:0.85rem;color:var(--kgu-blue)">一覧に戻る</a></div>';
        return;
    }

    const ev = res.evaluation;
    const scoreColor = getEvalScoreColor(ev.totalScore);

    let html = '<div class="portal-container">';
    html += '<a href="#/evaluations" style="font-size:0.8rem;color:var(--kgu-blue);text-decoration:none"><i class="fas fa-arrow-left"></i> 評価一覧に戻る</a>';

    // Header
    html += '<div class="eval-detail-header" style="margin-top:1rem">' +
        '<div class="eval-score-circle" style="border-color:' + scoreColor + '"><span style="color:' + scoreColor + '">' + ev.totalScore + '</span></div>' +
        '<div class="eval-detail-info"><h3>' + escapeHTML(ev.consultantName || '未設定') + '</h3>' +
        '<p>' + escapeHTML(ev.companyName || '') + ' | ' + escapeHTML(ev.evalDate || '') + '</p>' +
        '<p>AI: ' + ev.aiTotal + '/90 + 人間: ' + ev.humanTotal + '/10 = <strong>' + ev.totalScore + '/100</strong></p>' +
        '<span class="eval-status-badge ' + (EVAL_STATUS_MAP[ev.status] || '') + '">' + escapeHTML(ev.status || '') + '</span>' +
        '</div></div>';

    // Radar chart
    html += '<div class="portal-card"><h3><i class="fas fa-chart-radar"></i> カテゴリ別スコア</h3>' +
        '<div class="eval-radar-wrap"><canvas id="evalRadar" width="350" height="280"></canvas></div></div>';

    // 22-item breakdown
    html += '<div class="portal-card"><h3><i class="fas fa-list-ol"></i> 22項目ブレークダウン</h3>';
    const catOrder = ['c1', 'c2', 'c3', 'c4', 'c5', 'c6'];
    catOrder.forEach(cat => {
        const catLabel = EVAL_CATEGORY_LABELS[cat];
        const catScore = ev.categories[cat] || 0;
        const items = Object.keys(ITEM_TO_CATEGORY).filter(k => ITEM_TO_CATEGORY[k] === cat);

        html += '<div class="eval-category-section">' +
            '<div class="eval-category-header" onclick="this.nextElementSibling.style.display=this.nextElementSibling.style.display===\'none\'?\'\':\'none\'">' +
            '<h4>' + catLabel + '</h4><span class="eval-category-score">' + catScore.toFixed(1) + '</span></div>' +
            '<div style="display:none">';

        items.forEach(num => {
            const itemName = EVAL_ITEM_NAMES[num] || 'No.' + num;
            const score = (ev.itemScores && ev.itemScores[num]) || 3;
            const pct = (score / 5) * 100;
            const evid = (ev.evidence && ev.evidence[num]) || {};
            const fillColor = score >= 4 ? '#28a745' : (score >= 3 ? '#ffc107' : '#dc3545');

            html += '<div class="eval-item">' +
                '<div class="eval-item-header"><span class="eval-item-name">No.' + num + ' ' + itemName + '</span><span class="eval-item-score">' + score + '/5</span></div>' +
                '<div class="eval-progress"><div class="eval-progress-fill" style="width:' + pct + '%;background:' + fillColor + '"></div></div>';
            if (evid.evidence) {
                html += '<div class="eval-evidence">"' + escapeHTML(evid.evidence).substring(0, 200) + '"</div>';
            }
            html += '</div>';
        });

        html += '</div></div>';
    });
    html += '</div>';

    // NG words
    if (ev.ngWords && ev.ngWords.length > 0) {
        html += '<div class="portal-card"><h3><i class="fas fa-exclamation-triangle"></i> NG語句検出</h3><div class="eval-ng-section">';
        ev.ngWords.forEach(ng => {
            html += '<div class="eval-ng-item"><span class="eval-ng-cat ' + (ng.category || 'C') + '">' + escapeHTML(ng.category || 'C') + '</span>' +
                '<span>"' + escapeHTML(ng.text || '') + '"</span>' +
                '<span style="color:var(--text-muted);font-size:0.75rem">' + escapeHTML(ng.context || '') + '</span></div>';
        });
        html += '</div></div>';
    }

    // Human score form (leader+ and AI complete)
    if (PORTAL.session && (PORTAL.session.role === 'admin' || PORTAL.session.role === 'leader') &&
        (ev.status === 'AI完了' || ev.status === '人間評価中' || ev.status === '完了')) {
        html += '<div class="portal-card"><h3><i class="fas fa-user-edit"></i> 人間評価</h3>' +
            '<p style="font-size:0.8rem;color:var(--text-muted);margin-bottom:1rem">AI評価では測れない対面品質を評価してください</p>' +
            '<form class="eval-human-form" onsubmit="submitHumanScore(event,\'' + ev.evalId + '\')">' +
            '<div class="eval-slider-group"><label>H1: 相談時の雰囲気 <span id="h1Val">' + (ev.humanScores.h1 || 0) + '</span>/4</label>' +
            '<input type="range" id="h1" min="0" max="4" value="' + (ev.humanScores.h1 || 0) + '" oninput="document.getElementById(\'h1Val\').textContent=this.value"></div>' +
            '<div class="eval-slider-group"><label>H2: 相談者の態度変化 <span id="h2Val">' + (ev.humanScores.h2 || 0) + '</span>/3</label>' +
            '<input type="range" id="h2" min="0" max="3" value="' + (ev.humanScores.h2 || 0) + '" oninput="document.getElementById(\'h2Val\').textContent=this.value"></div>' +
            '<div class="eval-slider-group"><label>H3: 身だしなみ・マナー <span id="h3Val">' + (ev.humanScores.h3 || 0) + '</span>/3</label>' +
            '<input type="range" id="h3" min="0" max="3" value="' + (ev.humanScores.h3 || 0) + '" oninput="document.getElementById(\'h3Val\').textContent=this.value"></div>' +
            '<button type="submit" class="eval-submit-btn" id="humanScoreBtn"><i class="fas fa-save" style="margin-right:0.3rem"></i>採点を登録</button>' +
            '</form><div id="humanScoreMsg" style="margin-top:0.5rem;font-size:0.8rem"></div></div>';
    }

    // Growth chart (same consultant history)
    html += '<div class="portal-card"><h3><i class="fas fa-chart-line"></i> 成長推移</h3>' +
        '<div id="growthChartWrap"><div class="loading" style="text-align:center;padding:1rem"><i class="fas fa-spinner fa-spin"></i></div></div></div>';

    html += '</div>';
    container.innerHTML = html;

    // Draw radar chart
    drawEvalRadar(ev);

    // Load growth data
    if (ev.consultantName) {
        loadGrowthChart(ev.consultantName);
    }
}

function drawEvalRadar(ev) {
    const ctx = document.getElementById('evalRadar');
    if (!ctx) return;

    const cats = ev.categories || {};
    // Normalize category scores to a 0-5 display scale for radar
    const catMaxScaled = { c1: 16.4, c2: 20.5, c3: 16.4, c4: 12.3, c5: 12.3, c6: 12.3 };
    const data = Object.keys(EVAL_CATEGORY_LABELS).map(k => {
        const maxVal = catMaxScaled[k] || 15;
        return Math.round((cats[k] || 0) / maxVal * 5 * 10) / 10;
    });

    new Chart(ctx, {
        type: 'radar',
        data: {
            labels: Object.values(EVAL_CATEGORY_LABELS),
            datasets: [{
                data: data,
                backgroundColor: 'rgba(15,35,80,0.15)',
                borderColor: '#0F2350',
                borderWidth: 2,
                pointBackgroundColor: '#0F2350',
                pointRadius: 4
            }]
        },
        options: {
            responsive: false,
            plugins: { legend: { display: false } },
            scales: {
                r: {
                    min: 0, max: 5,
                    ticks: { stepSize: 1, font: { size: 10 } },
                    pointLabels: { font: { size: 11, family: 'Noto Sans JP' } }
                }
            }
        }
    });
}

async function loadGrowthChart(consultantName) {
    const wrap = document.getElementById('growthChartWrap');
    const res = await portalFetch('portal-evaluation-growth', { consultant: consultantName });

    if (!res.success || !res.history || res.history.length < 2) {
        wrap.innerHTML = '<div class="empty-state"><i class="fas fa-chart-line"></i>成長データが不足しています（2件以上の評価が必要）</div>';
        return;
    }

    wrap.innerHTML = '<canvas id="growthChart" width="500" height="250"></canvas>';
    const ctx = document.getElementById('growthChart');

    const labels = res.history.map(h => h.evalDate);
    const scores = res.history.map(h => h.totalScore);

    new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: '総合スコア',
                data: scores,
                borderColor: '#0F2350',
                backgroundColor: 'rgba(15,35,80,0.1)',
                borderWidth: 2,
                fill: true,
                tension: 0.3
            }]
        },
        options: {
            responsive: true,
            plugins: { legend: { display: false } },
            scales: {
                y: { min: 0, max: 100, ticks: { stepSize: 20 } }
            }
        }
    });
}

async function submitHumanScore(e, evalId) {
    e.preventDefault();
    const btn = document.getElementById('humanScoreBtn');
    const msg = document.getElementById('humanScoreMsg');
    const h1 = document.getElementById('h1').value;
    const h2 = document.getElementById('h2').value;
    const h3 = document.getElementById('h3').value;

    btn.disabled = true;
    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 登録中...';

    const res = await portalFetch('portal-evaluation-human-score', { id: evalId, h1, h2, h3 });

    if (res.success) {
        msg.innerHTML = '<span style="color:#28a745"><i class="fas fa-check-circle"></i> ' + escapeHTML(res.message || '登録完了') + '</span>';
        setTimeout(() => renderEvaluationDetail(document.getElementById('portalMain'), { id: evalId }), 1500);
    } else {
        msg.innerHTML = '<span style="color:#dc3545"><i class="fas fa-exclamation-circle"></i> ' + escapeHTML(res.message || 'エラー') + '</span>';
        btn.disabled = false;
        btn.innerHTML = '<i class="fas fa-save" style="margin-right:0.3rem"></i>採点を登録';
    }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// Init
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

document.addEventListener('DOMContentLoaded', async () => {
    // If no hash, check session and redirect appropriately
    if (!window.location.hash || window.location.hash === '#' || window.location.hash === '#/') {
        const valid = await validateCurrentSession();
        window.location.hash = valid ? '#/dashboard' : '#/login';
    } else {
        router();
    }
});
