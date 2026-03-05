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
