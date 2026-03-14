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

const ROLE_LABELS = { admin: '管理者', member: 'メンバー', observer: 'オブザーバー' };
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
        msg.style.display = 'block';
        btn.textContent = '送信しました';
        // 5秒後にボタンをリセット（再送信可能に）
        setTimeout(function() {
            btn.disabled = false;
            btn.textContent = 'ログインリンクを送信';
        }, 5000);
    } else {
        msg.className = 'login-msg error';
        msg.innerHTML = '<i class="fas fa-exclamation-circle"></i> ' + escapeHTML(res.message || 'エラーが発生しました');
        btn.disabled = false;
        btn.textContent = 'ログインリンクを送信';
    }
    msg.style.display = 'block';
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
                '<small>' + escapeHTML(c.method || '') + ' | 担当: ' + escapeHTML(c.leader || '未定') + ' | ' + escapeHTML(c.theme || '') + '</small></div></div>';
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

    // Build calendar
    const year = PORTAL.shiftYear;
    const month = PORTAL.shiftMonth;
    const firstDay = new Date(year, month - 1, 1).getDay(); // 0=Sun
    const daysInMonth = new Date(year, month, 0).getDate();
    const DOW = ['日', '月', '火', '水', '木', '金', '土'];

    // Map shifts by day number
    const shiftsByDay = {};
    if (res.success && res.shifts) {
        res.shifts.forEach(s => {
            const day = parseInt(s.date.split('/')[2]);
            if (!shiftsByDay[day]) shiftsByDay[day] = [];
            shiftsByDay[day].push(s);
        });
    }

    html += '<div class="cal-grid">';
    // Day-of-week header
    DOW.forEach((d, i) => {
        const cls = i === 0 ? ' cal-sun' : (i === 6 ? ' cal-sat' : '');
        html += '<div class="cal-dow' + cls + '">' + d + '</div>';
    });
    // Blank cells before 1st
    for (let i = 0; i < firstDay; i++) html += '<div class="cal-cell cal-empty"></div>';
    // Day cells
    for (let d = 1; d <= daysInMonth; d++) {
        const dow = (firstDay + d - 1) % 7;
        const isToday = (d === new Date().getDate() && month === new Date().getMonth() + 1 && year === new Date().getFullYear());
        let cls = 'cal-cell';
        if (dow === 0) cls += ' cal-sun';
        if (dow === 6) cls += ' cal-sat';
        if (isToday) cls += ' cal-today';

        const shifts = shiftsByDay[d] || [];
        const hasShift = shifts.length > 0;
        const activeShifts = shifts.filter(s => s.bookingStatus !== 'クローズ');
        const participating = shifts.some(s => s.participating);
        if (participating) cls += ' cal-participating';
        else if (activeShifts.length > 0) cls += ' cal-has-shift';

        html += '<div class="' + cls + '">';
        html += '<div class="cal-day-num">' + d + '</div>';
        if (hasShift) {
            shifts.forEach(s => {
                const booked = s.bookingStatus === '予約済み';
                const closed = s.bookingStatus === 'クローズ';
                // メソッド表示名変換
                let methodLabel = s.method;
                if (methodLabel === '両方') methodLabel = 'Zoom＋対面';
                else if (methodLabel === 'オンライン') methodLabel = 'Zoom';

                html += '<div class="cal-shift' + (s.participating ? ' mine' : '') + (booked ? ' booked' : '') + (closed ? ' closed' : '') + '">';
                html += '<span class="cal-shift-time">' + escapeHTML(s.time) + '</span>';
                html += '<span class="cal-shift-method">' + escapeHTML(methodLabel) + '</span>';
                if (closed) {
                    html += '<span class="cal-shift-badge closed">クローズ</span>';
                } else if (booked) {
                    html += '<span class="cal-shift-badge confirmed">確定</span>';
                }
                if (!closed) {
                    if (s.participating) {
                        html += '<span class="cal-shift-status-text">参加予定 <button class="cal-cancel-btn" onclick="event.stopPropagation();cancelShift(' + s.row + ',this)" title="キャンセル">&times;</button></span>';
                    } else if (s.bookable !== '不可' && !booked) {
                        html += '<button class="cal-join-btn" onclick="event.stopPropagation();toggleShift(' + s.row + ',true,this)">参加を申請</button>';
                    }
                }
                html += '</div>';
            });
        }
        html += '</div>';
    }
    html += '</div>';

    // Legend
    html += '<div class="cal-legend">' +
        '<span><span class="cal-legend-dot confirmed"></span>確定（予約済み）</span>' +
        '<span><span class="cal-legend-dot participating"></span>参加予定</span>' +
        '<span><span class="cal-legend-dot has-shift"></span>日程あり</span>' +
        '<span><span class="cal-legend-dot closed"></span>クローズ</span>' +
        '</div>';

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

    const res = await portalFetch('portal-shift-toggle', { row: String(row), join: join ? 'true' : 'false' });

    if (res.success) {
        renderShifts(document.getElementById('portalMain'));
    } else {
        alert(res.message || 'エラーが発生しました');
        btn.disabled = false;
        btn.textContent = join ? '参加を申請' : '×';
    }
}

async function cancelShift(row, el) {
    const ok = confirm('⚠ 警告\n\n本当にキャンセルしますか？\n\nキャンセルすることで他の方に迷惑がかかる可能性があるため、十分にご注意ください。');
    if (!ok) return;
    toggleShift(row, false, el);
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
            const isOnline = c.method && (c.method.indexOf('オンライン') >= 0 || c.method.indexOf('Zoom') >= 0 || c.method.indexOf('zoom') >= 0);
            const needsVenue = !isOnline && !c.location && c.status !== '確定' && c.status !== '完了' && c.status !== 'キャンセル';
            html += '<div class="case-item"><div class="case-item-header"><h4>' + escapeHTML(c.company || '企業名不明') + '</h4>' +
                '<span class="case-status ' + statusCls + '">' + escapeHTML(c.status) + '</span></div>' +
                '<div class="case-meta">' +
                '<span><i class="far fa-calendar"></i> ' + escapeHTML(c.confirmedDate || '日程未定') + '</span>' +
                '<span><i class="fas fa-user"></i> ' + escapeHTML(c.leader || '未定') + '</span>' +
                '<span><i class="fas fa-tag"></i> ' + escapeHTML(c.theme || '') + '</span>' +
                '<span><i class="fas fa-desktop"></i> ' + escapeHTML(c.method || '') + '</span>' +
                (c.location ? '<span><i class="fas fa-map-marker-alt"></i> ' + escapeHTML(c.location) + '</span>' : '') +
                (c.reportStatus ? '<span><i class="fas fa-file-alt"></i> ' + escapeHTML(c.reportStatus) + '</span>' : '') +
                '</div>' +
                (needsVenue ? '<button class="venue-unset-btn" onclick="openVenueDialog(' + c.row + ',\'' + escapeHTML(c.company || '').replace(/'/g, "\\'") + '\')"><i class="fas fa-map-marker-alt"></i> 会場未設定</button>' : '') +
                buildAudioSection(c) +
                '</div>';
        });
    }
    html += '</div>';

    html += '<p style="margin-top:1rem;font-size:0.8rem;color:var(--text-muted)">合計: ' + res.total + '件</p>';
    html += '</div>';
    container.innerHTML = html;

    // Store cases for filtering & update audio token
    window._portalCases = res;
    if (res.audioApiToken) AUDIO_CONFIG.apiToken = res.audioApiToken;

    // 処理中の文字起こしがあればポーリング開始
    startTranscriptPolling();
}

function filterCases(status) {
    PORTAL.caseFilter = status;
    if (window._portalCases) {
        renderCasesFromData(window._portalCases);
    } else {
        renderCases(document.getElementById('portalMain'));
    }
}

const VENUE_OPTIONS = ['アプローズタワー 14階', 'アプローズタワー 10階', 'ナレッジサロン', '貸会議室'];

function isApproseVenue(v) {
    return v.indexOf('アプローズタワー') === 0;
}

function openVenueDialog(row, company) {
    let html = '<div class="venue-overlay" id="venueOverlay" onclick="if(event.target===this)closeVenueDialog()">' +
        '<div class="venue-dialog">' +
        '<h3><i class="fas fa-map-marker-alt"></i> 会場を選択</h3>' +
        '<p class="venue-company">' + escapeHTML(company) + '</p>' +
        '<div class="venue-options">';
    VENUE_OPTIONS.forEach((v, i) => {
        html += '<label class="venue-option"><input type="radio" name="venueChoice" value="' + escapeHTML(v) + '"' + (i === 0 ? ' checked' : '') + ' onchange="toggleRoomInput()"> ' + escapeHTML(v) + '</label>';
    });
    html += '</div>' +
        '<div class="venue-room-input" id="venueRoomInput">' +
        '<label for="roomNumber">部屋番号</label>' +
        '<input type="text" id="roomNumber" placeholder="例: 1401">' +
        '</div>' +
        '<div class="venue-actions">' +
        '<button class="venue-cancel-btn" onclick="closeVenueDialog()">キャンセル</button>' +
        '<button class="venue-confirm-btn" id="venueConfirmBtn" onclick="confirmVenue(' + row + ')">確定する</button>' +
        '</div></div></div>';
    document.body.insertAdjacentHTML('beforeend', html);
    toggleRoomInput();
}

function toggleRoomInput() {
    var selected = document.querySelector('input[name="venueChoice"]:checked');
    var roomDiv = document.getElementById('venueRoomInput');
    if (selected && isApproseVenue(selected.value)) {
        roomDiv.style.display = 'block';
    } else {
        roomDiv.style.display = 'none';
    }
}

function closeVenueDialog() {
    const el = document.getElementById('venueOverlay');
    if (el) el.remove();
}

async function confirmVenue(row) {
    const selected = document.querySelector('input[name="venueChoice"]:checked');
    if (!selected) return;
    var venue = selected.value;
    if (isApproseVenue(venue)) {
        var room = (document.getElementById('roomNumber').value || '').trim();
        if (room) {
            venue += ' ' + room + '号室';
        }
    }
    const btn = document.getElementById('venueConfirmBtn');
    btn.disabled = true;
    btn.textContent = '処理中...';
    const res = await portalFetch('portal-set-venue', { row: String(row), venue: venue });
    closeVenueDialog();
    if (res.success) {
        alert(res.message);
        renderCases(document.getElementById('portalMain'));
    } else {
        alert(res.message || 'エラーが発生しました');
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
            const isOnline = c.method && (c.method.indexOf('オンライン') >= 0 || c.method.indexOf('Zoom') >= 0 || c.method.indexOf('zoom') >= 0);
            const needsVenue = !isOnline && !c.location && c.status !== '確定' && c.status !== '完了' && c.status !== 'キャンセル';
            html += '<div class="case-item"><div class="case-item-header"><h4>' + escapeHTML(c.company || '企業名不明') + '</h4>' +
                '<span class="case-status ' + statusCls + '">' + escapeHTML(c.status) + '</span></div>' +
                '<div class="case-meta">' +
                '<span><i class="far fa-calendar"></i> ' + escapeHTML(c.confirmedDate || '日程未定') + '</span>' +
                '<span><i class="fas fa-user"></i> ' + escapeHTML(c.leader || '未定') + '</span>' +
                '<span><i class="fas fa-tag"></i> ' + escapeHTML(c.theme || '') + '</span>' +
                '<span><i class="fas fa-desktop"></i> ' + escapeHTML(c.method || '') + '</span>' +
                (c.location ? '<span><i class="fas fa-map-marker-alt"></i> ' + escapeHTML(c.location) + '</span>' : '') +
                (c.reportStatus ? '<span><i class="fas fa-file-alt"></i> ' + escapeHTML(c.reportStatus) + '</span>' : '') +
                '</div>' +
                (needsVenue ? '<button class="venue-unset-btn" onclick="openVenueDialog(' + c.row + ',\'' + escapeHTML(c.company || '').replace(/'/g, "\\'") + '\')"><i class="fas fa-map-marker-alt"></i> 会場未設定</button>' : '') +
                buildAudioSection(c) +
                '</div>';
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
    if (PORTAL.session && (PORTAL.session.role === 'admin' || PORTAL.session.role === 'member')) {
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

        // Profile change request
        html += '<div class="portal-card"><h3><i class="fas fa-edit"></i> プロフィール変更依頼</h3>' +
            '<p style="font-size:0.8rem;color:var(--text-muted);margin-bottom:1rem">変更したい内容を記入して送信してください。管理者が確認後、更新いたします。</p>' +
            '<form onsubmit="submitProfileChange(event)">' +
            '<textarea class="news-form-textarea" id="profileChangeDetail" placeholder="例：電話番号を 090-xxxx-xxxx に変更したい\n得意テーマに「IT戦略」を追加したい" required style="min-height:100px"></textarea>' +
            '<button type="submit" class="news-submit-btn" id="profileChangeBtn" style="margin-top:0.8rem">' +
            '<i class="fas fa-paper-plane" style="margin-right:0.3rem"></i>変更を依頼する</button>' +
            '</form><div id="profileChangeMsg" style="margin-top:0.5rem;font-size:0.8rem"></div></div>';
    }

    // Logout button
    html += '<div style="margin-top:2rem;text-align:center">' +
        '<button onclick="portalLogout()" class="login-btn" style="max-width:300px;background:#dc3545">' +
        '<i class="fas fa-sign-out-alt" style="margin-right:0.5rem"></i>ログアウト</button></div>';

    html += '</div>';
    container.innerHTML = html;
}

async function submitProfileChange(e) {
    e.preventDefault();
    const btn = document.getElementById('profileChangeBtn');
    const msg = document.getElementById('profileChangeMsg');
    const detail = document.getElementById('profileChangeDetail').value.trim();
    if (!detail) return;

    btn.disabled = true;
    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 送信中...';

    const res = await portalFetch('portal-profile-change', { detail });

    if (res.success) {
        msg.innerHTML = '<span style="color:#155724;background:#d4edda;padding:6px 12px;border-radius:6px;display:inline-block"><i class="fas fa-check-circle"></i> ' + escapeHTML(res.message) + '</span>';
        document.getElementById('profileChangeDetail').value = '';
    } else {
        msg.innerHTML = '<span style="color:#721c24;background:#f8d7da;padding:6px 12px;border-radius:6px;display:inline-block"><i class="fas fa-exclamation-circle"></i> ' + escapeHTML(res.message || 'エラーが発生しました') + '</span>';
    }
    btn.disabled = false;
    btn.innerHTML = '<i class="fas fa-paper-plane" style="margin-right:0.3rem"></i>変更を依頼する';
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// Audio Upload & Transcription
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const AUDIO_CONFIG = {
    uploadUrl: 'https://iba-consulting.jp/site/api/audio/upload.php',
    apiToken: '',  // portal.htmlまたはconfig経由で設定
    maxSize: 100 * 1024 * 1024,
    allowedExts: ['mp3', 'm4a', 'wav', 'ogg', 'webm'],
    chunkSize: 2 * 1024 * 1024  // 2MB per chunk
};

/**
 * 案件カードの音声セクションHTML生成
 */
function buildAudioSection(c) {
    // 確定 or 完了ステータスのみ表示
    if (c.status !== '確定' && c.status !== '完了') return '';

    let html = '<div class="audio-section">';

    if (c.audioUrl) {
        // 音声あり: 再生ボタン + 文字起こしUI
        html += '<div class="audio-controls">';
        html += '<button class="audio-play-btn" onclick="playAudio(\'' + escapeHTML(c.audioUrl) + '\', this)"><i class="fas fa-play"></i> 再生</button>';
        html += '<button class="audio-delete-btn" onclick="deleteAudio(' + c.row + ', this)" title="音声を削除"><i class="fas fa-trash-alt"></i></button>';
        html += '</div>';

        // 文字起こし状態
        if (c.transcriptStatus === '完了' && c.transcriptFileId) {
            html += '<div class="audio-transcript-status">';
            html += '<span class="audio-status-badge completed"><i class="fas fa-check-circle"></i> 文字起こし完了</span>';
            html += '<a href="https://docs.google.com/document/d/' + escapeHTML(c.transcriptFileId) + '/edit" target="_blank" class="audio-doc-link"><i class="fas fa-external-link-alt"></i> 議事録を開く</a>';
            html += '</div>';
        } else if (c.transcriptStatus === '処理中') {
            html += '<div class="audio-transcript-status">';
            html += '<span class="audio-status-badge processing"><i class="fas fa-spinner fa-spin"></i> 文字起こし処理中...</span>';
            html += '</div>';
        } else if (c.transcriptStatus === 'エラー') {
            html += '<div class="audio-transcript-status">';
            html += '<span class="audio-status-badge error"><i class="fas fa-exclamation-triangle"></i> 文字起こしエラー</span>';
            html += '<button class="audio-transcribe-btn" onclick="requestTranscribe(' + c.row + ', this)"><i class="fas fa-redo"></i> 再実行</button>';
            html += '</div>';
        } else {
            // 未実行
            html += '<div class="audio-transcript-status">';
            html += '<button class="audio-transcribe-btn" onclick="requestTranscribe(' + c.row + ', this)"><i class="fas fa-file-alt"></i> 文字起こし実行</button>';
            html += '</div>';
        }
    } else {
        // 音声なし: アップロードボタン
        html += '<button class="audio-upload-btn" onclick="openAudioDialog(' + c.row + ',\'' + escapeHTML(c.company || '').replace(/'/g, "\\'") + '\')"><i class="fas fa-microphone"></i> 音声をアップロード</button>';
    }

    html += '</div>';
    return html;
}

/**
 * 音声アップロードダイアログを開く
 */
function openAudioDialog(row, company) {
    let html = '<div class="venue-overlay" id="audioOverlay" onclick="if(event.target===this)closeAudioDialog()">' +
        '<div class="venue-dialog audio-dialog">' +
        '<h3><i class="fas fa-microphone"></i> 音声をアップロード</h3>' +
        '<p class="venue-company">' + escapeHTML(company) + '</p>' +
        '<div class="audio-drop-zone" id="audioDropZone">' +
        '<i class="fas fa-cloud-upload-alt" style="font-size:2rem;color:var(--text-muted);margin-bottom:0.5rem"></i>' +
        '<p>ファイルをドラッグ＆ドロップ<br>または</p>' +
        '<label class="audio-file-label">' +
        '<input type="file" id="audioFileInput" accept=".mp3,.m4a,.wav,.ogg,.webm" onchange="onAudioFileSelected(this)" style="display:none">' +
        'ファイルを選択</label>' +
        '<p class="audio-hint">対応形式: MP3, M4A, WAV, OGG, WebM（最大100MB）</p>' +
        '</div>' +
        '<div id="audioFileInfo" style="display:none">' +
        '<div class="audio-file-name" id="audioFileName"></div>' +
        '<div class="audio-progress" id="audioProgress" style="display:none"><div class="audio-progress-bar" id="audioProgressBar"></div></div>' +
        '</div>' +
        '<div class="venue-actions">' +
        '<button class="venue-cancel-btn" onclick="closeAudioDialog()">キャンセル</button>' +
        '<button class="venue-confirm-btn" id="audioUploadBtn" onclick="uploadAudio(' + row + ')" disabled>アップロード</button>' +
        '</div></div></div>';
    document.body.insertAdjacentHTML('beforeend', html);

    // ドラッグ＆ドロップ
    const dropZone = document.getElementById('audioDropZone');
    dropZone.addEventListener('dragover', function(ev) {
        ev.preventDefault();
        dropZone.classList.add('dragover');
    });
    dropZone.addEventListener('dragleave', function() {
        dropZone.classList.remove('dragover');
    });
    dropZone.addEventListener('drop', function(ev) {
        ev.preventDefault();
        dropZone.classList.remove('dragover');
        if (ev.dataTransfer.files.length > 0) {
            document.getElementById('audioFileInput').files = ev.dataTransfer.files;
            onAudioFileSelected(document.getElementById('audioFileInput'));
        }
    });
}

function onAudioFileSelected(input) {
    const file = input.files[0];
    if (!file) return;

    const ext = file.name.split('.').pop().toLowerCase();
    if (!AUDIO_CONFIG.allowedExts.includes(ext)) {
        alert('対応していないファイル形式です。\n対応形式: ' + AUDIO_CONFIG.allowedExts.join(', '));
        input.value = '';
        return;
    }
    if (file.size > AUDIO_CONFIG.maxSize) {
        alert('ファイルサイズが100MBを超えています。');
        input.value = '';
        return;
    }

    const sizeMB = (file.size / (1024 * 1024)).toFixed(1);
    document.getElementById('audioFileName').textContent = file.name + ' (' + sizeMB + ' MB)';
    document.getElementById('audioFileInfo').style.display = 'block';
    document.getElementById('audioUploadBtn').disabled = false;
}

function closeAudioDialog() {
    const el = document.getElementById('audioOverlay');
    if (el) el.remove();
}

/**
 * 音声ファイルをさくらPHPにアップロード → GASにURL保存
 */
async function uploadAudio(row) {
    const fileInput = document.getElementById('audioFileInput');
    const file = fileInput.files[0];
    if (!file) return;

    const btn = document.getElementById('audioUploadBtn');
    btn.disabled = true;
    btn.textContent = 'アップロード中...';

    const progressDiv = document.getElementById('audioProgress');
    const progressBar = document.getElementById('audioProgressBar');
    progressDiv.style.display = 'block';

    try {
        let audioUrl;
        const chunkSize = AUDIO_CONFIG.chunkSize;

        if (file.size <= chunkSize) {
            // 小さいファイル: 通常アップロード
            audioUrl = await uploadAudioNormal(file, row, progressBar);
        } else {
            // 大きいファイル: チャンクアップロード
            audioUrl = await uploadAudioChunked(file, row, progressBar, btn);
        }

        // GASにURL保存
        const res = await portalFetch('portal-save-audio-url', { row: String(row), audioUrl: audioUrl });
        if (!res.success) {
            throw new Error(res.message || 'URL保存に失敗しました');
        }

        closeAudioDialog();
        alert('音声ファイルをアップロードしました');
        renderCases(document.getElementById('portalMain'));

    } catch (err) {
        alert('エラー: ' + err.message);
        btn.disabled = false;
        btn.textContent = 'アップロード';
        progressDiv.style.display = 'none';
    }
}

/** 通常アップロード（小さいファイル） */
function uploadAudioNormal(file, row, progressBar) {
    return new Promise((resolve, reject) => {
        const formData = new FormData();
        formData.append('file', file);
        formData.append('row', String(row));

        const xhr = new XMLHttpRequest();
        xhr.open('POST', AUDIO_CONFIG.uploadUrl);
        xhr.setRequestHeader('X-Audio-Token', AUDIO_CONFIG.apiToken);

        xhr.upload.onprogress = function(e) {
            if (e.lengthComputable) {
                progressBar.style.width = Math.round((e.loaded / e.total) * 100) + '%';
            }
        };
        xhr.onload = function() {
            if (xhr.status === 200) {
                const res = JSON.parse(xhr.responseText);
                res.success ? resolve(res.url) : reject(new Error(res.message || 'アップロード失敗'));
            } else {
                reject(new Error('HTTP ' + xhr.status));
            }
        };
        xhr.onerror = function() { reject(new Error('ネットワークエラー')); };
        xhr.send(formData);
    });
}

/** チャンクアップロード（大きいファイル） */
async function uploadAudioChunked(file, row, progressBar, btn) {
    const chunkSize = AUDIO_CONFIG.chunkSize;
    const totalChunks = Math.ceil(file.size / chunkSize);
    const uploadId = Date.now() + '-' + Math.random().toString(36).substr(2, 9);

    for (let i = 0; i < totalChunks; i++) {
        const start = i * chunkSize;
        const end = Math.min(start + chunkSize, file.size);
        const chunk = file.slice(start, end);

        btn.textContent = 'アップロード中... (' + (i + 1) + '/' + totalChunks + ')';
        progressBar.style.width = Math.round(((i + 1) / totalChunks) * 100) + '%';

        const formData = new FormData();
        formData.append('chunk', chunk);
        formData.append('row', String(row));
        formData.append('chunkIndex', String(i));
        formData.append('totalChunks', String(totalChunks));
        formData.append('fileName', file.name);
        formData.append('uploadId', uploadId);

        const resp = await fetch(AUDIO_CONFIG.uploadUrl, {
            method: 'POST',
            headers: { 'X-Audio-Token': AUDIO_CONFIG.apiToken },
            body: formData
        });

        if (!resp.ok) {
            const errText = await resp.text();
            let msg = 'HTTP ' + resp.status;
            try { msg = JSON.parse(errText).message || msg; } catch(e) {}
            throw new Error(msg);
        }

        const result = await resp.json();
        if (!result.success) throw new Error(result.message || 'チャンクアップロード失敗');

        // 最後のチャンクで完了
        if (result.complete && result.url) {
            return result.url;
        }
    }
    throw new Error('アップロード完了レスポンスが取得できませんでした');
}

/**
 * 音声再生
 */
async function playAudio(url, btn) {
    // 既存の再生中プレイヤーを停止
    const existing = document.getElementById('audioPlayer');
    if (existing) {
        existing.pause();
        if (existing._blobUrl) URL.revokeObjectURL(existing._blobUrl);
        existing.remove();
        // 同じURLなら停止のみ
        if (existing.dataset.url === url) {
            btn.innerHTML = '<i class="fas fa-play"></i> 再生';
            return;
        }
    }

    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 読込中...';

    try {
        // fetch+blobでトークンをURLに露出させない
        const res = await fetch(url, { headers: { 'X-Audio-Token': AUDIO_CONFIG.apiToken } });
        if (!res.ok) throw new Error('HTTP ' + res.status);
        const blob = await res.blob();
        const blobUrl = URL.createObjectURL(blob);

        const audio = document.createElement('audio');
        audio.id = 'audioPlayer';
        audio.dataset.url = url;
        audio._blobUrl = blobUrl;
        audio.src = blobUrl;
        audio.style.display = 'none';
        document.body.appendChild(audio);

        btn.innerHTML = '<i class="fas fa-stop"></i> 停止';

        audio.play().catch(function(e) {
            alert('再生に失敗しました: ' + e.message);
            btn.innerHTML = '<i class="fas fa-play"></i> 再生';
        });

        audio.onended = function() {
            btn.innerHTML = '<i class="fas fa-play"></i> 再生';
            URL.revokeObjectURL(blobUrl);
            audio.remove();
        };
    } catch (e) {
        alert('音声の読み込みに失敗しました: ' + e.message);
        btn.innerHTML = '<i class="fas fa-play"></i> 再生';
    }
}

/**
 * 音声を削除
 */
async function deleteAudio(row, btn) {
    if (!confirm('音声ファイルの登録を解除しますか？\n（サーバー上のファイルは削除されません）')) return;

    btn.disabled = true;
    const res = await portalFetch('portal-delete-audio', { row: String(row) });
    if (res.success) {
        renderCases(document.getElementById('portalMain'));
    } else {
        alert(res.message || 'エラーが発生しました');
        btn.disabled = false;
    }
}

/**
 * 文字起こしリクエスト
 */
async function requestTranscribe(row, btn) {
    if (!confirm('文字起こしを実行しますか？\n処理には数分〜数十分かかる場合があります。')) return;

    btn.disabled = true;
    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 送信中...';

    const res = await portalFetch('portal-request-transcribe', { row: String(row) });
    if (res.success) {
        alert('文字起こし処理を開始しました。完了するまでお待ちください。');
        renderCases(document.getElementById('portalMain'));
    } else {
        alert(res.message || '文字起こしの開始に失敗しました');
        btn.disabled = false;
        btn.innerHTML = '<i class="fas fa-file-alt"></i> 文字起こし実行';
    }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// Transcript Status Polling
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

let _transcriptPollTimer = null;

/**
 * 処理中の文字起こしがあれば30秒ごとにポーリング
 */
function startTranscriptPolling() {
    stopTranscriptPolling();
    if (!window._portalCases || !window._portalCases.cases) return;

    const hasProcessing = window._portalCases.cases.some(c => c.transcriptStatus === '処理中');
    if (!hasProcessing) return;

    _transcriptPollTimer = setInterval(async () => {
        // 案件ページ表示中のみ
        const { path } = getRoute();
        if (path !== '/cases') { stopTranscriptPolling(); return; }

        const res = await portalFetch('portal-cases');
        if (!res.success) return;

        // ステータス変化チェック
        const oldCases = window._portalCases.cases || [];
        const changed = res.cases.some(c => {
            const old = oldCases.find(o => o.row === c.row);
            return old && old.transcriptStatus !== c.transcriptStatus;
        });

        if (changed) {
            window._portalCases = res;
            if (res.audioApiToken) AUDIO_CONFIG.apiToken = res.audioApiToken;
            renderCasesFromData(res);
            // まだ処理中があるか確認
            if (!res.cases.some(c => c.transcriptStatus === '処理中')) {
                stopTranscriptPolling();
            }
        }
    }, 30000);
}

function stopTranscriptPolling() {
    if (_transcriptPollTimer) {
        clearInterval(_transcriptPollTimer);
        _transcriptPollTimer = null;
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
