/**
 * KG中小企業経営診断研究会 PRサイト — 共通JS
 */

// GAS API URL
const API_URL = 'https://script.google.com/macros/s/AKfycbzR7l1lyRF9dNZ0qqIov8LZwxDvkkyT4NNo2LSJKbQR_i46iqLfSRg4EuqQRflP76elAg/exec';

// LP URL（相談申込フォームへの導線）
const LP_URL = 'https://hirokazusugisugi-hue.github.io/kg-consultation/index14.html';

/**
 * GAS APIからデータを取得
 * @param {string} action - APIアクション名
 * @param {Object} params - 追加パラメータ
 * @returns {Promise<Object>}
 */
async function fetchAPI(action, params = {}) {
    const url = new URL(API_URL);
    url.searchParams.set('action', action);
    Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, v));

    try {
        const response = await fetch(url.toString());
        return await response.json();
    } catch (error) {
        console.error('API Error:', error);
        return { success: false, error: error.message };
    }
}

/**
 * Header scroll effect
 */
function initHeader() {
    const header = document.getElementById('header');
    if (!header) return;

    window.addEventListener('scroll', () => {
        header.classList.toggle('scrolled', window.scrollY > 80);
    });
}

/**
 * Mobile navigation toggle
 */
function initMobileNav() {
    const toggle = document.querySelector('.nav-toggle');
    const navList = document.querySelector('.nav-list');
    if (!toggle || !navList) return;

    toggle.addEventListener('click', () => {
        navList.classList.toggle('open');
        const isOpen = navList.classList.contains('open');
        toggle.innerHTML = isOpen ? '<i class="fas fa-times"></i>' : '<i class="fas fa-bars"></i>';
    });

    // Close on link click
    navList.querySelectorAll('a').forEach(a => {
        a.addEventListener('click', () => {
            navList.classList.remove('open');
            toggle.innerHTML = '<i class="fas fa-bars"></i>';
        });
    });
}

/**
 * Scroll reveal animations
 */
function initScrollReveal() {
    const els = document.querySelectorAll('.reveal, .reveal-left, .reveal-right');
    if (els.length === 0) return;

    const observer = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                entry.target.classList.add('visible');
                observer.unobserve(entry.target);
            }
        });
    }, { threshold: 0.1 });

    els.forEach(el => observer.observe(el));
}

/**
 * Active navigation link
 */
function initActiveNav() {
    const currentPath = window.location.pathname.split('/').pop() || 'index.html';
    document.querySelectorAll('.nav-list a').forEach(a => {
        const href = a.getAttribute('href');
        if (href && (href === currentPath || (currentPath === 'index.html' && href === './'))) {
            a.classList.add('active');
        }
    });
}

/**
 * Date formatting
 * @param {string} dateStr - Date string
 * @returns {string} Formatted date
 */
function formatDate(dateStr) {
    if (!dateStr) return '';
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr;
    return `${d.getFullYear()}.${String(d.getMonth() + 1).padStart(2, '0')}.${String(d.getDate()).padStart(2, '0')}`;
}

/**
 * HTML escape
 */
function escapeHTML(str) {
    if (!str) return '';
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

/**
 * Truncate text
 */
function truncate(str, len) {
    if (!str) return '';
    return str.length > len ? str.substring(0, len) + '...' : str;
}

// Initialize on DOM ready
document.addEventListener('DOMContentLoaded', () => {
    initHeader();
    initMobileNav();
    initScrollReveal();
    initActiveNav();
});
