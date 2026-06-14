/* =====================================================================
   auth.js — Client-side auth: JWT storage, login/logout UI
   ===================================================================== */

(function () {
  window.AE = window.AE || {};
  const TOKEN_KEY = 'ae_token';

  // ── Token helpers ──────────────────────────────────────────────────
  function getToken() {
    return localStorage.getItem(TOKEN_KEY);
  }

  function setToken(token) {
    localStorage.setItem(TOKEN_KEY, token);
  }

  function clearToken() {
    localStorage.removeItem(TOKEN_KEY);
  }

  // ── API fetch with auth header ──────────────────────────────────────
  async function apiFetch(url, options = {}) {
    const token = getToken();
    const headers = {
      'Content-Type': 'application/json',
      ...(options.headers || {})
    };
    if (token) headers['Authorization'] = `Bearer ${token}`;
    if (options.body instanceof FormData) {
      delete headers['Content-Type']; // Let browser set multipart boundary
    }
    const res = await fetch(url, { ...options, headers });
    return res;
  }

  // ── Current user ────────────────────────────────────────────────────
  let currentUser = null;

  async function loadCurrentUser() {
    const token = getToken();
    if (!token) return null;
    try {
      const res = await apiFetch('/api/auth/me');
      if (res.ok) {
        const data = await res.json();
        currentUser = data.user;
        return currentUser;
      }
    } catch (e) { /* silent */ }
    clearToken();
    currentUser = null;
    return null;
  }

  function getCurrentUser() {
    return currentUser;
  }

  function getInitials(name) {
    if (!name) return '?';
    return name.split(' ').slice(0, 2).map(w => w[0]).join('').toUpperCase();
  }

  // ── Update account button ────────────────────────────────────────────
  function updateAccountBtn(user) {
    const btn = document.getElementById('account-btn');
    if (!btn) return;
    if (user) {
      btn.textContent = getInitials(user.name);
      btn.classList.add('logged-in');
      btn.title = user.name;
    } else {
      btn.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>`;
      btn.classList.remove('logged-in');
      btn.title = 'Sign in';
    }
  }

  // ── Account Panel ────────────────────────────────────────────────────
  function renderAccountPanel(user) {
    const panel = document.getElementById('account-panel');
    if (!panel) return;

    if (user) {
      panel.innerHTML = `
        <div class="account-panel-header">
          <h3>Account</h3>
          <p>Signed in to AuditEase</p>
        </div>
        <div class="account-panel-body">
          <div class="account-panel-user">
            <div class="account-panel-avatar">${getInitials(user.name)}</div>
            <div class="account-panel-user-info">
              <h4>${escapeHtml(user.name)}</h4>
              <p>@${escapeHtml(user.username)}</p>
            </div>
          </div>
          <button class="btn btn-ghost w-full" id="signout-btn" style="justify-content:center;">
            <svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
            Sign Out
          </button>
        </div>
      `;
      document.getElementById('signout-btn').addEventListener('click', handleSignOut);
    } else {
      panel.innerHTML = `
        <div class="account-panel-header">
          <h3>Sign In</h3>
          <p>Access AuditEase</p>
        </div>
        <div class="account-panel-body">
          <div class="form-group">
            <label for="login-username">Username</label>
            <input type="text" id="login-username" placeholder="Enter your username" autocomplete="username" />
          </div>
          <div class="form-group">
            <label for="login-password">Password</label>
            <input type="password" id="login-password" placeholder="Enter your password" autocomplete="current-password" />
          </div>
          <p class="login-error" id="login-error"></p>
          <button class="btn btn-primary" id="login-btn" style="width:100%;justify-content:center;">
            <svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"/><polyline points="10 17 15 12 10 7"/><line x1="15" y1="12" x2="3" y2="12"/></svg>
            Sign In
          </button>
        </div>
      `;
      const loginBtn = document.getElementById('login-btn');
      const pwInput = document.getElementById('login-password');
      loginBtn.addEventListener('click', handleLogin);
      pwInput.addEventListener('keydown', e => { if (e.key === 'Enter') handleLogin(); });
      document.getElementById('login-username').addEventListener('keydown', e => { if (e.key === 'Enter') pwInput.focus(); });
    }
  }

  async function handleLogin() {
    const username = document.getElementById('login-username').value.trim();
    const password = document.getElementById('login-password').value;
    const errEl = document.getElementById('login-error');
    const btn = document.getElementById('login-btn');

    if (!username || !password) {
      errEl.textContent = 'Please enter username and password';
      return;
    }

    btn.disabled = true;
    btn.innerHTML = '<span class="spinner"></span> Signing in…';
    errEl.textContent = '';

    try {
      const res = await fetch('/api/auth/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username, password })
      });
      const data = await res.json();
      if (res.ok) {
        setToken(data.token);
        currentUser = data.user;
        updateAccountBtn(currentUser);
        closeAccountPanel();
        renderAccountPanel(currentUser);
        // Refresh page data
        if (typeof window.onAuthChange === 'function') window.onAuthChange(currentUser);
      } else {
        errEl.textContent = data.error || 'Login failed';
        btn.disabled = false;
        btn.textContent = 'Sign In';
      }
    } catch (e) {
      errEl.textContent = 'Connection error. Please try again.';
      btn.disabled = false;
      btn.textContent = 'Sign In';
    }
  }

  async function handleSignOut() {
    try {
      await apiFetch('/api/auth/logout', { method: 'POST' });
    } catch (e) { /* silent */ }
    clearToken();
    currentUser = null;
    updateAccountBtn(null);
    closeAccountPanel();
    renderAccountPanel(null);
    if (typeof window.onAuthChange === 'function') window.onAuthChange(null);
  }

  // ── Panel visibility ─────────────────────────────────────────────────
  function openAccountPanel() {
    const panel = document.getElementById('account-panel');
    if (panel) panel.classList.remove('hidden');
  }

  function closeAccountPanel() {
    const panel = document.getElementById('account-panel');
    if (panel) panel.classList.add('hidden');
  }

  function toggleAccountPanel() {
    const panel = document.getElementById('account-panel');
    if (panel) panel.classList.toggle('hidden');
  }

  // ── Escape helper ────────────────────────────────────────────────────
  function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }

  // ── Auth guard (for protected pages) ────────────────────────────────
  function showAuthGuard() {
    const guard = document.createElement('div');
    guard.className = 'auth-guard';
    guard.innerHTML = `
      <div class="auth-guard-icon">
        <svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="var(--text-muted)" stroke-width="1.5">
          <rect x="3" y="11" width="18" height="11" rx="2" ry="2"/>
          <path d="M7 11V7a5 5 0 0 1 10 0v4"/>
        </svg>
      </div>
      <h2>Authentication Required</h2>
      <p>Please sign in to access this page.</p>
      <button class="btn btn-primary" onclick="document.getElementById('account-btn').click()">Sign In</button>
    `;
    document.body.appendChild(guard);
  }

  // ── initAuthUI — called by topbar.js AFTER topbar HTML is injected ──────
  // This must NOT run in DOMContentLoaded because #account-btn and
  // #account-panel don't exist until initTopbar() injects the topbar HTML.
  async function initAuthUI() {
    const accountBtn = document.getElementById('account-btn');
    if (accountBtn) {
      accountBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        toggleAccountPanel();
      });
    } else {
      console.error('[AuditEase] #account-btn not found — topbar may not have been injected yet');
    }

    // Close panel when clicking outside
    document.addEventListener('click', (e) => {
      const panel = document.getElementById('account-panel');
      const btn = document.getElementById('account-btn');
      if (panel && !panel.classList.contains('hidden')) {
        if (!panel.contains(e.target) && e.target !== btn) {
          closeAccountPanel();
        }
      }
    });

    const user = await loadCurrentUser();
    updateAccountBtn(user);
    renderAccountPanel(user);

    // Show auth guard on protected pages
    const isProtected = document.body.dataset.protected === 'true';
    if (isProtected && !user) {
      showAuthGuard();
    }
  }

  // ── Exports ───────────────────────────────────────────────────────────
  window.AE.getToken = getToken;
  window.AE.apiFetch = apiFetch;
  window.AE.getCurrentUser = getCurrentUser;
  window.AE.loadCurrentUser = loadCurrentUser;
  window.AE.escapeHtml = escapeHtml;
  window.AE.updateAccountBtn = updateAccountBtn;
  window.AE.renderAccountPanel = renderAccountPanel;
  window.AE.initAuthUI = initAuthUI;  // Called by topbar.js after HTML injection
})();
