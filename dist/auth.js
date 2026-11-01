/* Maestro authentication. Public configuration only; project data is never stored here. */
(function (global) {
  'use strict';

  const listeners = new Set();
  const DEFAULTS = {appSelector: '#app', gateSelector: '#auth-gate', render: true};
  const REFRESH_MARGIN = 60;
  const NETWORK_TIMEOUT = 15000;
  let options = {...DEFAULTS};
  let config = null;
  let session = null;
  let refreshPromise = null;
  let refreshTimer = null;
  let started = false;
  let startPromise = null;
  let authEpoch = 0;
  let storageBlocked = false;
  let accountArea = null;

  class AuthError extends Error {
    constructor(message, code = 'AUTH_ERROR', status = 0) {
      super(message);
      this.name = 'MaestroAuthError';
      this.code = code;
      this.status = status;
    }
  }

  const copy = value => value ? JSON.parse(JSON.stringify(value)) : null;
  const nowSeconds = () => Math.floor(Date.now() / 1000);

  function decodeClaims(token) {
    try {
      const segment = String(token).split('.')[1];
      if (!segment) return {};
      const padded = segment.replace(/-/g, '+').replace(/_/g, '/').padEnd(Math.ceil(segment.length / 4) * 4, '=');
      return JSON.parse(global.atob(padded));
    } catch (_) { return {}; }
  }

  function readConfig() {
    const raw = global.MAESTRO_CONFIG || {};
    const rawUrl = String(raw.supabaseUrl || '').trim();
    const key = String(raw.publishableKey || '').trim();
    if (!rawUrl || !key || /YOUR_|REPLACE|PASTE|PROJECT_REF|example/i.test(rawUrl + ' ' + key)) {
      throw new AuthError('الربط السحابي غير مكتمل. يحتاج مسؤول المشروع إلى ضبط عنوان Supabase والمفتاح العام في إعدادات الموقع.', 'CONFIG_REQUIRED');
    }
    let url;
    try { url = new URL(rawUrl); } catch (_) {
      throw new AuthError('عنوان Supabase في إعدادات الموقع غير صالح.', 'CONFIG_REQUIRED');
    }
    const local = ['localhost', '127.0.0.1', '[::1]'].includes(url.hostname);
    if ((url.protocol !== 'https:' && !(local && url.protocol === 'http:')) || url.username || url.password || url.search || url.hash) {
      throw new AuthError('يجب أن يستخدم اتصال Supabase عنوان HTTPS صالحًا دون بيانات دخول ضمن الرابط.', 'CONFIG_REQUIRED');
    }
    if (key.startsWith('sb_secret_') || decodeClaims(key).role === 'service_role') {
      throw new AuthError('إعداد المفتاح غير آمن. استخدم المفتاح العام publishable أو anon فقط؛ لا تضع مفتاح service_role أو secret في الموقع.', 'CONFIG_REQUIRED');
    }
    const base = url.toString().replace(/\/+$/, '');
    return {url: base, key, storageKey: 'maestro-auth:' + url.host + url.pathname.replace(/\/+$/, '')};
  }

  function configured() {
    if (!config) config = readConfig();
    return config;
  }

  function normalizeSession(value) {
    if (!value || typeof value.access_token !== 'string' || !value.access_token ||
        typeof value.refresh_token !== 'string' || !value.refresh_token) {
      throw new AuthError('لم تصل جلسة دخول صالحة من الخادم. أعد المحاولة أو تواصل مع مسؤول المشروع.', 'INVALID_SESSION');
    }
    let expiresAt = Number(value.expires_at);
    if (!Number.isFinite(expiresAt) || expiresAt <= 0) {
      const lifetime = Number(value.expires_in);
      expiresAt = Number.isFinite(lifetime) && lifetime > 0 ? nowSeconds() + lifetime : Number(decodeClaims(value.access_token).exp);
    }
    if (!Number.isFinite(expiresAt) || expiresAt <= 0) {
      throw new AuthError('تعذر تحديد صلاحية جلسة الدخول. أعد تسجيل الدخول.', 'INVALID_SESSION');
    }
    return {
      access_token: value.access_token,
      refresh_token: value.refresh_token,
      expires_at: Math.floor(expiresAt),
      user: {id: String(value.user?.id || ''), email: String(value.user?.email || '')}
    };
  }

  function readStored() {
    try {
      const raw = global.localStorage.getItem(configured().storageKey);
      if (!raw) return null;
      return normalizeSession(JSON.parse(raw));
    } catch (error) {
      if (error instanceof AuthError && error.code === 'CONFIG_REQUIRED') throw error;
      return null;
    }
  }

  function writeStored(value) {
    try {
      const key = configured().storageKey;
      if (value) global.localStorage.setItem(key, JSON.stringify(value));
      else global.localStorage.removeItem(key);
      storageBlocked = false;
    } catch (_) {
      // Private browsing or browser policy can prevent persistent login.
      storageBlocked = true;
    }
  }

  function emit(event) {
    for (const listener of listeners) {
      try { listener(event, copy(session)); } catch (_) { /* A view callback must not corrupt authentication. */ }
    }
  }

  function clearRefreshTimer() {
    if (refreshTimer) global.clearTimeout(refreshTimer);
    refreshTimer = null;
  }

  function scheduleRefresh(delay) {
    clearRefreshTimer();
    if (!session) return;
    const wait = delay ?? Math.max(5000, (session.expires_at - nowSeconds() - REFRESH_MARGIN) * 1000);
    refreshTimer = global.setTimeout(async () => {
      try { await requestToken(); }
      catch (error) {
        if (session && error.code === 'NETWORK_ERROR') {
          showGate(error.message, 'error', true);
          scheduleRefresh(30000);
          emit('AUTH_ERROR');
        }
      }
    }, Math.min(wait, 2147483000));
  }

  function setSession(value, event, persist = true) {
    const next = value ? normalizeSession(value) : null;
    // A different account must clear the previous account's rendered data,
    // including when a storage event arrives from another browser tab.
    if (session && next && session.user.id !== next.user.id) {
      ++authEpoch;
      session = null;
      clearRefreshTimer();
      emit('SIGNED_OUT');
      event = 'SIGNED_IN';
    }
    if (event === 'SIGNED_IN') ++authEpoch;
    if (event === 'SIGNED_IN' || event === 'SIGNED_OUT') refreshPromise = null;
    session = next;
    if (persist) writeStored(session);
    scheduleRefresh();
    updateAccount();
    if (session) showApp();
    if (event) emit(event);
  }

  function messageFor(code, status) {
    if (status === 429 || ['over_request_rate_limit', 'over_email_send_rate_limit'].includes(code)) {
      return 'محاولات الدخول متقاربة. انتظر قليلًا ثم حاول مجددًا.';
    }
    if (['invalid_credentials', 'invalid_grant'].includes(code)) return 'البريد الإلكتروني أو كلمة المرور غير صحيحة.';
    if (code === 'email_not_confirmed') return 'لم يُؤكَّد البريد الإلكتروني بعد. أكمل تأكيد الحساب أو تواصل مع مسؤول المشروع.';
    if (['user_banned', 'user_not_found'].includes(code)) return 'هذا الحساب غير متاح للدخول. تواصل مع مسؤول المشروع.';
    if (['captcha_failed', 'captcha_verification_failed'].includes(code)) return 'يتطلب تسجيل الدخول تحققًا أمنيًا غير مهيأ في هذه الواجهة. تواصل مع مسؤول المشروع.';
    if (['insufficient_aal', 'mfa_challenge_expired', 'mfa_verification_failed'].includes(code)) return 'يتطلب الحساب تحققًا إضافيًا. تواصل مع مسؤول المشروع لإكمال طريقة الدخول المناسبة.';
    if (status === 401 || status === 403) return 'تعذر اعتماد جلسة الدخول. أعد تسجيل الدخول أو تواصل مع مسؤول المشروع.';
    return 'تعذر تسجيل الدخول حاليًا. أعد المحاولة أو تواصل مع مسؤول المشروع.';
  }

  async function authFetch(path, body, token = null) {
    const current = configured();
    const controller = new AbortController();
    const timeout = global.setTimeout(() => controller.abort(), NETWORK_TIMEOUT);
    try {
      const headers = {'apikey': current.key, 'Accept': 'application/json'};
      if (body !== undefined) headers['Content-Type'] = 'application/json';
      if (token) headers.Authorization = 'Bearer ' + token;
      const response = await global.fetch(current.url + '/auth/v1/' + path, {
        method: body === undefined ? 'GET' : 'POST',
        headers,
        body: body === undefined ? undefined : JSON.stringify(body),
        cache: 'no-store',
        credentials: 'omit',
        redirect: 'error',
        referrerPolicy: 'no-referrer',
        signal: controller.signal
      });
      const data = response.status === 204 ? {} : await response.json().catch(() => ({}));
      if (!response.ok) {
        const code = String(data.error_code || data.code || data.error || 'AUTH_ERROR');
        throw new AuthError(messageFor(code, response.status), code, response.status);
      }
      return data;
    } catch (error) {
      if (error instanceof AuthError) throw error;
      throw new AuthError('تعذر الاتصال بخدمة تسجيل الدخول. تحقق من اتصالك بالإنترنت ثم أعد المحاولة.', 'NETWORK_ERROR');
    } finally { global.clearTimeout(timeout); }
  }

  function appNode() {
    return global.document?.querySelector(options.appSelector) || global.document?.querySelector('.app-shell');
  }

  function gateNode() {
    return global.document?.querySelector(options.gateSelector);
  }

  function showGate(message = '', kind = 'info', canRetry = false) {
    if (!options.render || !global.document) return;
    ensureUI();
    const app = appNode();
    if (app) { app.hidden = true; app.setAttribute('inert', ''); }
    // The original local application keeps its editor dialog outside .app-shell.
    global.document.querySelectorAll('dialog[open]').forEach(dialog => { if (dialog.close) dialog.close(); });
    const gate = gateNode();
    if (gate) gate.hidden = false;
    const notice = global.document.getElementById('auth-message');
    if (notice) {
      notice.textContent = message;
      notice.className = 'auth-message' + (kind === 'error' ? ' error' : '');
      notice.hidden = !message;
    }
    const retry = global.document.getElementById('auth-retry');
    if (retry) retry.hidden = !canRetry;
    let valid = true;
    try { configured(); } catch (_) { valid = false; }
    const submit = global.document.getElementById('auth-submit');
    if (submit) submit.disabled = !valid;
  }

  function showApp() {
    if (!options.render || !global.document) return;
    const gate = gateNode();
    if (gate) gate.hidden = true;
    const app = appNode();
    if (app) { app.hidden = false; app.removeAttribute('inert'); }
  }

  function busy(value, label = 'جارٍ التحقق…') {
    if (!options.render || !global.document) return;
    const form = global.document.getElementById('auth-form');
    if (!form) return;
    form.setAttribute('aria-busy', value ? 'true' : 'false');
    for (const field of form.querySelectorAll('input,button')) field.disabled = value;
    const button = global.document.getElementById('auth-submit');
    if (button) {
      button.textContent = value ? label : 'تسجيل الدخول';
      if (!value) { try { configured(); } catch (_) { button.disabled = true; } }
    }
  }

  function ensureUI() {
    if (!options.render || !global.document) return;
    const document = global.document;
    let gate = gateNode();
    if (!gate) {
      gate = document.createElement('section');
      if (options.gateSelector.startsWith('#')) gate.id = options.gateSelector.slice(1);
      else gate.id = 'auth-gate';
      document.body.prepend(gate);
    }
    gate.classList.add('auth-gate');
    gate.setAttribute('aria-labelledby', 'auth-title');
    if (gate.dataset.authReady === 'true') return;
    gate.dataset.authReady = 'true';
    gate.innerHTML = `
      <div class="auth-frame">
        <div class="auth-brand"><img src="./assets/logo.png" alt="MAESTRO مايسترو"><span>إدارة محتوى شاشات المدينة</span></div>
        <section class="auth-card">
          <div class="auth-heading"><span class="auth-eyebrow">مساحة عمل الفريق</span><h1 id="auth-title">تسجيل الدخول</h1><p>استخدم الحساب الذي جهّزه مسؤول المشروع للوصول إلى البيانات المشتركة.</p></div>
          <div id="auth-message" class="auth-message" role="alert" hidden></div>
          <form id="auth-form">
            <label class="auth-field" for="auth-email"><span>البريد الإلكتروني</span><input id="auth-email" name="email" type="email" autocomplete="username" inputmode="email" dir="ltr" required></label>
            <label class="auth-field" for="auth-password"><span>كلمة المرور</span><span class="auth-password-wrap"><input id="auth-password" name="password" type="password" autocomplete="current-password" dir="ltr" required><button id="auth-reveal" type="button" aria-controls="auth-password" aria-pressed="false">إظهار</button></span></label>
            <button id="auth-submit" class="btn primary auth-submit" type="submit">تسجيل الدخول</button>
          </form>
          <button id="auth-retry" class="btn secondary auth-retry" type="button" hidden>إعادة التحقق من الجلسة</button>
          <p class="auth-support">إذا لم يكن لديك حساب، تواصل مع مسؤول المشروع. لا يمكن إنشاء حساب عام من هذه الصفحة.</p>
        </section>
        <p class="auth-footnote">لا تُخزّن كلمة المرور في هذا الموقع. يُحفظ رمز الجلسة فقط في هذا المتصفح لاستمرار تسجيل الدخول.</p>
      </div>`;
    document.getElementById('auth-form').addEventListener('submit', async event => {
      event.preventDefault();
      const email = document.getElementById('auth-email').value.trim();
      const password = document.getElementById('auth-password').value;
      busy(true);
      try { await signIn(email, password); }
      catch (error) { if (error.code !== 'AUTH_CANCELLED') showGate(error.message, 'error'); }
      finally { busy(false); }
    });
    document.getElementById('auth-reveal').addEventListener('click', event => {
      const input = document.getElementById('auth-password');
      const reveal = input.type === 'password';
      input.type = reveal ? 'text' : 'password';
      event.currentTarget.textContent = reveal ? 'إخفاء' : 'إظهار';
      event.currentTarget.setAttribute('aria-pressed', reveal ? 'true' : 'false');
    });
    document.getElementById('auth-retry').addEventListener('click', async () => {
      busy(true);
      try { await restoreSession(); }
      catch (error) { showGate(error.message, 'error', error.code === 'NETWORK_ERROR'); }
      finally { busy(false); }
    });
  }

  function updateAccount() {
    if (!options.render || !global.document) return;
    const document = global.document;
    if (!accountArea || !accountArea.isConnected) {
      const host = appNode()?.querySelector('.topbar-actions');
      if (!host) return;
      accountArea = document.createElement('div');
      accountArea.className = 'auth-account';
      const email = document.createElement('span');
      email.className = 'auth-account-email';
      email.dir = 'ltr';
      const logout = document.createElement('button');
      logout.className = 'btn secondary auth-logout';
      logout.type = 'button';
      logout.textContent = 'تسجيل الخروج';
      logout.addEventListener('click', async () => {
        if (options.beforeSignOut && !(await options.beforeSignOut())) return;
        logout.disabled = true;
        try { await signOut(); } finally { logout.disabled = false; }
      });
      accountArea.append(email, logout);
      host.append(accountArea);
    }
    accountArea.hidden = !session;
    if (session) {
      accountArea.querySelector('.auth-account-email').textContent = session.user.email || 'حساب مسجّل';
      accountArea.title = storageBlocked ? 'تعذر حفظ الجلسة في المتصفح. ستحتاج إلى تسجيل الدخول بعد إعادة فتح الصفحة.' : '';
    }
  }

  async function signIn(email, password) {
    configured();
    if (!String(email || '').trim() || !String(password || '')) throw new AuthError('أدخل البريد الإلكتروني وكلمة المرور.', 'INPUT_REQUIRED');
    const epoch = ++authEpoch;
    const data = await authFetch('token?grant_type=password', {email: String(email).trim(), password: String(password)});
    if (epoch !== authEpoch) throw new AuthError('أُلغيت محاولة الدخول.', 'AUTH_CANCELLED');
    const next = normalizeSession(data);
    setSession(next, 'SIGNED_IN');
    if (options.render && global.document) {
      const field = global.document.getElementById('auth-password');
      if (field) field.value = '';
    }
    return copy(session);
  }

  async function invalidate(message = 'انتهت جلسة الدخول. سجّل الدخول مرة أخرى.') {
    ++authEpoch;
    setSession(null, 'SIGNED_OUT');
    showGate(message, 'error');
  }

  async function refresh(forceRefresh = false) {
    const epoch = authEpoch;
    const previousToken = session?.access_token;
    const work = async () => {
      const stored = readStored();
      if (stored && session && stored.user.id !== session.user.id) {
        setSession(stored, 'SIGNED_IN', false);
        throw new AuthError('تغير حساب الدخول. أعد المحاولة من مساحة العمل الحالية.', 'AUTH_CANCELLED');
      }
      if (stored) session = stored;
      else if (!storageBlocked && !session) throw new AuthError('سجّل الدخول للوصول إلى بيانات المشروع.', 'SESSION_REQUIRED');
      if (!session) throw new AuthError('سجّل الدخول للوصول إلى بيانات المشروع.', 'SESSION_REQUIRED');
      const fresh = session.expires_at > nowSeconds() + REFRESH_MARGIN;
      if (fresh && (!forceRefresh || (previousToken && session.access_token !== previousToken))) {
        scheduleRefresh();
        return session.access_token;
      }
      try {
        const data = await authFetch('token?grant_type=refresh_token', {refresh_token: session.refresh_token});
        if (epoch !== authEpoch) throw new AuthError('انتهت الجلسة الحالية.', 'AUTH_CANCELLED');
        setSession(data, 'TOKEN_REFRESHED');
        return session.access_token;
      } catch (error) {
        if (epoch !== authEpoch) throw new AuthError('انتهت الجلسة الحالية.', 'AUTH_CANCELLED');
        if (error.code === 'NETWORK_ERROR' || error.code === 'AUTH_CANCELLED') throw error;
        const newer = readStored();
        if (newer && newer.access_token !== previousToken && newer.expires_at > nowSeconds() + 5) {
          setSession(newer, 'TOKEN_REFRESHED', false);
          return newer.access_token;
        }
        await invalidate();
        throw new AuthError('انتهت جلسة الدخول. سجّل الدخول مرة أخرى.', 'SESSION_EXPIRED', error.status);
      }
    };
    if (global.navigator?.locks?.request) return global.navigator.locks.request(configured().storageKey + ':refresh', work);
    return work();
  }

  async function requestToken(requestOptions = {}) {
    configured();
    if (!session) session = readStored();
    if (!session) {
      showGate('سجّل الدخول للوصول إلى بيانات المشروع.');
      throw new AuthError('سجّل الدخول للوصول إلى بيانات المشروع.', 'SESSION_REQUIRED');
    }
    const forceRefresh = requestOptions?.forceRefresh === true;
    if (!forceRefresh && session.expires_at > nowSeconds() + REFRESH_MARGIN) return session.access_token;
    if (!refreshPromise) {
      const pending = refresh(forceRefresh).finally(() => { if (refreshPromise === pending) refreshPromise = null; });
      refreshPromise = pending;
    }
    return refreshPromise;
  }

  async function restoreSession() {
    configured();
    if (!session) session = readStored();
    if (!session) {
      showGate();
      emit('INITIAL_SESSION');
      return null;
    }
    const epoch = authEpoch;
    try {
      const token = await requestToken();
      let user;
      try { user = await authFetch('user', undefined, token); }
      catch (error) {
        if (error.status !== 401) throw error;
        const renewed = await requestToken({forceRefresh: true});
        user = await authFetch('user', undefined, renewed);
      }
      if (epoch !== authEpoch) return null;
      if (!user?.id) throw new AuthError('تعذر التحقق من هوية الحساب.', 'INVALID_SESSION');
      setSession({...session, user}, 'INITIAL_SESSION');
      return copy(session);
    } catch (error) {
      if (epoch !== authEpoch) return null;
      if (error.code === 'NETWORK_ERROR') {
        showGate(error.message, 'error', true);
        scheduleRefresh(30000);
        return null;
      }
      if (error.code !== 'AUTH_CANCELLED') await invalidate();
      return null;
    }
  }

  async function signOut() {
    const previous = session;
    const epoch = ++authEpoch;
    clearRefreshTimer();
    setSession(null, 'SIGNED_OUT');
    if (global.document) {
      const email = global.document.getElementById('auth-email');
      if (email) email.value = '';
      const password = global.document.getElementById('auth-password');
      if (password) { password.value = ''; password.type = 'password'; }
      const reveal = global.document.getElementById('auth-reveal');
      if (reveal) { reveal.textContent = 'إظهار'; reveal.setAttribute('aria-pressed', 'false'); }
      const fileInput = global.document.getElementById('importFile');
      if (fileInput) fileInput.value = '';
    }
    showGate('تم تسجيل خروجك من هذا المتصفح.');
    if (previous?.access_token) {
      try { await authFetch('logout?scope=local', {}, previous.access_token); }
      catch (_) { if (epoch === authEpoch && !session) showGate('أُغلقت الجلسة في هذا المتصفح. تعذر تأكيد إنهائها لدى الخادم؛ تحقق من الاتصال قبل استخدام جهاز مشترك.'); }
    }
  }

  function subscribe(listener) {
    if (typeof listener !== 'function') throw new TypeError('Authentication subscriber must be a function.');
    listeners.add(listener);
    return () => listeners.delete(listener);
  }

  async function start(startOptions = {}) {
    options = {...options, ...startOptions};
    if (startPromise) return startPromise;
    if (started) { updateAccount(); return restoreSession(); }
    startPromise = (async () => {
      ensureUI();
      showGate('جارٍ التحقق من جلسة الدخول…');
      try { configured(); }
      catch (error) { showGate(error.message, 'error'); return null; }
      started = true;
      session = readStored();
      if (options.render && global.document && session?.user?.email) global.document.getElementById('auth-email').value = session.user.email;
      return restoreSession();
    })().finally(() => { startPromise = null; });
    return startPromise;
  }

  if (global.addEventListener) {
    global.addEventListener('storage', event => {
      if (!config || event.key !== config.storageKey) return;
      const next = readStored();
      if (!next) {
        ++authEpoch;
        setSession(null, 'SIGNED_OUT', false);
        showGate('تم تسجيل الخروج من نافذة أخرى. سجّل الدخول مجددًا للمتابعة.');
      } else if (!session || next.access_token !== session.access_token) {
        setSession(next, session ? 'TOKEN_REFRESHED' : 'SIGNED_IN', false);
      }
    });
    global.addEventListener('focus', () => {
      if (session && session.expires_at <= nowSeconds() + REFRESH_MARGIN) requestToken().catch(() => {});
    });
  }

  global.MaestroAuth = Object.freeze({
    getSession: () => copy(session),
    requestToken,
    signIn,
    signOut,
    subscribe,
    start
  });
})(window);
