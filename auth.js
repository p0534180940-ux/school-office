// כניסה משותפת לכל דפי המערכת + סרגל ניווט עליון
(function () {
  const CFG = window.APP_CONFIG || {};
  const LS = {
    get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
    set(k, v) { try { localStorage.setItem(k, v); } catch (e) {} },
    del(k) { try { localStorage.removeItem(k); } catch (e) {} },
  };
  const page = location.pathname.split('/').pop() || 'index.html';
  const isLogin = page === 'login.html';

  const Auth = window.Auth = {
    cfg: CFG,
    enabled: !!CFG.API_URL,
    async api(action, data) {
      const r = await fetch(CFG.API_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(Object.assign({ action }, data || {})) });
      return r.json();
    },
    token() { return LS.get('so_token'); },
    user() { try { return JSON.parse(LS.get('so_user') || 'null'); } catch (e) { return null; } },
    save(token, user) { LS.set('so_token', token); LS.set('so_user', JSON.stringify(user)); LS.set('so_checked', String(Date.now())); },
    logout() { LS.del('so_token'); LS.del('so_user'); LS.del('so_checked'); location.href = 'login.html'; },
    ai: LS.get('so_ai') === '1',
    toLogin() { location.replace('login.html?next=' + encodeURIComponent(page + location.hash)); },
  };

  // מסתירים את הדף עד שהכניסה נבדקה
  const st = document.createElement('style');
  st.textContent = `html.auth-wait body{visibility:hidden}
  #soBar{background:#0f3d73;color:#fff;font-family:Assistant,Arial,sans-serif;display:flex;align-items:center;gap:4px;padding:6px 14px;font-size:15px;flex-wrap:wrap;direction:rtl}
  #soBar a{color:#fff;text-decoration:none;padding:4px 10px;border-radius:7px;opacity:.85;font-weight:600}
  #soBar a:hover{background:rgba(255,255,255,.12);opacity:1}
  #soBar a.on{background:rgba(255,255,255,.2);opacity:1}
  #soBar .who{margin-right:auto;font-size:14px;opacity:.9;display:flex;gap:8px;align-items:center}
  #soBar .who button{background:rgba(255,255,255,.15);border:0;color:#fff;border-radius:7px;padding:4px 10px;font:inherit;cursor:pointer}
  #soBar .days{background:#fdf3e2;color:#8a4f06;border-radius:20px;padding:1px 9px;font-size:13px;font-weight:700}
  @media print{#soBar{display:none!important}}`;
  document.head.appendChild(st);

  function bar(user) {
    if (isLogin || document.getElementById('soBar')) return;
    const links = [['index.html', '🏠 ראשי'], ['register.html', '📝 רישום תלמידים'], ['labels.html', '🏷️ מדבקות'], ['tasks-setup.html', '📋 משימות ותזכורות'], ['spec.html', '🧭 אפיון']];
    const b = document.createElement('div'); b.id = 'soBar';
    b.innerHTML = links.map(([h, t]) => `<a href="${h}" class="${h === page ? 'on' : ''}">${t}</a>`).join('') +
      (user ? `<span class="who">${user.daysLeft != null && user.daysLeft <= 14 ? `<span class="days">נותרו ${user.daysLeft} ימים</span>` : ''}שלום, ${String(user.name || user.email).replace(/</g, '')}<button id="soOut">יציאה</button></span>` : '');
    document.body.prepend(b);
    const o = document.getElementById('soOut'); if (o) o.onclick = () => Auth.logout();
  }
  function ready(user) {
    const go = () => { bar(user); document.documentElement.classList.remove('auth-wait'); document.dispatchEvent(new CustomEvent('auth-ready', { detail: user })); };
    if (document.body) go(); else document.addEventListener('DOMContentLoaded', go);
  }

  Auth.ready = (async () => {
    if (!Auth.enabled || isLogin) { ready(null); return null; }       // מצב הקמה – בלי מסך כניסה
    document.documentElement.classList.add('auth-wait');
    const t = Auth.token();
    if (!t) { Auth.toLogin(); return null; }
    const cached = Auth.user(), last = +(LS.get('so_checked') || 0);
    if (cached && Date.now() - last < 6 * 3600000) { ready(cached); return cached; }   // נבדק לאחרונה
    try {
      const r = await Auth.api('me', { token: t });
      if (r.ok) { Auth.save(t, r.user); LS.set('so_ai', r.ai ? '1' : '0'); Auth.ai = !!r.ai; ready(r.user); return r.user; }
      LS.del('so_token'); LS.del('so_user'); Auth.toLogin(); return null;
    } catch (e) {
      if (cached) { ready(cached); return cached; }                  // אין אינטרנט – ממשיכים עם הכניסה האחרונה
      Auth.toLogin(); return null;
    }
  })();
})();
