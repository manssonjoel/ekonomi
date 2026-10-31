/* Inloggning (Microsoft, OAuth2 + PKCE) och lagring i OneDrive-appmappen.
   Inga externa bibliotek. Demoläge sparar i webbläsaren. */
(function (root) {
  'use strict';
  const CFG = root.EKONOMI_CONFIG || {};
  const AUTH = 'https://login.microsoftonline.com/consumers/oauth2/v2.0';
  const GRAPH = 'https://graph.microsoft.com/v1.0/me/drive/special/approot';
  const SCOPE = 'openid profile offline_access https://graph.microsoft.com/Files.ReadWrite.AppFolder';
  const DATAFILE = 'ekonomi-data.json', AVZFILE = 'ekonomi-avanza.json';
  const LS = { tok: 'ek.tokens', pkce: 'ek.pkce', cache: 'ek.cache', demo: 'ek.demo', silentTried: 'ek.silent' };

  const redirectUri = () => location.origin + location.pathname.replace(/index\.html$/, '');
  const b64url = buf => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  const rand = n => b64url(crypto.getRandomValues(new Uint8Array(n)));
  const getJ = k => { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } };
  const setJ = (k, v) => { try { v === null ? localStorage.removeItem(k) : localStorage.setItem(k, JSON.stringify(v)); } catch (e) { } };

  const isDemo = () => !CFG.clientId || new URLSearchParams(location.search).has('demo');

  // ---------------------------------------------------------------- auth
  async function login(silent) {
    const verifier = rand(48);
    const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
    const state = rand(16);
    setJ(LS.pkce, { verifier, state, silent: !!silent, back: location.hash });
    const p = new URLSearchParams({ client_id: CFG.clientId, response_type: 'code', redirect_uri: redirectUri(),
      scope: SCOPE, code_challenge: challenge, code_challenge_method: 'S256', state, response_mode: 'query' });
    if (silent) p.set('prompt', 'none');
    location.assign(`${AUTH}/authorize?${p}`);
    return new Promise(() => { }); // sidan lämnas
  }

  async function tokenRequest(params) {
    const r = await fetch(`${AUTH}/token`, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({ client_id: CFG.clientId, scope: SCOPE, ...params }) });
    const j = await r.json().catch(() => ({}));
    if (!r.ok) throw new Error(j.error_description || j.error || ('HTTP ' + r.status));
    const now = Date.now();
    const old = getJ(LS.tok) || {};
    const tok = { access: j.access_token, exp: now + (j.expires_in || 3600) * 1000, refresh: j.refresh_token || old.refresh,
      idClaims: j.id_token ? parseJwt(j.id_token) : old.idClaims };
    setJ(LS.tok, tok);
    return tok;
  }
  function parseJwt(t) { try { return JSON.parse(decodeURIComponent(escape(atob(t.split('.')[1].replace(/-/g, '+').replace(/_/g, '/'))))); } catch (e) { return null; } }

  // Körs vid sidladdning: tar emot svaret från Microsoft efter inloggning.
  async function handleRedirect() {
    const q = new URLSearchParams(location.search);
    if (!q.has('code') && !q.has('error')) return null;
    const pk = getJ(LS.pkce); setJ(LS.pkce, null);
    const back = (pk && pk.back) || '';
    history.replaceState(null, '', redirectUri() + back);
    if (q.has('error')) {
      if (pk && pk.silent) return { needLogin: true };
      throw new Error(q.get('error_description') || q.get('error'));
    }
    if (!pk || pk.state !== q.get('state')) throw new Error('Inloggningen kunde inte verifieras (state). Försök igen.');
    await tokenRequest({ grant_type: 'authorization_code', code: q.get('code'), redirect_uri: redirectUri(), code_verifier: pk.verifier });
    sessionStorage.removeItem(LS.silentTried);
    return { ok: true };
  }

  async function token() {
    const t = getJ(LS.tok);
    if (t && t.access && t.exp - 60000 > Date.now()) return t.access;
    if (t && t.refresh) {
      try { return (await tokenRequest({ grant_type: 'refresh_token', refresh_token: t.refresh })).access; } catch (e) { /* utgången */ }
    }
    // tyst omdirigering (Microsoft minns oftast inloggningen) – max en gång per session
    if (t && !sessionStorage.getItem(LS.silentTried)) { sessionStorage.setItem(LS.silentTried, '1'); await login(true); }
    const err = new Error('Du behöver logga in.'); err.needLogin = true; throw err;
  }
  const signedIn = () => !!getJ(LS.tok);
  const account = () => { const t = getJ(LS.tok); const c = t && t.idClaims; return c ? (c.preferred_username || c.email || c.name) : null; };
  function logout() { setJ(LS.tok, null); setJ(LS.cache, null); }

  // ---------------------------------------------------------------- OneDrive
  async function g(path, opt = {}) {
    const tk = await token();
    const r = await fetch(GRAPH + path, { ...opt, headers: { Authorization: 'Bearer ' + tk, ...(opt.headers || {}) } });
    return r;
  }
  async function loadRemote(file = DATAFILE) {
    const meta = await g(`:/${file}`);
    if (meta.status === 404) return { data: null, etag: null };
    if (!meta.ok) throw new Error('Kunde inte läsa från OneDrive (' + meta.status + ')');
    const m = await meta.json();
    let r = m['@microsoft.graph.downloadUrl'] ? await fetch(m['@microsoft.graph.downloadUrl']).catch(() => null) : null;
    if (!r || !r.ok) r = await g(`:/${file}:/content`);
    if (!r.ok) throw new Error('Kunde inte hämta datafilen (' + r.status + ')');
    const data = await r.json();
    return { data, etag: m.eTag, modified: m.lastModifiedDateTime };
  }
  async function saveRemote(data, etag, file = DATAFILE) {
    const h = { 'Content-Type': 'application/json' };
    if (etag) h['If-Match'] = etag;
    const r = await g(`:/${file}:/content`, { method: 'PUT', headers: h, body: JSON.stringify(data) });
    if (r.status === 412) { const e = new Error('Datan har ändrats från en annan enhet. Den har laddats om – gör om ändringen.'); e.conflict = true; throw e; }
    if (!r.ok) throw new Error('Kunde inte spara till OneDrive (' + r.status + ')');
    const j = await r.json();
    return j.eTag;
  }
  async function putFile(name, blob) {
    const r = await g(`:/${encodeURIComponent(name)}:/content`, { method: 'PUT', headers: { 'Content-Type': blob.type || 'application/octet-stream' }, body: blob });
    if (!r.ok) throw new Error('Kunde inte spara ' + name + ' (' + r.status + ')');
  }

  // ---------------------------------------------------------------- publikt gränssnitt
  const Store = {
    isDemo, handleRedirect, login: () => login(false), logout, signedIn, account,
    async load() {
      if (isDemo()) return { data: getJ(LS.demo), etag: 'demo' };
      const res = await loadRemote();
      if (res.data) setJ(LS.cache, { data: res.data, etag: res.etag, at: Date.now() });
      return res;
    },
    cached() { return isDemo() ? null : getJ(LS.cache); },
    async save(data, etag) {
      if (isDemo()) { setJ(LS.demo, data); return 'demo'; }
      const e = await saveRemote(data, etag);
      setJ(LS.cache, { data, etag: e, at: Date.now() });
      return e;
    },
    async saveFile(name, blob) { if (isDemo()) return false; await putFile(name, blob); return true; },
    // Avanza-transaktioner och innehav i en egen fil (större, ändras sällan)
    async loadAvanza() {
      if (isDemo()) return { data: getJ(LS.demo + '.avanza'), etag: 'demo' };
      return loadRemote(AVZFILE);
    },
    async saveAvanza(data, etag) {
      if (isDemo()) { setJ(LS.demo + '.avanza', data); return 'demo'; }
      return saveRemote(data, etag, AVZFILE);
    }
  };
  root.Store = Store;
})(this);
