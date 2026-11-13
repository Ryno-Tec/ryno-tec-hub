/* Ryno-Tec Business Hub: Microsoft 365 runtime.
 *
 * The hub page was first built as a Claude artifact, which gives the page a small runtime
 * (window.claude.use('db' | 'mcp' | 'assets' | 'downloads')). This file provides the same
 * interface on Microsoft 365, so the hub code itself stays almost unchanged:
 *   db        -> SharePoint lists on the Ryno-Tec Hub site (one Hub_<section> list each)
 *   mcp       -> Outlook mail through Microsoft Graph (answers the hub's Gmail-style calls)
 *   assets    -> files in the site's Compliance library
 *   downloads -> a normal browser download
 * Sign-in is MSAL (single-page app, no secrets). Config is in hub-config.js.
 */
(function () {
  'use strict';
  const CFG = window.HUB_CONFIG;
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  const SCOPES = CFG.scopes;
  window.HUB_M365 = true;
  document.documentElement.classList.add('m365');

  /* ---------------- sign-in ---------------- */
  let account = null, resolveReady, pca = null;
  const ready = new Promise(res => { resolveReady = res; });

  function gate(msg, showButton) {
    let el = document.getElementById('m365gate');
    if (!el) {
      el = document.createElement('div');
      el.id = 'm365gate';
      el.setAttribute('role', 'dialog');
      el.setAttribute('aria-modal', 'true');
      el.setAttribute('aria-labelledby', 'm365gateTitle');
      document.body.appendChild(el);
    }
    el.innerHTML = `<div class="m365card"><h1 id="m365gateTitle">Ryno-Tec Business Hub</h1><p>${msg}</p>${showButton ? '<button type="button" class="btn primary" id="m365signin">Sign in with Microsoft</button>' : ''}</div>`;
    const b = document.getElementById('m365signin');
    if (b && pca) { b.addEventListener('click', () => pca.loginRedirect({ scopes: SCOPES, prompt: 'select_account' })); b.focus(); }
  }
  function ungate() { const el = document.getElementById('m365gate'); if (el) el.remove(); }

  const domReady = new Promise(r => document.readyState === 'loading' ? document.addEventListener('DOMContentLoaded', r, { once: true }) : r());
  (async () => {
    await domReady;
    try {
      if (!window.isSecureContext) throw new Error('Open the hub over https (hub.ryno-tec.co.za)');
      pca = new msal.PublicClientApplication({
        auth: { clientId: CFG.clientId, authority: CFG.authority, redirectUri: location.origin, postLogoutRedirectUri: location.origin, navigateToLoginRequestUrl: false },
        cache: { cacheLocation: 'localStorage' }
      });
      await pca.initialize();
      const r = await pca.handleRedirectPromise();
      account = (r && r.account) || pca.getActiveAccount() || pca.getAllAccounts()[0] || null;
      if (!account) { gate('Sign in with your Ryno-Tec Microsoft 365 account to open the hub.', true); return; }
      pca.setActiveAccount(account);
      ungate();
      resolveReady(account);
      addSignOut();
    } catch (e) {
      console.error(e);
      gate('Sign-in failed: ' + esc(e.errorMessage || e.message || String(e)) + '. If this keeps happening, ask your administrator to assign you to the Ryno-Tec Hub app.', true);
    }
  })();

  function addSignOut() {
    const foot = document.querySelector('.rail-foot') || document.getElementById('rail');
    if (!foot || document.getElementById('m365who')) return;
    const p = document.createElement('p');
    p.id = 'm365who';
    p.className = 'muted small';
    p.innerHTML = `${esc(account.username)} · <button type="button" class="link" id="m365out">Sign out</button>`;
    foot.appendChild(p);
    document.getElementById('m365out').addEventListener('click', () => pca.logoutRedirect({ account }));
  }

  async function token() {
    await ready;
    try { return (await pca.acquireTokenSilent({ scopes: SCOPES, account })).accessToken; }
    catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError) { await pca.acquireTokenRedirect({ scopes: SCOPES, account }); }
      throw e;
    }
  }
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  async function graph(method, path, body, extra) {
    const url = /^https?:/.test(path) ? path : GRAPH + path;
    const headers = { Authorization: 'Bearer ' + await token(), ...((extra && extra.headers) || {}) };
    let payload;
    if (body instanceof Blob || body instanceof ArrayBuffer) payload = body;
    else if (body !== undefined && body !== null) { headers['Content-Type'] = 'application/json'; payload = JSON.stringify(body); }
    for (let attempt = 0; attempt < 4; attempt++) {
      const r = await fetch(url, { method, headers, body: payload });
      if (r.status === 429 || r.status === 503) { await sleep((+r.headers.get('Retry-After') || 2) * 1000); continue; }
      if (!r.ok) {
        let msg = r.status + ' ' + r.statusText;
        try { const j = await r.json(); msg = (j.error && j.error.message) || msg; } catch (_) {}
        const err = new Error(msg); err.status = r.status;
        err.code = r.status === 403 ? 'invalid_argument' : r.status === 404 ? 'not_found' : 'upstream_error';
        throw err;
      }
      if (extra && extra.raw) return r;
      if (r.status === 204 || r.status === 202) return null;
      const t = await r.text();
      return t ? JSON.parse(t) : null;
    }
    const e = new Error('Microsoft 365 is busy. Try again in a moment.'); e.code = 'rate_limited'; throw e;
  }
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  window.hubGraph = graph;

  /* ---------------- db: SharePoint lists ---------------- */
  const site = CFG.siteId;
  const LISTS = CFG.lists;
  const cache = {};          // listKey -> Map(hubId -> {itemId, data})
  const listeners = {};      // listKey -> [{kind:'col'|'doc', hubId, cb}]
  const HDR = { headers: { Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' } };

  function listFor(path) {
    // 'tasks' -> tasks; 'settings/business' -> settings + business; 'meta/etender' -> settings + meta-etender
    const [col, id] = path.split('/');
    if (col === 'meta') return { key: 'settings', hubId: 'meta-' + id };
    return { key: col, hubId: id };
  }
  async function load(key) {
    const listId = LISTS[key];
    if (!listId) { cache[key] = new Map(); return cache[key]; }
    const m = new Map();
    let next = `/sites/${site}/lists/${listId}/items?$expand=fields($select=HubId,Data)&$top=500`;
    while (next) {
      const r = await graph('GET', next, null, HDR);
      for (const it of r.value || []) {
        const f = it.fields || {};
        if (!f.HubId) continue;
        let data = {};
        try { data = f.Data ? JSON.parse(f.Data) : {}; } catch (_) {}
        m.set(f.HubId, { itemId: it.id, data });
      }
      next = r['@odata.nextLink'];
    }
    cache[key] = m;
    return m;
  }
  function emit(key) {
    const m = cache[key] || new Map();
    for (const l of listeners[key] || []) {
      try {
        if (l.kind === 'col') l.cb({ docs: [...m.entries()].map(([id, v]) => ({ id, data: () => ({ ...v.data }) })) });
        else { const v = m.get(l.hubId); l.cb({ exists: !!v, data: () => (v ? { ...v.data } : undefined) }); }
      } catch (e) { console.error(e); }
    }
  }
  const loading = {};
  async function refresh(key) {
    if (loading[key]) return loading[key];
    loading[key] = load(key).then(() => emit(key)).catch(e => {
      for (const l of listeners[key] || []) if (l.err) l.err(e);
    }).finally(() => { loading[key] = null; });
    return loading[key];
  }
  function titleOf(d, id) {
    for (const k of ['subject', 'title', 'name', 'doc', 'number', 'service', 'legalName']) if (d[k]) return String(d[k]).slice(0, 250);
    return id;
  }
  const statusOf = d => String(d.status || d.stage || d.category || '');
  async function write(key, hubId, data) {
    const listId = LISTS[key];
    if (!listId) { const e = new Error('No SharePoint list for ' + key + '. Run Import-RynoTecHub.ps1 again.'); e.code = 'not_found'; throw e; }
    if (!cache[key]) await load(key);
    const fields = { Title: titleOf(data, hubId), HubId: hubId, Status: statusOf(data), Data: JSON.stringify(data) };
    if (data.updatedAt) fields.UpdatedAt = new Date(data.updatedAt).toISOString();
    const have = cache[key].get(hubId);
    if (have) await graph('PATCH', `/sites/${site}/lists/${listId}/items/${have.itemId}/fields`, fields);
    else {
      const it = await graph('POST', `/sites/${site}/lists/${listId}/items`, { fields });
      cache[key].set(hubId, { itemId: it.id, data });
    }
    cache[key].set(hubId, { itemId: (cache[key].get(hubId) || {}).itemId, data });
    emit(key);
  }
  async function remove(key, hubId) {
    if (!cache[key]) await load(key);
    const have = cache[key].get(hubId);
    if (have) await graph('DELETE', `/sites/${site}/lists/${LISTS[key]}/items/${have.itemId}`);
    cache[key].delete(hubId);
    emit(key);
  }
  const newId = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  function docRef(key, hubId) {
    return {
      id: hubId,
      set: data => write(key, hubId, data),
      update: async data => { if (!cache[key]) await load(key); const old = (cache[key].get(hubId) || {}).data || {}; return write(key, hubId, { ...old, ...data }); },
      delete: () => remove(key, hubId),
      onSnapshot(cb, err) {
        (listeners[key] = listeners[key] || []).push({ kind: 'doc', hubId, cb, err });
        cache[key] ? emit(key) : refresh(key);
        return () => { listeners[key] = listeners[key].filter(l => l.cb !== cb); };
      }
    };
  }
  const DBX = {
    collection(k) {
      return {
        doc: id => docRef(k, id || newId()),
        onSnapshot(cb, err) {
          (listeners[k] = listeners[k] || []).push({ kind: 'col', cb, err });
          cache[k] ? emit(k) : refresh(k);
          return () => { listeners[k] = listeners[k].filter(l => l.cb !== cb); };
        }
      };
    },
    doc(path) { const { key, hubId } = listFor(path); return docRef(key, hubId); }
  };
  // keep in step with edits made on other devices
  let lastPoll = Date.now();
  const pollAll = () => { lastPoll = Date.now(); Object.keys(listeners).forEach(k => { if ((listeners[k] || []).length) refresh(k); }); };
  setInterval(() => { if (!document.hidden) pollAll(); }, 120000);
  document.addEventListener('visibilitychange', () => { if (!document.hidden && Date.now() - lastPoll > 30000) pollAll(); });

  /* ---------------- mcp: Outlook mail in the shape the hub expects ---------------- */
  // Outlook conversation ids can contain '/', '+', '=' and also '-', '_'; the hub puts the id in the
  // page address (#/mail/<id>), so only the characters that would break it are swapped, reversibly.
  const b64u = s => String(s).replace(/\//g, '!').replace(/\+/g, '*').replace(/=/g, '~');
  const unb64u = s => String(s).replace(/!/g, '/').replace(/\*/g, '+').replace(/~/g, '=');
  const seen = new Map(); // thread id -> message ids from the last list, as a fallback for opening it
  const who = r => r && r.emailAddress ? (r.emailAddress.name && r.emailAddress.name !== r.emailAddress.address ? `${r.emailAddress.name} <${r.emailAddress.address}>` : r.emailAddress.address) : '';
  const addr = r => (r && r.emailAddress && r.emailAddress.address) || '';
  const SEL = 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,bodyPreview,conversationId,isRead,hasAttachments,webLink';
  const toMsg = m => ({
    id: m.id, threadId: b64u(m.conversationId), subject: m.subject || '', sender: who(m.from),
    toRecipients: (m.toRecipients || []).map(addr), ccRecipients: (m.ccRecipients || []).map(addr),
    date: m.receivedDateTime || m.sentDateTime, snippet: m.bodyPreview || '',
    labelIds: m.isRead === false ? ['UNREAD', 'INBOX'] : ['INBOX'], viewUrl: m.webLink
  });
  function parseQuery(q) {
    q = String(q || '');
    const folder = /\bin:sent\b/i.test(q) ? 'sentitems' : 'inbox';
    const addrs = [...new Set((q.match(/\b(?:to|cc|from|deliveredto):([^\s()]+@[^\s()]+)/gi) || []).map(x => x.split(':').slice(1).join(':').toLowerCase()))];
    const text = q.replace(/\([^)]*\b(to|cc|from|deliveredto):[^)]*\)/gi, ' ').replace(/\b(in|to|cc|from|deliveredto):\S+/gi, ' ').replace(/\b(OR|AND)\b/g, ' ').replace(/\s+/g, ' ').trim();
    return { folder, text, addrs };
  }
  async function searchThreads(args) {
    const { folder, text, addrs } = parseQuery(args.query);
    const top = addrs.length ? 100 : Math.min(50, Math.max(10, (args.pageSize || 30) * 2));
    let path = `/me/mailFolders/${folder}/messages?$top=${top}&$select=${SEL}`;
    path += text ? `&$search=${encodeURIComponent('"' + text.replace(/"/g, '') + '"')}` : '&$orderby=receivedDateTime desc';
    const r = await graph('GET', path, null, text ? { headers: { ConsistencyLevel: 'eventual' } } : undefined);
    const byConv = new Map();
    const mine = m => !addrs.length || [m.from, ...(m.toRecipients || []), ...(m.ccRecipients || [])].some(x => addrs.includes(addr(x).toLowerCase()));
    for (const m of (r.value || []).filter(mine)) {
      const k = m.conversationId;
      if (!byConv.has(k)) byConv.set(k, []);
      byConv.get(k).push(m);
    }
    const threads = [...byConv.entries()].map(([conv, ms]) => {
      ms.sort((a, b) => String(a.receivedDateTime).localeCompare(String(b.receivedDateTime)));
      seen.set(b64u(conv), ms.map(m => m.id));
      return { id: b64u(conv), messageCount: ms.length, messages: ms.map(toMsg), viewUrl: ms[ms.length - 1].webLink };
    }).sort((a, b) => String(b.messages[b.messages.length - 1].date).localeCompare(String(a.messages[a.messages.length - 1].date)));
    return { threads: threads.slice(0, args.pageSize || 30) };
  }
  async function getThread(args) {
    const conv = unb64u(args.threadId);
    const r = await graph('GET', `/me/messages?$filter=${encodeURIComponent(`conversationId eq '${conv.replace(/'/g, "''")}'`)}&$top=50&$select=${SEL},body`, null, { headers: { Prefer: 'outlook.body-content-type="text"' } });
    let list = r.value || [];
    if (!list.length && seen.has(args.threadId)) {
      list = (await Promise.all(seen.get(args.threadId).map(id => graph('GET', `/me/messages/${id}?$select=${SEL},body`, null, { headers: { Prefer: 'outlook.body-content-type="text"' } }).catch(() => null)))).filter(Boolean);
    }
    const ms = list.sort((a, b) => String(a.receivedDateTime).localeCompare(String(b.receivedDateTime)));
    const out = [];
    for (const m of ms) {
      const msg = { ...toMsg(m), plaintextBody: (m.body && m.body.content) || m.bodyPreview || '' };
      if (m.hasAttachments) {
        try { const a = await graph('GET', `/me/messages/${m.id}/attachments?$select=id,name,contentType,size`); msg.attachments = (a.value || []).filter(x => !x.isInline).map(x => ({ id: x.id, filename: x.name, mimeType: x.contentType, size: x.size })); } catch (_) { msg.attachments = []; }
      }
      out.push(msg);
    }
    return { id: args.threadId, messages: out, viewUrl: ms.length ? ms[ms.length - 1].webLink : '' };
  }
  const rcpts = list => (list || []).filter(Boolean).map(a => ({ emailAddress: { address: a } }));
  const mailBody = a => a.htmlBody ? { contentType: 'HTML', content: a.htmlBody } : { contentType: 'Text', content: a.body || '' };
  const mailAtts = a => (a.attachments || []).map(x => ({ '@odata.type': '#microsoft.graph.fileAttachment', name: x.filename, contentType: x.mimeType || 'application/octet-stream', contentBytes: x.content }));
  const message = a => ({ subject: a.subject || '', body: mailBody(a), toRecipients: rcpts(a.to), ccRecipients: rcpts(a.cc), attachments: mailAtts(a) });
  const replyComment = a => a.htmlBody || esc(a.body || '').replace(/\n/g, '<br>');
  async function callTool(_server, name, args) {
    args = args || {};
    let payload;
    switch (name) {
      case 'search_threads': payload = await searchThreads(args); break;
      case 'get_thread': payload = await getThread(args); break;
      case 'send_message': await graph('POST', '/me/sendMail', { message: message(args), saveToSentItems: true }); payload = { ok: true }; break;
      case 'create_draft':
        if (args.replyToMessageId) {
          const d = await graph('POST', `/me/messages/${args.replyToMessageId}/createReply`, { comment: replyComment(args) });
          payload = { id: d.id };
        } else { const d = await graph('POST', '/me/messages', message(args)); payload = { id: d.id }; }
        break;
      case 'reply':
        await graph('POST', `/me/messages/${args.messageId}/${args.replyAll ? 'replyAll' : 'reply'}`, { comment: replyComment(args) });
        payload = { ok: true }; break;
      default: { const e = new Error('Not available in the Microsoft 365 hub: ' + name); e.code = 'tool_error'; throw e; }
    }
    return { payload };
  }
  function watchTool(server, name, args, onEvent, opts) {
    let stopped = false;
    const run = () => callTool(server, name, args)
      .then(r => { if (!stopped) onEvent({ type: 'data', result: { payload: r.payload, cache: { storedAt: Date.now() } } }); })
      .catch(e => { if (!stopped) onEvent({ type: 'error', error: e }); });
    run();
    const t = setInterval(() => { if (!document.hidden) run(); }, (opts && opts.refetchInterval) || 120000);
    watchers.add(run);
    return () => { stopped = true; clearInterval(t); watchers.delete(run); };
  }
  const watchers = new Set();
  const MCPX = { callTool, watchTool, invalidate: async () => { watchers.forEach(f => f()); } };

  /* ---------------- assets: Compliance library ---------------- */
  const drive = CFG.complianceDrive;
  const ASSETSX = {
    async upload(file, opts) {
      const type = (opts && opts.type) || file.type || 'application/octet-stream';
      const safe = String(file.name || 'document').replace(/[\\/:*?"<>|#%]/g, '-');
      const path = `/drives/${drive}/root:/${encodeURIComponent('Hub uploads')}/${encodeURIComponent(safe)}:`;
      let item;
      if (file.size <= 4 * 1024 * 1024) {
        item = await graph('PUT', `${path}/content?@microsoft.graph.conflictBehavior=rename`, file);
      } else {
        const s = await graph('POST', `${path}/createUploadSession`, { item: { '@microsoft.graph.conflictBehavior': 'rename' } });
        const chunk = 5 * 320 * 1024;
        for (let start = 0; start < file.size; start += chunk) {
          const end = Math.min(start + chunk, file.size);
          const r = await fetch(s.uploadUrl, { method: 'PUT', headers: { 'Content-Range': `bytes ${start}-${end - 1}/${file.size}` }, body: file.slice(start, end) });
          if (!r.ok && r.status !== 202) { const e = new Error('Upload failed (' + r.status + ')'); e.code = 'upstream_error'; throw e; }
          if (r.status === 200 || r.status === 201) item = await r.json();
        }
      }
      return { id: item.id, sizeBytes: item.size, contentType: type, sp: true, driveId: drive, itemId: item.id, webUrl: item.webUrl };
    }
  };
  // Fetch a stored document for the viewer or a supplier pack
  window.fetchHubFile = async function (f) {
    if (f && f.sp && f.itemId) return graph('GET', `/drives/${f.driveId || drive}/items/${f.itemId}/content`, null, { raw: true });
    return new Response(null, { status: 404 });
  };

  /* ---------------- downloads ---------------- */
  const DLX = {
    async save({ filename, data }) {
      const url = URL.createObjectURL(data instanceof Blob ? data : new Blob([data]));
      const a = document.createElement('a');
      a.href = url; a.download = filename || 'download';
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 5000);
    }
  };

  /* ---------------- runtime the hub asks for ---------------- */
  window.claude = {
    async use(name) {
      await ready;
      switch (name) {
        case 'db': return DBX;
        case 'mcp': return MCPX;
        case 'assets': return ASSETSX;
        case 'downloads': return DLX;
        default: return null; // 'sample' (Claude drafting) and 'permissions' are Claude-only
      }
    }
  };
})();
