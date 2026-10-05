// ScoreDesk – Call Stats: import the 3CX CallStats workbook straight from SharePoint.
//
// An admin saves the SharePoint link to the file (or to the folder it lands in)
// plus an Azure app Client ID. Each time Call Stats opens, the file is checked
// through Microsoft Graph and, if it changed since the last import, the roster
// agents' figures are saved, exactly as a manual upload with the roster ticked.
// Relies on globals from agent-progression.html (supabase, parseWorkbook,
// rosterSet, key, NOT_AGENT_RE, saveStats, loadData, openPreview, showToast).

const SP = (() => {
  const CFG_KEY = 'scoredesk.sharepoint.v1';
  const MSAL_URL = 'https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.28.1/lib/msal-browser.min.js';
  const SCOPES = ['Files.Read.All'];
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  const el = id => document.getElementById(id);
  const h = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  const cfg = () => { try { return JSON.parse(localStorage.getItem(CFG_KEY) || '{}'); } catch (e) { return {}; } };
  const saveCfg = c => { try { localStorage.setItem(CFG_KEY, JSON.stringify(c)); } catch (e) {} };
  const configured = () => { const c = cfg(); return !!(c.clientId && c.url); };
  const when = iso => new Date(iso).toLocaleString('en-GB', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' });

  // ------------------------------------------------------------ Microsoft sign-in
  let msalReady = null, pca = null, pcaFor = '';
  function loadMsal() {
    if (window.msal) return Promise.resolve();
    if (!msalReady) msalReady = new Promise((res, rej) => {
      const s = document.createElement('script'); s.src = MSAL_URL;
      s.onload = res; s.onerror = () => { msalReady = null; rej(new Error('Could not load Microsoft sign-in')); };
      document.head.appendChild(s);
    });
    return msalReady;
  }
  async function client() {
    const c = cfg(), sig = c.clientId + '|' + (c.tenant || '');
    await loadMsal();
    if (!pca || pcaFor !== sig) {
      pca = new msal.PublicClientApplication({
        auth: {
          clientId: c.clientId,
          authority: 'https://login.microsoftonline.com/' + (c.tenant || 'organizations'),
          // a blank page, so the sign-in popup doesn't boot the whole app
          redirectUri: new URL('msal-redirect.html', location.href).href.split('?')[0]
        },
        cache: { cacheLocation: 'localStorage' }
      });
      await pca.initialize(); pcaFor = sig;
    }
    return pca;
  }
  class NeedsSignIn extends Error {}
  async function token(interactive) {
    const app = await client();
    const account = app.getAllAccounts()[0];
    if (account) {
      try { return (await app.acquireTokenSilent({ scopes: SCOPES, account })).accessToken; }
      catch (e) { if (!interactive) throw new NeedsSignIn('Microsoft sign-in expired'); }
    }
    if (!interactive) throw new NeedsSignIn('Sign in to Microsoft to check SharePoint');
    const r = await app.acquireTokenPopup({ scopes: SCOPES, prompt: 'select_account' });
    return r.accessToken;
  }

  // ------------------------------------------------------------ Graph
  async function graph(path, tok) {
    const r = await fetch(path.startsWith('http') ? path : GRAPH + path, { headers: { Authorization: 'Bearer ' + tok } });
    if (r.status === 401) throw new NeedsSignIn('Microsoft sign-in expired');
    if (r.status === 403) throw new Error("Your Microsoft account doesn't have access to that SharePoint file");
    if (r.status === 404) throw new Error("SharePoint file not found – check the link in SharePoint settings");
    if (!r.ok) throw new Error(`SharePoint error ${r.status}: ${(await r.text()).slice(0, 160)}`);
    return r;
  }
  // Any SharePoint/OneDrive link (sharing link or address-bar URL) -> Graph share id
  const shareId = url => 'u!' + btoa(unescape(encodeURIComponent(url.trim()))).replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');

  // The workbook to import: the linked file, or the newest Excel file in the linked folder.
  async function locate(tok) {
    const item = await (await graph(`/shares/${shareId(cfg().url)}/driveItem`, tok)).json();
    if (!item.folder) return item;
    const kids = await (await graph(`/drives/${item.parentReference.driveId}/items/${item.id}/children?$top=999`, tok)).json();
    const xl = (kids.value || []).filter(f => f.file && /\.xlsx?$/i.test(f.name))
      .sort((a, b) => b.lastModifiedDateTime.localeCompare(a.lastModifiedDateTime));
    if (!xl.length) throw new Error(`No Excel file in the SharePoint folder "${item.name}"`);
    return xl[0];
  }
  async function download(item, tok) {
    const direct = item['@microsoft.graph.downloadUrl'];
    const r = direct ? await fetch(direct) : await graph(`/drives/${item.parentReference.driveId}/items/${item.id}/content`, tok);
    if (!r.ok) throw new Error('Could not download the file from SharePoint (' + r.status + ')');
    return r.arrayBuffer();
  }

  // ------------------------------------------------------------ sync
  let busy = false;
  async function sync(interactive, force) {
    if (!configured()) { if (interactive) settings(); return; }
    if (busy) return; busy = true;
    const btn = el('spSyncBtn'); btn.disabled = true; btn.textContent = '⟳ Syncing…';
    try {
      const tok = await token(interactive);
      banner('Checking SharePoint for a new CallStats file…');
      const item = await locate(tok);
      const c = cfg();
      if (!force && c.lastETag && c.lastETag === item.eTag) {
        banner(`SharePoint: no new file since ${when(c.lastSync)} · <b>${h(item.name)}</b>, updated ${when(item.lastModifiedDateTime)}`, 'info', 'Re-import');
        return;
      }
      banner(`Downloading <b>${h(item.name)}</b> from SharePoint…`);
      const wb = XLSX.read(new Uint8Array(await download(item, tok)), { type: 'array', cellDates: false });
      const parsed = parseWorkbook(wb);

      // Same choice as a manual upload with the roster ticked, every queue counted.
      const roster = rosterSet();
      const names = [...new Set(parsed.rows.map(r => r.agent))];
      const sel = new Set(names.filter(n => roster.has(key(n)) && !NOT_AGENT_RE.test(n)));
      if (!sel.size) throw new Error('None of the roster agents are in the SharePoint file – check it is the 3CX CallStats export');
      const queues = new Set(parsed.rows.map(r => r.queue));
      const { data: { user } } = await supabase.auth.getUser();
      const saved = await saveStats(parsed, sel, queues, user.id);

      // Agents with recent calls on our 6xx extensions who aren't on the roster yet.
      const last = parsed.rows.reduce((m, r) => r.date > m ? r.date : m, '');
      const recentFrom = new Date(Date.parse(last) - 30 * 864e5).toISOString().slice(0, 10);
      const newcomers = [...new Set(parsed.rows.filter(r => r.date >= recentFrom && r.calls > 0 && /^6\d\d$/.test(r.ext) &&
        !roster.has(key(r.agent)) && !NOT_AGENT_RE.test(r.agent)).map(r => r.agent))];

      saveCfg({ ...cfg(), lastETag: item.eTag, lastSync: new Date().toISOString(), lastFile: item.name, lastModified: item.lastModifiedDateTime });
      await loadData();
      banner(`SharePoint synced ${when(new Date().toISOString())} · <b>${h(item.name)}</b> (updated ${when(item.lastModifiedDateTime)}) · ${saved.toLocaleString('en-GB')} agent-days for ${sel.size} agents` +
        (newcomers.length ? `<br>⚠ Not on the roster, so skipped: <b>${newcomers.map(h).join(', ')}</b>` : ''),
        newcomers.length ? 'warn' : 'info', newcomers.length ? 'Review agents' : null,
        newcomers.length ? () => openPreview({ name: item.name }, parsed) : null);
      if (interactive) showToast(`Imported ${saved.toLocaleString('en-GB')} agent-days from SharePoint`, true);
    } catch (e) {
      if (e instanceof NeedsSignIn) banner(`${h(e.message)} – click to sign in, then new CallStats files import automatically.`, 'info', 'Sign in & sync');
      else { console.error(e); banner('SharePoint sync failed: ' + h(e.message || e), 'bad', 'Try again'); }
    } finally { busy = false; btn.disabled = false; btn.textContent = '⟳ Sync SharePoint'; }
  }

  function banner(html, kind = 'info', action, onAction) {
    el('spBanner').innerHTML = `<div class="banner ${kind}"><span>☁️</span><span class="grow">${html}</span>${action ? `<button class="btn" id="spAct">${h(action)}</button>` : ''}</div>`;
    if (action) el('spAct').onclick = onAction || (() => sync(true, action === 'Re-import'));
  }

  // ------------------------------------------------------------ settings
  function settings() {
    const c = cfg();
    const redirect = new URL('msal-redirect.html', location.href).href.split('?')[0];
    const root = el('modalRoot');
    root.innerHTML = `<div class="ov" id="spOv"><div class="modal">
      <div class="mh"><div><div class="eyebrow">SharePoint</div><div class="t">CallStats file</div>
        <div class="sub">Each time Call Stats opens, this file is checked and imported if it has changed.</div></div><button class="xbtn" data-x>✕</button></div>
      <div class="mb2">
        <div class="field"><label>SharePoint link to the file (or its folder)</label><input class="inp" id="spUrl" placeholder="https://yourcompany.sharepoint.com/:x:/s/…" value="${h(c.url || '')}"></div>
        <div class="facts" style="grid-template-columns:1fr 1fr">
          <div class="field"><label>Azure app Client ID</label><input class="inp" id="spClient" placeholder="00000000-0000-0000-0000-000000000000" value="${h(c.clientId || '')}"></div>
          <div class="field"><label>Tenant ID (optional)</label><input class="inp" id="spTenant" placeholder="organizations" value="${h(c.tenant || '')}"></div>
        </div>
        <details ${c.clientId ? '' : 'open'}><summary class="sub" style="cursor:pointer">How to get a Client ID (one-off, ~5 minutes, Microsoft 365 admin may be needed)</summary>
        <ol class="steps" style="margin-top:8px">
          <li>Open <code>entra.microsoft.com</code> → App registrations → <b>New registration</b> (e.g. "ScoreDesk").</li>
          <li>Supported accounts: <b>this organisation only</b>.</li>
          <li>Redirect URI: platform <b>Single-page application (SPA)</b>, URL <code>${h(redirect)}</code></li>
          <li>API permissions → Add → Microsoft Graph → Delegated → <b>Files.Read.All</b> (grant admin consent if your tenant requires it).</li>
          <li>Copy the <b>Application (client) ID</b> and <b>Directory (tenant) ID</b> from the Overview page into the boxes above.</li>
          <li>File link: in SharePoint, open the file's <b>⋯ → Copy link</b> (or copy the folder link if a new file is dropped there each day).</li>
        </ol></details>
        ${c.lastSync ? `<p class="sub" style="margin-top:12px">Last import ${when(c.lastSync)} · ${h(c.lastFile || '')}</p>` : ''}
      </div>
      <div class="mf">${c.url ? '<button class="btn danger" id="spForget">Disconnect</button>' : ''}<span class="grow"></span>
        <button class="btn" data-x>Cancel</button><button class="btn primary" id="spSave">Save &amp; sync</button></div>
    </div></div>`;
    const close = () => { root.innerHTML = ''; };
    root.querySelectorAll('[data-x]').forEach(b => b.onclick = close);
    el('spOv').onclick = e => { if (e.target.id === 'spOv') close(); };
    el('spSave').onclick = () => {
      const url = el('spUrl').value.trim(), clientId = el('spClient').value.trim(), tenant = el('spTenant').value.trim();
      if (!/^https:\/\/[^/]+\.(sharepoint\.com|onedrive\.live\.com|1drv\.ms)/i.test(url)) return showToast('Paste a SharePoint link (https://…sharepoint.com/…)', false);
      if (!/^[0-9a-f-]{36}$/i.test(clientId)) return showToast('The Client ID looks like 00000000-0000-0000-0000-000000000000', false);
      const prev = cfg();
      saveCfg({ ...prev, url, clientId, tenant, ...(prev.url !== url ? { lastETag: null } : {}) });
      close(); sync(true, true);
    };
    if (el('spForget')) el('spForget').onclick = () => { localStorage.removeItem(CFG_KEY); el('spBanner').innerHTML = ''; close(); showToast('SharePoint disconnected', true); };
  }

  // ------------------------------------------------------------ wiring
  function onOpen() {
    el('spCfgBtn').onclick = settings;
    el('spSyncBtn').onclick = () => sync(true, false);
    if (!configured()) return;
    const c = cfg();
    if (c.lastSync) banner(`SharePoint: last import ${when(c.lastSync)} · ${h(c.lastFile || '')}`);
    loadMsal().then(() => sync(false, false)).catch(e => banner('SharePoint sync failed: ' + h(e.message), 'bad', 'Try again'));
  }

  return { onOpen, sync, settings };
})();
