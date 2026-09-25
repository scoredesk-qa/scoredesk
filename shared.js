// ScoreDesk – shared helpers for the tab pages (Dashboard, Quality, Zero Score).
// Loaded after supabase-js (and xlsx where scorecards are parsed).

const SD = (() => {
  const SUPABASE_URL = 'https://qbvjgiamqzkscdykenvi.supabase.co';
  const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InFidmpnaWFtcXprc2NkeWtlbnZpIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzUwMzUxMTIsImV4cCI6MjA5MDYxMTExMn0.OeV5qa9y_PCGw6nG0wbfPfiMBjYC_jvzsUXHOip1RTU';
  const sb = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
    auth: { autoRefreshToken: true, persistSession: true }
  });

  // ---------------------------------------------------------------- basics
  const $ = id => document.getElementById(id);
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const norm = s => String(s ?? '').trim().replace(/\s+/g, ' ');
  const pad = n => String(n).padStart(2, '0');
  const fmtN = n => Math.round(n).toLocaleString('en-GB');
  const pct = (a, b) => b > 0 ? a / b * 100 : null;

  function toast(msg, ok) {
    const t = document.createElement('div');
    t.className = 'toast' + (ok === true ? ' ok' : ok === false ? ' err' : '');
    t.textContent = msg;
    document.body.appendChild(t);
    setTimeout(() => t.classList.add('show'), 10);
    setTimeout(() => { t.classList.remove('show'); setTimeout(() => t.remove(), 300); }, 4000);
  }

  // ---------------------------------------------------------------- dates ('YYYY-MM-DD' strings, UTC maths)
  const dUTC = s => { const [y, m, d] = s.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)); };
  const sUTC = d => `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
  const addDays = (s, n) => { const d = dUTC(s); d.setUTCDate(d.getUTCDate() + n); return sUTC(d); };
  const mondayOf = s => { const d = dUTC(s); const wd = d.getUTCDay() || 7; d.setUTCDate(d.getUTCDate() - wd + 1); return sUTC(d); };
  const fmtDate = (s, o = { day: '2-digit', month: 'short', year: 'numeric' }) => s ? dUTC(s).toLocaleDateString('en-GB', { ...o, timeZone: 'UTC' }) : '—';
  const shortDate = s => fmtDate(s, { day: 'numeric', month: 'short' });
  const monthLbl = ym => fmtDate(ym + '-01', { month: 'short', year: 'numeric' });
  const today = () => sUTC(new Date());
  const daysBetween = (a, b) => Math.round((dUTC(b) - dUTC(a)) / 86400000);
  function weekList(from, to) { const out = []; let w = mondayOf(from); while (w <= to) { out.push(w); w = addDays(w, 7); } return out; }

  function rangeBounds(range, maxDate, minDate) {
    if (range === 'all') return [minDate, maxDate];
    if (range === 'mtd') return [maxDate.slice(0, 8) + '01', maxDate];
    if (range === 'lm') { const d = dUTC(maxDate.slice(0, 8) + '01'); d.setUTCDate(0); const l = sUTC(d); return [l.slice(0, 8) + '01', l]; }
    return [addDays(maxDate, -(parseInt(range, 10) - 1)), maxDate];
  }

  const MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, sept: 9, oct: 10, nov: 11, dec: 12 };

  // Parses any call-date format seen in the scorecards and returns
  // { date:'YYYY-MM-DD', time:'HH:MM'|'' } or null.
  //  - Excel serial numbers (number or numeric string)
  //  - "2026-09-13", "2026-09-13 14:05", ISO timestamps
  //  - "13 Sept 2026, 14:05" / "13 Sep 2026" (what older uploads stored)
  //  - "13/09/2026 14:05" (UK order; swapped only when the month would be > 12)
  function parseCallDate(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date && !isNaN(v)) return { date: `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`, time: `${pad(v.getHours())}:${pad(v.getMinutes())}` };
    const s = String(v).trim();
    if (/^\d{5}(\.\d+)?$/.test(s)) {
      const n = parseFloat(s);
      if (n > 20000 && n < 80000) {
        const ms = Math.round((n - 25569) * 86400000);
        const d = new Date(ms);
        const hasTime = n % 1 > 0;
        return { date: sUTC(d), time: hasTime ? `${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}` : '' };
      }
    }
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s](\d{1,2}):(\d{2}))?/);
    if (m) return { date: `${m[1]}-${pad(m[2])}-${pad(m[3])}`, time: m[4] ? `${pad(m[4])}:${m[5]}` : '' };
    m = s.match(/^(\d{1,2})\s+([A-Za-z]{3,9})\.?,?\s+(\d{4})(?:[,\s]+(\d{1,2}):(\d{2}))?/);
    if (m) {
      const mo = MONTHS[m[2].toLowerCase()] || MONTHS[m[2].slice(0, 3).toLowerCase()];
      if (mo) return { date: `${m[3]}-${pad(mo)}-${pad(m[1])}`, time: m[4] ? `${pad(m[4])}:${m[5]}` : '' };
    }
    m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})(?:[,\sT]+(\d{1,2}):(\d{2}))?/);
    if (m) {
      let day = +m[1], mon = +m[2], yr = +m[3];
      if (yr < 100) yr += 2000;
      if (mon > 12 && day <= 12) [day, mon] = [mon, day];
      if (mon >= 1 && mon <= 12 && day >= 1 && day <= 31) return { date: `${yr}-${pad(mon)}-${pad(day)}`, time: m[4] ? `${pad(m[4])}:${m[5]}` : '' };
    }
    m = s.match(/^([A-Za-z]{3,9})\s+(\d{1,2}),?\s+(\d{4})/); // "Sep 13, 2026"
    if (m && MONTHS[m[1].slice(0, 3).toLowerCase()]) return { date: `${m[3]}-${pad(MONTHS[m[1].slice(0, 3).toLowerCase()])}-${pad(m[2])}`, time: '' };
    return null;
  }

  // ---------------------------------------------------------------- auth / data
  async function session() {
    const { data: { user } } = await sb.auth.getUser();
    if (!user) return null;
    const meta = user.user_metadata || {};
    return { user, isAdmin: user.email === 'admin@scoredesk.com' || meta.role === 'admin', agentName: meta.agent_name || null };
  }

  function notSignedIn() {
    document.querySelector('.main').innerHTML = '<div class="loading"><h2 style="font-family:Syne;color:var(--text);margin-bottom:8px">Not signed in</h2>Please log in via the ScoreDesk portal.</div>';
  }

  // Pages through a table (PostgREST returns at most 1000 rows per request).
  async function fetchAll(table, cols, tweak) {
    const PAGE = 1000; let out = [], from = 0;
    while (true) {
      let q = sb.from(table).select(cols).order('id', { ascending: true }).range(from, from + PAGE - 1);
      if (tweak) q = tweak(q);
      const { data, error } = await q;
      if (error) throw error;
      out = out.concat(data || []);
      if (!data || data.length < PAGE) break;
      from += PAGE;
    }
    return out;
  }

  // ---------------------------------------------------------------- agent names
  const nameKey = n => norm(n).toLowerCase().replace(/[^a-z0-9\s]/g, '').split(' ').filter(Boolean).sort().join(' ');
  function lev(a, b) {
    const m = a.length, n = b.length, d = Array.from({ length: m + 1 }, (_, i) => [i]);
    for (let j = 1; j <= n; j++) d[0][j] = j;
    for (let i = 1; i <= m; i++) for (let j = 1; j <= n; j++)
      d[i][j] = a[i - 1] === b[j - 1] ? d[i - 1][j - 1] : 1 + Math.min(d[i - 1][j - 1], d[i - 1][j], d[i][j - 1]);
    return d[m][n];
  }
  // Maps spelling variants of the same person onto one display name.
  function nameResolver(canonicalNames = []) {
    const known = new Map();
    const add = n => { const k = nameKey(n); if (k && !known.has(k)) known.set(k, norm(n)); };
    canonicalNames.forEach(add);
    return raw => {
      const k = nameKey(raw);
      if (!k) return norm(raw);
      if (known.has(k)) return known.get(k);
      for (const [kk, v] of known) if (Math.abs(kk.length - k.length) <= 2 && lev(kk, k) <= 2) { known.set(k, v); return v; }
      add(raw); return norm(raw);
    };
  }

  // ---------------------------------------------------------------- scorecards
  const AF_ALIASES = {
    'data protection/auto fail if not passed': 'Data Protection',
    'were three points for data protection confirmed?': 'Data Protection',
    'has authority, if needed, been appropriately captured or granted to an unauthorised third party? auto fail if not passed': 'Authority for Third Party',
    'was relevant authority obtained for third party': 'Authority for Third Party',
    'pci dss compliant? call recording restarted? (auto fail if not)': 'PCI DSS Compliant',
    'have the notes been added? (autofail if not)': 'Notes Added',
    'do the notes clearly and accurately reflect the call content? has all data captured been correctly recorded? (auto fail if not)': 'Notes Accurate & Complete'
  };
  const critKey = l => norm(l).toLowerCase();
  const canonicalAF = l => AF_ALIASES[critKey(l)] || norm(l);

  const cellStr = v => String(v ?? '').trim();

  // Finds the sheet holding the scorecard (the one with "Advisor Name"), else the first.
  function scorecardRows(wb) {
    for (const n of wb.SheetNames) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: '', raw: true });
      if (rows.some(r => /advisor name/i.test(cellStr(r[0])))) return rows;
    }
    return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '', raw: true });
  }

  // Returns a row ready to insert into `scorecards` (plus parse warnings).
  function extractScorecard(rows, fileName) {
    const warnings = [];
    const field = re => { const r = rows.find(r => re.test(cellStr(r[0]))); if (!r) return ''; return r[2] !== '' && r[2] != null ? r[2] : r[1]; };
    const name = norm(cellStr(field(/advisor name/i))) || fileName.replace(/\.(xlsx?|xlsm)$/i, '');
    if (!cellStr(field(/advisor name/i))) warnings.push('No "Advisor Name" – used the file name');
    const account = cellStr(field(/account reference/i));
    const auditor = cellStr(field(/auditor/i));
    const rawDate = field(/call date/i);
    const cd = parseCallDate(rawDate);
    if (!cd) warnings.push(`Call date not recognised: "${cellStr(rawDate)}"`);

    const pfRow = rows.find(r => /^pass\s*\/?\s*fail/i.test(cellStr(r[0])));
    const pfResult = pfRow ? cellStr(pfRow[1]).toUpperCase() : '';
    let pfReason = pfRow ? cellStr(pfRow[2]) : '';
    if (!pfRow) warnings.push('No Pass/Fail row');
    const pass = pfResult === 'PASS';

    const zzRow = rows.find(r => /ZZPS Quality Score/i.test(cellStr(r[0])));
    const zzRaw = zzRow ? (zzRow[2] !== '' ? zzRow[2] : zzRow[1]) : '';
    const zzNum = parseFloat(zzRaw);
    const zzPct = isFinite(zzNum) ? (zzNum <= 1 ? Math.round(zzNum * 100) : Math.round(zzNum)) : null;

    const autoFails = [], criteria = [];
    let inZ = false, section = '';
    for (const r of rows) {
      const l = cellStr(r[0]), c1 = cellStr(r[1]), c2 = cellStr(r[2]), c3 = cellStr(r[3]);
      if (!l) continue;
      if (/ZZPS Quality Score/i.test(l)) { inZ = true; continue; }
      if (/^pass\s*\/?\s*fail/i.test(l)) { inZ = false; continue; }
      if (inZ) {
        const v = c1.toUpperCase();
        autoFails.push({ source: 'zzps', section: 'ZZPS Summary', label: l, value: v || 'N/A', feedback: c2, triggered: v === 'NO' });
        continue;
      }
      if (/^(Introduction|Data Protection|Body|Conclusion)\s*$/i.test(l)) { section = l; continue; }
      if (/auto.?fail/i.test(c1)) {
        const v = c2.toUpperCase();
        autoFails.push({ source: 'criteria', section, label: l, value: v || 'N/A', feedback: c3, triggered: v === 'NO' });
        continue;
      }
      const w = parseFloat(r[1]), s = parseFloat(r[2]);
      if (isFinite(w) && w > 0 && !/^(total|overall|zzps quality score)\b/i.test(l)) criteria.push({ section, label: l, weight: w, score: isFinite(s) ? s : null, feedback: c3 });
    }
    if (!pass && !pfReason) {
      const trig = autoFails.filter(a => a.triggered);
      if (trig.length) pfReason = 'AF ' + trig.map(a => canonicalAF(a.label)).join(', ');
      else if (zzPct != null) pfReason = zzPct + '% - Below Pass Rate';
    }
    if (!criteria.length) warnings.push('No scored criteria found');
    return {
      row: {
        agent_name: name, file_name: fileName,
        call_date: cd ? (cd.date + (cd.time ? ' ' + cd.time : '')) : cellStr(rawDate),
        auditor, account_ref: account, pass, pf_reason: pfReason,
        auto_fails: autoFails, criteria
      },
      warnings
    };
  }

  // Same call = same agent + call date/time + account. Falls back to the file name.
  function scorecardKey(r) {
    const cd = parseCallDate(r.call_date);
    if (cd && r.account_ref) return [nameKey(r.agent_name), cd.date, cd.time, norm(r.account_ref).toLowerCase()].join('|');
    return 'file|' + nameKey(r.agent_name) + '|' + norm(r.file_name).toLowerCase();
  }

  // Normalises a stored scorecard row for analysis.
  function enrichScorecard(r, resolveName) {
    const cd = parseCallDate(r.call_date) || parseCallDate(r.created_at);
    const crit = Array.isArray(r.criteria) ? r.criteria : [];
    let got = 0, max = 0;
    for (const c of crit) { const w = parseFloat(c.weight), s = parseFloat(c.score); if (w > 0 && isFinite(s)) { got += s; max += w; } }
    const afs = (Array.isArray(r.auto_fails) ? r.auto_fails : []).map(a => {
      const value = cellStr(a.value ?? a.val).toUpperCase();
      return { ...a, value, feedback: a.feedback ?? a.fb ?? '', triggered: a.triggered ?? value === 'NO' };
    });
    return {
      ...r, agent: resolveName ? resolveName(r.agent_name) : norm(r.agent_name),
      date: cd ? cd.date : null, time: cd ? cd.time : '',
      score: max > 0 ? got / max * 100 : null,
      afs, crit, afTriggered: afs.filter(a => a.triggered).length
    };
  }

  // ---------------------------------------------------------------- issues (zero scores + triggered auto-fails)
  // Each agent × criterion is tracked through that agent's scored calls in date order.
  // A failure opens an issue; `cleanNeeded` consecutive clean calls on that criterion resolve it.
  function buildIssues(cards, cleanNeeded = 2) {
    const tl = new Map();
    const sorted = cards.filter(c => c.date).sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time));
    for (const c of sorted) {
      const seen = new Set();
      const push = (key, label, type, section, fail, feedback) => {
        if (seen.has(key)) return; seen.add(key);
        const k = c.agent + '§' + key;
        if (!tl.has(k)) tl.set(k, { agent: c.agent, key, label, type, section, occ: [] });
        tl.get(k).occ.push({ date: c.date, fail, feedback: cellStr(feedback), card: c });
      };
      for (const cr of c.crit) {
        const w = parseFloat(cr.weight), s = parseFloat(cr.score);
        if (!(w > 0) || !isFinite(s) || !cr.label) continue;
        push('c:' + critKey(cr.label), norm(cr.label), 'Zero score', cr.section || '', s === 0, cr.feedback);
      }
      for (const a of c.afs) {
        if (!a.label || !a.value || a.value === 'N/A') continue;
        const lbl = canonicalAF(a.label);
        push('a:' + critKey(lbl), lbl, 'Auto-fail', a.section || 'ZZPS Summary', a.value === 'NO', a.feedback);
      }
    }
    const episodes = [];
    for (const t of tl.values()) {
      let open = null, clean = 0, count = 0;
      for (const o of t.occ) {
        if (o.fail) {
          if (!open) { open = { ...t, occ: undefined, emerged: o.date, reopened: count > 0, fails: 0, assessed: 0, lastFail: null, lastFeedback: '', resolved: null, cleanSince: 0, lastCard: null }; count++; }
          open.fails++; open.lastFail = o.date; open.lastCard = o.card; if (o.feedback) open.lastFeedback = o.feedback;
          clean = 0;
        } else if (open) {
          clean++;
          if (clean >= cleanNeeded) { open.resolved = o.date; open.cleanSince = clean; episodes.push(open); open = null; clean = 0; }
        }
        if (open) { open.assessed++; open.cleanSince = clean; }
      }
      if (open) episodes.push(open);
    }
    return episodes;
  }

  return {
    sb, $, esc, norm, pad, fmtN, pct, toast,
    dUTC, sUTC, addDays, mondayOf, fmtDate, shortDate, monthLbl, today, daysBetween, weekList, rangeBounds,
    parseCallDate, session, notSignedIn, fetchAll,
    nameKey, nameResolver, canonicalAF, critKey,
    scorecardRows, extractScorecard, scorecardKey, enrichScorecard, buildIssues
  };
})();

// Chart.js defaults shared by every page
if (window.Chart) {
  Chart.defaults.color = '#6b7694';
  Chart.defaults.font.family = 'DM Mono, monospace';
  Chart.defaults.font.size = 10;
}
const GRID = 'rgba(255,255,255,0.05)';
const TIP = { backgroundColor: '#191d28', borderColor: '#2e3448', borderWidth: 1, titleColor: '#e8ecf4', bodyColor: '#a8b1cc', padding: 10 };
