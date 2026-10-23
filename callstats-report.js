// ScoreDesk – Call Stats: printable A4 performance report per agent, compared
// with the team. Uses globals from agent-progression.html: ROWS, sumRows,
// groupBy, mondayOf, fmtDate, fmtTalk, esc, showToast.

const CSReport = (() => {
  // print palette (light), chosen to stay distinct in colour-blind vision and in greyscale
  const C = { agent: '#7f9f12', team: '#6b6e66', crit: '#c93c3c', grid: '#e4e6df', ink: '#121310', muted: '#6b6e66', soft: '#f3f4ef', lime: '#e5ff8f' };
  const LOGO = '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="#121400" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 5.5h12a2.5 2.5 0 0 1 2.5 2.5v7a2.5 2.5 0 0 1-2.5 2.5h-5.2L8.5 20.8v-3.3H6A2.5 2.5 0 0 1 3.5 15V8A2.5 2.5 0 0 1 6 5.5z"/><path d="M8.6 11.6l2.3 2.3 4.6-4.7"/></svg>';
  const n1 = v => v == null || !isFinite(v) ? '—' : v.toFixed(1);
  const pc = v => v == null || !isFinite(v) ? '—' : v.toFixed(1) + '%';
  const int = v => Math.round(v).toLocaleString('en-GB');
  const sd = k => fmtDate(k, { day: 'numeric', month: 'short' });

  // ------------------------------------------------------------ data
  function compute(agent, from, to) {
    const inR = ROWS.filter(r => r.date >= from && r.date <= to);
    const days = Math.round((Date.parse(to) - Date.parse(from)) / 864e5) + 1;
    const unit = days <= 31 ? 'day' : days <= 190 ? 'week' : 'month';
    const keyOf = unit === 'day' ? d => d : unit === 'week' ? mondayOf : d => d.slice(0, 7);
    const label = unit === 'day' ? sd : unit === 'week' ? sd : k => fmtDate(k + '-01', { month: 'short', year: '2-digit' });

    const mine = inR.filter(r => r.agent === agent);
    const keys = [...new Set(inR.filter(r => r.calls > 0).map(r => keyOf(r.date)))].sort();
    const byKeyTeam = groupBy(inR, r => keyOf(r.date)), byKeyMe = groupBy(mine, r => keyOf(r.date));
    const series = keys.map(k => {
      const t = byKeyTeam.get(k) || [], me = sumRows(byKeyMe.get(k) || []), ts = sumRows(t);
      const active = new Set(t.filter(r => r.calls > 0).map(r => r.agent)).size;
      return { k, label: label(k), me, team: ts, teamPerAgent: active ? ts.calls / active : null };
    });

    // every agent's totals for the ranking and team averages
    const byAgent = groupBy(inR, r => r.agent);
    const agents = [...byAgent.entries()].map(([name, rs]) => ({ name, ...sumRows(rs) })).filter(a => a.calls > 0);
    const avg = f => { const v = agents.map(f).filter(x => x != null && isFinite(x)); return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null; };
    const me = sumRows(mine), team = sumRows(inR);
    const rank = [...agents].sort((a, b) => b.perDay - a.perDay).findIndex(a => a.name === agent) + 1;
    return {
      agent, from, to, unit, series, me, agents,
      teamAvg: { calls: avg(a => a.calls), perDay: avg(a => a.perDay), days: avg(a => a.days), talkMin: avg(a => a.talkMin), ignRate: team.ignRate, avgTalk: team.avgTalk },
      rank, of: agents.length
    };
  }

  // ------------------------------------------------------------ svg charts
  function scale(max) {
    if (!(max > 0)) return { top: 1, step: .25 };
    const raw = max / 4, mag = Math.pow(10, Math.floor(Math.log10(raw)));
    const step = [1, 2, 2.5, 5, 10].map(m => m * mag).find(s => s >= raw);
    return { top: step * Math.ceil(max / step), step };
  }
  // bars (agent) with an optional dashed reference line (team), one y-axis
  function barLine({ labels, bars, line, w = 680, h = 190, fmt = int, barColor = C.agent, lineAsBars = false }) {
    const pl = 40, pr = 8, pt = 8, pb = 24, iw = w - pl - pr, ih = h - pt - pb;
    const { top, step } = scale(Math.max(...bars.filter(v => v != null), ...(line || []).filter(v => v != null), 0));
    const y = v => pt + ih - (v / top) * ih, n = labels.length, bw = iw / Math.max(n, 1);
    let s = `<svg viewBox="0 0 ${w} ${h}" width="100%" style="display:block;height:auto" font-family="Inter,Arial,sans-serif">`;
    for (let v = 0; v <= top + 1e-9; v += step) s += `<line x1="${pl}" x2="${w - pr}" y1="${y(v)}" y2="${y(v)}" stroke="${C.grid}"/><text x="${pl - 6}" y="${y(v) + 3}" font-size="9" fill="${C.muted}" text-anchor="end">${fmt(v)}</text>`;
    const every = Math.ceil(n / Math.max(3, Math.floor(iw / 58)));  // keep date labels from colliding
    bars.forEach((v, i) => {
      const x = pl + i * bw, bwid = Math.min(26, bw * .62);
      if (v != null && v > 0) s += `<rect x="${x + (bw - bwid) / 2}" y="${y(v)}" width="${bwid}" height="${Math.max(0, pt + ih - y(v))}" rx="2.5" fill="${barColor}"/>`;
      if (i % every === 0) s += `<text x="${x + bw / 2}" y="${h - 8}" font-size="9" fill="${C.muted}" text-anchor="middle">${esc(labels[i])}</text>`;
    });
    if (line) {
      const pts = line.map((v, i) => v == null ? null : [pl + i * bw + bw / 2, y(v)]);
      let d = ''; pts.forEach((p, i) => { if (p) d += (d && pts[i - 1] ? 'L' : 'M') + p[0].toFixed(1) + ',' + p[1].toFixed(1); });
      s += `<path d="${d}" fill="none" stroke="${C.team}" stroke-width="1.8" stroke-dasharray="5 4"/>`;
    }
    return s + `<line x1="${pl}" x2="${w - pr}" y1="${pt + ih}" y2="${pt + ih}" stroke="${C.muted}" stroke-width=".8"/></svg>`;
  }
  // two lines on one axis: agent (solid) vs team (dashed)
  function lines({ labels, a, b, w = 330, h = 160, fmt = int, min0 = true }) {
    const pl = 38, pr = 8, pt = 8, pb = 22, iw = w - pl - pr, ih = h - pt - pb;
    const { top, step } = scale(Math.max(...a.filter(v => v != null), ...b.filter(v => v != null), 0));
    const n = labels.length, xs = i => pl + (n === 1 ? iw / 2 : (i / (n - 1)) * iw), y = v => pt + ih - (v / top) * ih;
    let s = `<svg viewBox="0 0 ${w} ${h}" width="100%" style="display:block;height:auto" font-family="Inter,Arial,sans-serif">`;
    for (let v = 0; v <= top + 1e-9; v += step) s += `<line x1="${pl}" x2="${w - pr}" y1="${y(v)}" y2="${y(v)}" stroke="${C.grid}"/><text x="${pl - 5}" y="${y(v) + 3}" font-size="8.5" fill="${C.muted}" text-anchor="end">${fmt(v)}</text>`;
    const every = Math.ceil(n / 6);
    labels.forEach((l, i) => { if (i % every === 0) s += `<text x="${xs(i)}" y="${h - 7}" font-size="8.5" fill="${C.muted}" text-anchor="middle">${esc(l)}</text>`; });
    const path = (vals, col, dash, wdt) => { let d = ''; vals.forEach((v, i) => { if (v != null) d += (d && vals[i - 1] != null ? 'L' : 'M') + xs(i).toFixed(1) + ',' + y(v).toFixed(1); }); return `<path d="${d}" fill="none" stroke="${col}" stroke-width="${wdt}" ${dash ? 'stroke-dasharray="5 4"' : ''} stroke-linejoin="round"/>`; };
    s += path(b, C.team, true, 1.8) + path(a, C.agent, false, 2.2);
    if (n <= 16) a.forEach((v, i) => { if (v != null) s += `<circle cx="${xs(i)}" cy="${y(v)}" r="2.6" fill="#fff" stroke="${C.agent}" stroke-width="1.6"/>`; });
    return s + '</svg>';
  }
  // horizontal ranking with the agent highlighted and the team average marked
  function ranking({ items, me, avg, w = 330, rowH = 13 }) {
    const pl = 104, pr = 30, h = items.length * rowH + 22, iw = w - pl - pr, max = Math.max(...items.map(i => i.v), avg || 0) || 1;
    let s = `<svg viewBox="0 0 ${w} ${h}" width="100%" style="display:block;height:auto" font-family="Inter,Arial,sans-serif">`;
    items.forEach((it, i) => {
      const yy = 4 + i * rowH, bw = (it.v / max) * iw, on = it.name === me;
      s += `<text x="${pl - 6}" y="${yy + 9}" font-size="8.5" fill="${on ? C.ink : C.muted}" font-weight="${on ? 600 : 400}" text-anchor="end">${esc(it.name)}</text>`;
      s += `<rect x="${pl}" y="${yy + 1.5}" width="${Math.max(1, bw)}" height="${rowH - 4}" rx="2" fill="${on ? C.agent : '#d7dad2'}"/>`;
      s += `<text x="${pl + bw + 4}" y="${yy + 9}" font-size="8.5" fill="${on ? C.ink : C.muted}" font-weight="${on ? 600 : 400}">${n1(it.v)}</text>`;
    });
    if (avg) { const x = pl + (avg / max) * iw; s += `<line x1="${x}" x2="${x}" y1="2" y2="${h - 16}" stroke="${C.team}" stroke-dasharray="3 3" stroke-width="1.2"/><text x="${x}" y="${h - 5}" font-size="8.5" fill="${C.muted}" text-anchor="middle">team avg ${n1(avg)}</text>`; }
    return s + '</svg>';
  }

  // ------------------------------------------------------------ page
  function vs(me, team, better, fmt, unit = '') {
    if (me == null || team == null || !isFinite(me) || !isFinite(team) || team === 0) return '<span class="vs">—</span>';
    const d = (me / team - 1) * 100, good = better === 'higher' ? d >= 0 : d <= 0, flat = Math.abs(d) < 2;
    return `<span class="vs ${flat ? '' : good ? 'up' : 'dn'}">${d >= 0 ? '+' : ''}${d.toFixed(0)}% vs team (${fmt(team)}${unit})</span>`;
  }

  function page(R, idx, total, period) {
    const { me, teamAvg: T, series } = R;
    const unitWord = R.unit === 'day' ? 'day' : R.unit === 'week' ? 'week' : 'month';
    const labels = series.map(s => s.label);
    const tiles = [
      ['Answered calls', int(me.calls), vs(me.calls, T.calls, 'higher', int)],
      ['Calls per day worked', n1(me.perDay), vs(me.perDay, T.perDay, 'higher', n1)],
      ['Ignore rate', pc(me.ignRate), vs(me.ignRate, T.ignRate, 'lower', n1, '%')],
      ['Average talk time', fmtTalk(me.avgTalk), vs(me.avgTalk, T.avgTalk, 'neither', fmtTalk)],
      ['Days worked', int(me.days), vs(me.days, T.days, 'higher', n1)],
      ['Volume rank', R.rank ? `${R.rank}<small> of ${R.of}</small>` : '—', '<span class="vs">by calls per day</span>']
    ];
    // "neither": talk time is shown against the team without judging good/bad
    const rows = series.slice(-14);
    const rank = [...R.agents].sort((a, b) => b.perDay - a.perDay).map(a => ({ name: a.name, v: a.perDay }));
    return `<section class="page"><div class="inner">
      <header><div class="brand"><span class="logo">${LOGO}</span>ScoreDesk</div><div class="doc">Call performance report · ${idx + 1} of ${total}</div></header>
      <div class="titlerow"><div><h1>${esc(R.agent)}</h1><div class="period">${esc(period)} · compared with the team average (${R.of} agents)</div></div></div>
      <div class="tiles">${tiles.map(([l, v, s], i) => `<div class="tile ${i === 0 ? 'hl' : ''}"><div class="tl">${l}</div><div class="tv">${v}</div>${s}</div>`).join('')}</div>

      <div class="box"><div class="bh"><b>Calls answered per ${unitWord}</b><span class="lg"><i style="background:${C.agent}"></i>${esc(R.agent.split(' ')[0])}<i class="dash"></i>Team average per agent</span></div>
        ${series.length ? barLine({ labels, bars: series.map(s => s.me.calls), line: series.map(s => s.teamPerAgent) }) : '<p class="none">No calls in this period.</p>'}</div>

      <div class="two">
        <div class="box"><div class="bh"><b>Average talk time</b><span class="lg"><i style="background:${C.agent}"></i>${esc(R.agent.split(' ')[0])}<i class="dash"></i>Team</span></div>
          ${lines({ labels, a: series.map(s => s.me.avgTalk), b: series.map(s => s.team.avgTalk), fmt: fmtTalk })}</div>
        <div class="box"><div class="bh"><b>Ignore rate</b><span class="lg"><i style="background:${C.crit}"></i>${esc(R.agent.split(' ')[0])}<i class="dash"></i>Team</span></div>
          ${barLine({ labels, bars: series.map(s => s.me.ignRate), line: series.map(s => s.team.ignRate), w: 330, h: 160, fmt: v => v.toFixed(v < 10 ? 1 : 0) + '%', barColor: C.crit })}</div>
      </div>

      <div class="two">
        <div class="box"><div class="bh"><b>Calls per day worked – all agents</b></div>${ranking({ items: rank, me: R.agent, avg: T.perDay })}</div>
        <div class="box"><div class="bh"><b>By ${unitWord}</b>${series.length > rows.length ? `<span class="lg">last ${rows.length}</span>` : ''}</div>
          <table><thead><tr><th>${unitWord[0].toUpperCase() + unitWord.slice(1)}</th><th>Answered</th><th>Team avg</th><th>Ignore</th><th>Avg talk</th><th>Team talk</th></tr></thead>
          <tbody>${rows.map(s => `<tr><td>${esc(s.label)}</td><td><b>${int(s.me.calls)}</b></td><td>${s.teamPerAgent == null ? '—' : n1(s.teamPerAgent)}</td><td class="${s.me.ignRate >= 5 ? 'bad' : ''}">${pc(s.me.ignRate)}</td><td>${fmtTalk(s.me.avgTalk)}</td><td>${fmtTalk(s.team.avgTalk)}</td></tr>`).join('')}</tbody></table></div>
      </div>

      <div class="notes"><b>Notes / actions</b></div>
      <footer><div class="sign"><span>Agent signature</span><span>Team leader signature</span><span>Date</span></div>
        <div class="gen">Generated ${new Date().toLocaleString('en-GB', { dateStyle: 'medium', timeStyle: 'short' })} from 3CX call statistics</div></footer>
    </div></section>`;
  }

  function html(reports, period) {
    return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Call performance – ${esc(period)}</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap" rel="stylesheet"><style>
@page { size: A4 portrait; margin: 0; }
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: Inter, Arial, sans-serif; color: ${C.ink}; background: #e7e9e2; -webkit-print-color-adjust: exact; print-color-adjust: exact; font-feature-settings: 'tnum' 1; }
.bar { position: sticky; top: 0; z-index: 2; background: #121310; color: #fff; padding: 10px 16px; display: flex; gap: 12px; align-items: center; font-size: 13px; }
.bar button { background: ${C.lime}; color: #121400; border: 0; padding: 8px 18px; border-radius: 999px; font-weight: 600; cursor: pointer; font-family: inherit; }
.page { width: 210mm; height: 297mm; padding: 11mm 12mm; margin: 12px auto; background: #fff; overflow: hidden; box-shadow: 0 2px 12px rgba(0,0,0,.15); }
.inner { height: 100%; display: flex; flex-direction: column; gap: 3.2mm; font-size: 8.6pt; transform-origin: top left; }
header { display: flex; justify-content: space-between; align-items: center; }
.brand { display: flex; align-items: center; gap: 7px; font-weight: 600; font-size: 10.5pt; }
.logo { width: 26px; height: 26px; border-radius: 8px; background: ${C.lime}; display: inline-flex; align-items: center; justify-content: center; }
.doc { font-size: 8pt; color: ${C.muted}; }
h1 { font-size: 21pt; font-weight: 500; letter-spacing: -.5px; line-height: 1.1; }
.period { color: ${C.muted}; font-size: 8.8pt; margin-top: 1mm; }
.tiles { display: grid; grid-template-columns: repeat(6, 1fr); gap: 2mm; }
.tile { background: ${C.soft}; border-radius: 3mm; padding: 2.6mm 3mm; }
.tile.hl { background: ${C.lime}; }
.tl { font-size: 7.4pt; color: ${C.muted}; }
.tile.hl .tl, .tile.hl .vs { color: #3a3f1f; }
.tv { font-size: 15pt; font-weight: 500; letter-spacing: -.4px; margin: 1mm 0 .6mm; }
.tv small { font-size: 8.5pt; color: ${C.muted}; }
.vs { font-size: 7pt; color: ${C.muted}; display: block; }
.vs.up { color: #1f7a45; } .vs.dn { color: ${C.crit}; }
.box { border: 1px solid ${C.grid}; border-radius: 3mm; padding: 2.6mm 3mm 2mm; }
.bh { display: flex; justify-content: space-between; align-items: center; margin-bottom: 1.6mm; font-size: 9pt; }
.bh b { font-weight: 600; }
.lg { display: inline-flex; align-items: center; gap: 4px; font-size: 7.4pt; color: ${C.muted}; }
.lg i { display: inline-block; width: 9px; height: 9px; border-radius: 2px; margin-left: 6px; }
.lg i.dash { width: 14px; height: 0; border-top: 1.8px dashed ${C.team}; border-radius: 0; }
.two { display: grid; grid-template-columns: 1fr 1fr; gap: 3mm; }
table { width: 100%; border-collapse: collapse; font-size: 7.8pt; }
th { text-align: left; font-weight: 500; color: ${C.muted}; padding: .9mm 1mm; border-bottom: 1px solid ${C.ink}; }
td { padding: .8mm 1mm; border-bottom: 1px solid ${C.grid}; }
td.bad { color: ${C.crit}; font-weight: 600; }
.none { color: ${C.muted}; padding: 6mm 0; text-align: center; }
.notes { flex: 1; min-height: 14mm; border: 1px solid ${C.grid}; border-radius: 3mm; padding: 2.4mm 3mm; font-size: 8.4pt; }
footer { display: flex; flex-direction: column; gap: 2mm; }
.sign { display: flex; gap: 8mm; } .sign span { flex: 1; border-top: 1px solid ${C.ink}; padding-top: 1mm; font-size: 7.6pt; color: ${C.muted}; }
.gen { font-size: 7pt; color: ${C.muted}; }
@media print { body { background: #fff; } .bar { display: none; } .page { margin: 0; box-shadow: none; break-after: page; } .page:last-child { break-after: auto; } }
</style></head><body>
<div class="bar"><span><b>${reports.length}</b> report${reports.length === 1 ? '' : 's'} · one A4 page each</span><button onclick="window.print()">Print</button></div>
${reports.map((r, i) => page(r, i, reports.length, period)).join('')}
<script>
// If a page's content is taller than A4 (lots of agents in the ranking), scale it down to fit.
function fit(){ document.querySelectorAll('.page').forEach(function(p){ var inn=p.querySelector('.inner'); inn.style.transform=''; inn.style.width='';
  var cs=getComputedStyle(p), avail=p.clientHeight-parseFloat(cs.paddingTop)-parseFloat(cs.paddingBottom);
  inn.style.height='auto'; var need=inn.scrollHeight; inn.style.height='';
  if(need>avail){ var k=avail/need; inn.style.transform='scale('+k+')'; inn.style.width=(100/k)+'%'; inn.style.height=(avail/k)+'px'; } }); }
window.addEventListener('load', function(){ fit(); setTimeout(function(){ window.print(); }, 400); });
<\/script></body></html>`;
  }

  function print(agents, from, to) {
    const w = window.open('', '_blank');
    if (!w) return showToast('Your browser blocked the report window – allow pop-ups for ScoreDesk', false);
    const period = `${fmtDate(from)} – ${fmtDate(to)}`;
    const reports = [...agents].sort().map(a => compute(a, from, to)).filter(r => r.me.calls || r.me.ignored);
    if (!reports.length) { w.close(); return showToast('No calls for those agents in that period', false); }
    w.document.open(); w.document.write(html(reports, period)); w.document.close();
  }

  return { print, compute };
})();
