// ScoreDesk – Call Stats: download an Excel report for chosen agents and dates.
// Uses globals from agent-progression.html: ROWS, VIEW, XLSX, sumRows, groupBy,
// mondayOf, fmtDate, esc, showToast.

const CSExport = (() => {
  const el = id => document.getElementById(id);
  const ic = (n, s) => SDTheme.icon(n, s || 16);

  // ------------------------------------------------------------ dialog
  function open(preset = {}) {
    if (!ROWS.length) return showToast('No call data loaded yet', false);
    const dates = ROWS.map(r => r.date).sort();
    const all = [...new Set(ROWS.map(r => r.agent))].sort();
    const st = {
      from: preset.from || (VIEW && VIEW.from) || dates[0],
      to: preset.to || (VIEW && VIEW.to) || dates[dates.length - 1],
      agents: new Set(preset.agents && preset.agents.length ? preset.agents : all),
      sheets: new Set(['daily', 'weekly', 'monthly'])
    };
    const root = el('modalRoot');
    const close = () => { root.innerHTML = ''; };

    function draw() {
      const inRange = ROWS.filter(r => r.date >= st.from && r.date <= st.to);
      const count = a => inRange.filter(r => r.agent === a && r.calls > 0).length;
      const n = st.agents.size;
      root.innerHTML = `<div class="ov" id="exOv"><div class="modal">
        <div class="mh"><div><div class="eyebrow">Download</div><div class="t">Call stats report</div>
          <div class="sub">An Excel file with a summary plus daily, weekly and monthly figures.</div></div>
          <button class="xbtn" data-x>${ic('x')}</button></div>
        <div class="mb2">
          <div class="facts" style="grid-template-columns:1fr 1fr 2fr">
            <div class="field" style="margin:0"><label>From</label><input class="inp" type="date" id="exFrom" min="${dates[0]}" max="${dates[dates.length - 1]}" value="${st.from}"></div>
            <div class="field" style="margin:0"><label>To</label><input class="inp" type="date" id="exTo" min="${dates[0]}" max="${dates[dates.length - 1]}" value="${st.to}"></div>
            <div class="field" style="margin:0"><label>Include</label><div class="qchips" style="margin:0">
              ${[['daily', 'Daily'], ['weekly', 'Weekly'], ['monthly', 'Monthly']].map(([k, l]) =>
                `<label class="qchip ${st.sheets.has(k) ? 'on' : ''}"><input type="checkbox" data-sheet="${k}" ${st.sheets.has(k) ? 'checked' : ''}>${l}</label>`).join('')}
            </div></div>
          </div>
          <div class="lbl"><span>Agents</span><span><a data-all>All</a> · <a data-none>None</a></span></div>
          <div class="alist" style="max-height:320px"><table class="dt" style="min-width:0"><tbody>
            ${all.map(a => `<tr class="${st.agents.has(a) ? '' : 'off'}">
              <td style="width:34px"><input type="checkbox" data-a="${esc(a)}" ${st.agents.has(a) ? 'checked' : ''}></td>
              <td style="font-weight:500">${esc(a)}</td><td class="cm" style="text-align:right">${count(a)} days worked</td></tr>`).join('')}
          </tbody></table></div>
        </div>
        <div class="mf"><span class="grow">${fmtDate(st.from)} – ${fmtDate(st.to)} · ${n} agent${n === 1 ? '' : 's'}</span>
          <button class="btn" data-x>Cancel</button>
          <button class="btn primary" id="exGo" ${n && st.from <= st.to ? '' : 'disabled'}>${ic('download', 15)}Download Excel</button></div>
      </div></div>`;
      root.querySelectorAll('[data-x]').forEach(b => b.onclick = close);
      el('exOv').onclick = e => { if (e.target.id === 'exOv') close(); };
      el('exFrom').onchange = e => { if (e.target.value) st.from = e.target.value; draw(); };
      el('exTo').onchange = e => { if (e.target.value) st.to = e.target.value; draw(); };
      root.querySelectorAll('[data-sheet]').forEach(cb => cb.onchange = () => { cb.checked ? st.sheets.add(cb.dataset.sheet) : st.sheets.delete(cb.dataset.sheet); draw(); });
      root.querySelector('[data-all]').onclick = () => { all.forEach(a => st.agents.add(a)); draw(); };
      root.querySelector('[data-none]').onclick = () => { st.agents.clear(); draw(); };
      root.querySelectorAll('[data-a]').forEach(cb => cb.onchange = () => { cb.checked ? st.agents.add(cb.dataset.a) : st.agents.delete(cb.dataset.a); draw(); });
      el('exGo').onclick = () => {
        try { const name = build(st); close(); showToast('Downloaded ' + name, true); }
        catch (e) { console.error(e); showToast('Could not build the file: ' + e.message, false); }
      };
    }
    draw();
  }

  // ------------------------------------------------------------ workbook
  const DAY = 1440; // minutes per day: Excel stores durations as fractions of a day
  const pct = v => v == null ? null : v / 100;
  const row = s => [s.days, s.calls, s.ignored, s.calls + s.ignored, pct(s.handled), pct(s.ignRate),
    s.days ? +(s.calls / s.days).toFixed(1) : null, s.talkMin ? s.talkMin / DAY : null, s.avgTalk != null ? s.avgTalk / DAY : null];
  const HEAD = ['Days worked', 'Answered', 'Ignored', 'Offered', 'Handled %', 'Ignore %', 'Calls / day', 'Talk time', 'Avg talk'];
  // number formats for the HEAD columns (after any leading text columns)
  const FMT = [null, '#,##0', '#,##0', '#,##0', '0.0%', '0.0%', '0.0', '[h]:mm', 'm:ss'];

  function sheet(title, lead, rows, sub) {
    const aoa = [[title], [sub], [], [...lead, ...HEAD], ...rows];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const off = lead.length;
    for (let r = 4; r < aoa.length; r++) FMT.forEach((f, i) => {
      const cell = ws[XLSX.utils.encode_cell({ r, c: off + i })];
      if (f && cell && typeof cell.v === 'number') cell.z = f;
    });
    ws['!cols'] = [...lead.map(h => ({ wch: Math.max(14, h.length + 2) })), ...HEAD.map(h => ({ wch: Math.max(11, h.length + 2) }))];
    ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 3, c: 0 }, e: { r: Math.max(3, aoa.length - 1), c: off + HEAD.length - 1 } }) };
    return ws;
  }

  function build(st) {
    const rows = ROWS.filter(r => r.date >= st.from && r.date <= st.to && st.agents.has(r.agent));
    if (!rows.some(r => r.calls || r.ignored)) throw new Error('no calls for those agents in that period');
    const agents = [...st.agents].sort();
    const sub = `${fmtDate(st.from)} to ${fmtDate(st.to)} · generated ${new Date().toLocaleString('en-GB', { dateStyle: 'medium', timeStyle: 'short' })} · ScoreDesk`;
    const wb = XLSX.utils.book_new();

    // Summary: one line per agent + team total
    const byAgent = groupBy(rows, r => r.agent);
    const sum = agents.map(a => [a, ...row(sumRows(byAgent.get(a) || []))]);
    if (agents.length > 1) sum.push(['Total', ...row(sumRows(rows))]);
    XLSX.utils.book_append_sheet(wb, sheet('Call stats – summary', ['Agent'], sum, sub), 'Summary');

    const per = (keyOf, label) => {
      const out = [];
      for (const a of agents) {
        const g = groupBy((byAgent.get(a) || []), r => keyOf(r.date));
        [...g.keys()].sort().forEach(k => { const s = sumRows(g.get(k)); if (s.calls || s.ignored) out.push([...label(k), a, ...row(s)]); });
      }
      return out;
    };
    if (st.sheets.has('daily')) {
      const d = per(x => x, k => [k, new Date(k + 'T00:00:00Z').toLocaleDateString('en-GB', { weekday: 'short', timeZone: 'UTC' })]);
      XLSX.utils.book_append_sheet(wb, sheet('Call stats – daily', ['Date', 'Day', 'Agent'], d, sub), 'Daily');
    }
    if (st.sheets.has('weekly')) XLSX.utils.book_append_sheet(wb, sheet('Call stats – weekly', ['Week commencing', 'Agent'], per(mondayOf, k => [k]), sub), 'Weekly');
    if (st.sheets.has('monthly')) XLSX.utils.book_append_sheet(wb, sheet('Call stats – monthly', ['Month', 'Agent'], per(x => x.slice(0, 7), k => [k]), sub), 'Monthly');

    const who = agents.length === 1 ? agents[0].replace(/[^\w]+/g, '_') : `${agents.length}_agents`;
    const name = `CallStats_${who}_${st.from}_to_${st.to}.xlsx`;
    XLSX.writeFile(wb, name, { compression: true });
    return name;
  }

  return { open, build };
})();
