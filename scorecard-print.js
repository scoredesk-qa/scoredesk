// ScoreDesk – print scorecards, one A4 page each.
// Used by quality.html: SDPrint.dialog(cards) opens the picker, SDPrint.print(cards) prints.

const SDPrint = (() => {
  const { $, esc, fmtDate, addDays } = SD;

  // ------------------------------------------------------------ picker
  function dialog(cards, preset = {}) {
    const dated = cards.filter(c => c.date);
    if (!dated.length) return SD.toast('No dated scorecards to print', false);
    const latest = dated.map(c => c.date).sort().pop();
    const st = {
      from: preset.from || addDays(latest, -29), to: preset.to || latest,
      result: 'all', agents: new Set(preset.agents || [])
    };
    const all = [...new Set(dated.map(c => c.agent))].sort();
    if (!st.agents.size) all.forEach(a => st.agents.add(a));

    const root = document.createElement('div');
    document.body.appendChild(root);
    const close = () => root.remove();
    const pick = () => dated.filter(c => c.date >= st.from && c.date <= st.to && st.agents.has(c.agent) &&
      (st.result === 'all' || (st.result === 'fail') === !c.pass));

    function draw() {
      const inDates = dated.filter(c => c.date >= st.from && c.date <= st.to);
      const n = pick().length;
      root.innerHTML = `<div class="ov" id="prOv"><div class="modal">
        <div class="mh"><div><div class="eyebrow">Print</div><div class="t">Print scorecards</div>
          <div class="sub">Each scorecard prints on its own A4 page.</div></div><button class="xbtn" data-x>${SDTheme.icon('x', 16)}</button></div>
        <div class="mb2">
          <div class="facts" style="grid-template-columns:1fr 1fr 1fr">
            <div class="field"><label>From</label><input class="inp" type="date" id="prFrom" value="${st.from}"></div>
            <div class="field"><label>To</label><input class="inp" type="date" id="prTo" value="${st.to}"></div>
            <div class="field"><label>Result</label><select class="inp" id="prRes">
              <option value="all" ${st.result === 'all' ? 'selected' : ''}>Passed &amp; failed</option>
              <option value="fail" ${st.result === 'fail' ? 'selected' : ''}>Failed only</option>
              <option value="pass" ${st.result === 'pass' ? 'selected' : ''}>Passed only</option></select></div>
          </div>
          <div class="lbl"><span>Agents</span><span><a data-all>All</a> · <a data-none>None</a></span></div>
          <div class="alist" style="max-height:300px"><table class="dt" style="min-width:0"><tbody>
            ${all.map(a => { const k = inDates.filter(c => c.agent === a).length; return `<tr class="${st.agents.has(a) ? '' : 'off'}">
              <td style="width:30px"><input type="checkbox" data-a="${esc(a)}" ${st.agents.has(a) ? 'checked' : ''}></td>
              <td style="font-weight:700">${esc(a)}</td><td class="n cm" style="text-align:right">${k} scorecard${k === 1 ? '' : 's'}</td></tr>`; }).join('')}
          </tbody></table></div>
        </div>
        <div class="mf"><span class="grow">${fmtDate(st.from)} – ${fmtDate(st.to)} · ${st.agents.size} agent${st.agents.size === 1 ? '' : 's'}</span>
          <button class="btn" data-x>Cancel</button>
          <button class="btn primary" id="prGo" ${n ? '' : 'disabled'}>${SDTheme.icon('print', 15)}Print ${n} scorecard${n === 1 ? '' : 's'}</button></div>
      </div></div>`;
      root.querySelectorAll('[data-x]').forEach(b => b.onclick = close);
      $('prOv').onclick = e => { if (e.target.id === 'prOv') close(); };
      $('prFrom').onchange = e => { st.from = e.target.value || st.from; draw(); };
      $('prTo').onchange = e => { st.to = e.target.value || st.to; draw(); };
      $('prRes').onchange = e => { st.result = e.target.value; draw(); };
      root.querySelector('[data-all]').onclick = () => { all.forEach(a => st.agents.add(a)); draw(); };
      root.querySelector('[data-none]').onclick = () => { st.agents.clear(); draw(); };
      root.querySelectorAll('[data-a]').forEach(cb => cb.onchange = () => { cb.checked ? st.agents.add(cb.dataset.a) : st.agents.delete(cb.dataset.a); draw(); });
      $('prGo').onclick = () => { const list = pick(); close(); print(list, `${fmtDate(st.from)} – ${fmtDate(st.to)}`); };
    }
    draw();
  }

  // ------------------------------------------------------------ print window
  async function loadNotes(cards) {
    const keys = cards.map(noteKey), out = new Map();
    try {
      for (let i = 0; i < keys.length; i += 100) {
        const { data } = await SD.sb.from('notes').select('note_key,content').in('note_key', keys.slice(i, i + 100));
        (data || []).forEach(n => { if (n.content) out.set(n.note_key, n.content); });
      }
    } catch (e) {}
    return out;
  }
  const noteKey = c => 'call:' + encodeURIComponent(c.agent_name + '||' + c.file_name);

  async function print(cards, periodLabel) {
    if (!cards.length) return;
    // Open the window straight away (inside the click) so popup blockers allow it.
    const w = window.open('', '_blank');
    if (!w) return SD.toast('Your browser blocked the print window – allow pop-ups for ScoreDesk', false);
    w.document.write('<p style="font-family:sans-serif;padding:40px">Preparing scorecards…</p>');
    const list = [...cards].sort((a, b) => a.agent.localeCompare(b.agent) || (a.date + a.time).localeCompare(b.date + b.time));
    const notes = await loadNotes(list);
    w.document.open();
    w.document.write(pageHtml(list, notes, periodLabel));
    w.document.close();
  }

  // The same check can appear in the ZZPS summary and in the criteria: show it once,
  // as failed if either failed, else YES if either passed.
  function afChecks(c) {
    const m = new Map(), rank = a => a.triggered ? 3 : a.value === 'YES' ? 2 : 1;
    for (const a of c.afs) {
      const label = SD.canonicalAF(a.label), cur = m.get(label);
      if (!cur || rank(a) > rank(cur)) m.set(label, { ...a, label });
    }
    return [...m.values()];
  }

  const val = v => { const n = parseFloat(v); return isFinite(n) ? String(+n.toFixed(2)) : '—'; };

  function sheet(c, i, n, notes, periodLabel) {
    const sections = [];
    c.crit.forEach(cr => { const s = (cr.section || 'General').trim(); if (!sections.includes(s)) sections.push(s); });
    const trig = c.afs.filter(a => a.triggered);
    const note = notes.get(noteKey(c));
    const rows = sections.map(s => `<tr class="sec"><td colspan="3">${esc(s)}</td></tr>` + c.crit.filter(cr => (cr.section || 'General').trim() === s).map(cr => {
      const w = parseFloat(cr.weight), v = parseFloat(cr.score);
      const cls = !isFinite(v) ? 'na' : v >= w ? 'ok' : v > 0 ? 'part' : 'zero';
      return `<tr class="${cls}"><td class="q">${esc(cr.label)}</td><td class="s">${isFinite(v) ? val(v) : 'N/A'} / ${val(w)}</td><td class="f">${esc(cr.feedback || '')}</td></tr>`;
    }).join('')).join('');
    return `<section class="page"><div class="inner">
      <header>
        <div><div class="doc">Call Accreditation · ${esc(periodLabel || '')}</div><h1>${esc(c.agent)}</h1></div>
        <div class="res ${c.pass ? 'pass' : 'fail'}"><b>${c.pass ? 'PASS' : 'FAIL'}</b><span>${c.score == null ? '' : c.score.toFixed(0) + '%'}</span></div>
      </header>
      <table class="meta"><tr>
        <td><i>Call date</i>${fmtDate(c.date)}${c.time ? ' ' + esc(c.time) : ''}</td>
        <td><i>Account</i>${esc(c.account_ref || '—')}</td>
        <td><i>Auditor</i>${esc(c.auditor || '—')}</td>
        <td><i>Pass mark</i>90%</td></tr>
        <tr><td colspan="4"><i>Recording</i>${esc(c.file_name || '—')}</td></tr></table>
      ${!c.pass && c.pf_reason ? `<div class="reason"><b>Reason:</b> ${esc(c.pf_reason)}</div>` : ''}
      <div class="afs">${afChecks(c).map(a => `<span class="af ${a.triggered ? 'bad' : a.value === 'YES' ? 'good' : ''}">${esc(a.label)}: <b>${esc(a.value || 'N/A')}</b></span>`).join('')}</div>
      ${trig.filter(a => a.feedback).map(a => `<div class="reason"><b>Auto-fail – ${esc(SD.canonicalAF(a.label))}:</b> ${esc(a.feedback)}</div>`).join('')}
      <table class="crit"><thead><tr><th>Criterion</th><th>Score</th><th>Feedback</th></tr></thead><tbody>${rows}</tbody></table>
      ${note ? `<div class="notes"><b>Notes:</b> ${esc(note)}</div>` : ''}
      <footer><div class="sign"><span>Agent signature</span><span>Auditor signature</span><span>Date</span></div>
        <div class="pg">${i + 1} / ${n}</div></footer>
    </div></section>`;
  }

  function pageHtml(list, notes, periodLabel) {
    return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Scorecards – ${esc(periodLabel || '')}</title><style>
      @page { size: A4 portrait; margin: 0; }
      * { box-sizing: border-box; margin: 0; padding: 0; }
      body { font-family: Arial, Helvetica, sans-serif; color: #111; background: #e5e7eb; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      .bar { position: sticky; top: 0; background: #121310; color: #fff; padding: 10px 16px; display: flex; gap: 12px; align-items: center; font-size: 13px; z-index: 2; }
      .bar button { background: #e5ff8f; color: #121400; border: 0; padding: 8px 18px; border-radius: 999px; font-weight: 600; cursor: pointer; }
      .page { width: 210mm; height: 297mm; padding: 10mm 11mm; margin: 12px auto; background: #fff; overflow: hidden; position: relative; box-shadow: 0 2px 10px rgba(0,0,0,.15); }
      .inner { font-size: var(--fs, 9.5pt); line-height: 1.3; display: flex; flex-direction: column; min-height: 100%; }
      .page.overflow::after { content: 'Too long for one page – shortened'; position: absolute; bottom: 3mm; left: 11mm; font-size: 7pt; color: #b91c1c; }
      header { display: flex; justify-content: space-between; align-items: flex-start; border-bottom: 2px solid #111; padding-bottom: 2.5mm; margin-bottom: 2.5mm; }
      .doc { font-size: .8em; color: #555; text-transform: uppercase; letter-spacing: .06em; }
      h1 { font-size: 1.9em; margin-top: 1mm; }
      .res { border: 2px solid; border-radius: 2mm; padding: 1.5mm 4mm; text-align: center; min-width: 26mm; }
      .res b { display: block; font-size: 1.6em; } .res span { font-size: 1.1em; font-weight: 700; }
      .res.pass { color: #047857; border-color: #047857; background: #ecfdf5; } .res.fail { color: #b91c1c; border-color: #b91c1c; background: #fef2f2; }
      table { width: 100%; border-collapse: collapse; }
      .meta td { border: 1px solid #d1d5db; padding: 1.2mm 2mm; font-size: .92em; vertical-align: top; }
      .meta i { display: block; font-style: normal; font-size: .78em; color: #6b7280; text-transform: uppercase; letter-spacing: .04em; }
      .reason { margin-top: 2mm; padding: 1.5mm 2.5mm; background: #fef2f2; border-left: 3px solid #b91c1c; font-size: .92em; }
      .afs { display: flex; flex-wrap: wrap; gap: 1.5mm; margin-top: 2mm; }
      .af { border: 1px solid #d1d5db; border-radius: 1.5mm; padding: .6mm 2mm; font-size: .82em; }
      .af.good b { color: #047857; } .af.bad { border-color: #b91c1c; background: #fef2f2; } .af.bad b { color: #b91c1c; }
      .crit { margin-top: 2.5mm; }
      .crit th { text-align: left; font-size: .78em; text-transform: uppercase; letter-spacing: .04em; color: #6b7280; border-bottom: 1.5px solid #111; padding: 1mm 1.5mm; }
      .crit td { border-bottom: 1px solid #e5e7eb; padding: 1mm 1.5mm; vertical-align: top; }
      .crit .sec td { background: #f3f4f6; font-weight: 700; font-size: .85em; text-transform: uppercase; letter-spacing: .04em; padding-top: 1.3mm; }
      .crit .q { width: 34%; } .crit .s { width: 11%; white-space: nowrap; font-weight: 700; text-align: center; } .crit .f { color: #374151; font-size: .9em; }
      .crit .zero .s { color: #b91c1c; } .crit .zero .q { font-weight: 700; } .crit .part .s { color: #b45309; } .crit .ok .s { color: #047857; } .crit .na .s { color: #9ca3af; font-weight: 400; }
      .notes { margin-top: 2.5mm; padding: 1.5mm 2.5mm; border: 1px solid #d1d5db; border-radius: 1.5mm; font-size: .92em; white-space: pre-wrap; }
      footer { margin-top: auto; padding-top: 4mm; display: flex; justify-content: space-between; align-items: flex-end; gap: 6mm; }
      .sign { display: flex; gap: 8mm; flex: 1; } .sign span { flex: 1; border-top: 1px solid #111; padding-top: 1mm; font-size: .8em; color: #555; }
      .pg { font-size: .8em; color: #6b7280; }
      @media print { body { background: #fff; } .bar { display: none; } .page { margin: 0; box-shadow: none; break-after: page; page-break-after: always; } .page:last-child { break-after: auto; page-break-after: auto; } }
    </style></head><body>
      <div class="bar"><span><b>${list.length}</b> scorecard${list.length === 1 ? '' : 's'} · one A4 page each</span><button onclick="window.print()">Print</button></div>
      ${list.map((c, i) => sheet(c, i, list.length, notes, periodLabel)).join('')}
      <script>
        // Shrink the text of any scorecard that would spill onto a second page.
        function fit() {
          document.querySelectorAll('.page').forEach(function (p) {
            var cs = getComputedStyle(p), inner = p.querySelector('.inner'), fs = 9.5;
            var avail = p.clientHeight - parseFloat(cs.paddingTop) - parseFloat(cs.paddingBottom);
            inner.style.setProperty('--fs', fs + 'pt');
            while (inner.scrollHeight > avail + 0.5 && fs > 5) { fs -= 0.25; inner.style.setProperty('--fs', fs + 'pt'); }
            if (inner.scrollHeight > avail + 0.5) p.classList.add('overflow');
          });
        }
        window.addEventListener('load', function () { fit(); setTimeout(function () { window.print(); }, 300); });
      <\/script>
    </body></html>`;
  }

  return { dialog, print };
})();
