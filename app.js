/* Ekonomi – appens gränssnitt */
window.EKONOMI_VERSION = '1.2.0';
(function () {
  'use strict';
  const $ = (s, el = document) => el.querySelector(s);
  const h = (tag, attrs = {}, ...kids) => {
    const el = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (v === null || v === undefined || v === false) continue;
      if (k === 'class') el.className = v; else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
      else if (k === 'html') el.innerHTML = v; else el.setAttribute(k, v === true ? '' : v);
    }
    for (const k of kids.flat()) if (k !== null && k !== undefined && k !== false) el.append(k.nodeType ? k : document.createTextNode(k));
    return el;
  };
  const nf0 = new Intl.NumberFormat('sv-SE', { maximumFractionDigits: 0 });
  const nf1 = new Intl.NumberFormat('sv-SE', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  const nf2 = new Intl.NumberFormat('sv-SE', { maximumFractionDigits: 2 });
  const kr = x => (x === null || x === undefined || !isFinite(x)) ? '–' : nf0.format(Math.round(x)) + ' kr';
  const krS = x => (x > 0 ? '+' : '') + kr(x);
  const pct = (x, sign = true) => (x === null || x === undefined || !isFinite(x)) ? '–' : (sign && x > 0 ? '+' : '') + nf1.format(x * 100) + ' %';
  const cls = x => !isFinite(x) || x === null ? '' : x > 0 ? 'pos' : x < 0 ? 'neg' : '';
  const MON = ['jan', 'feb', 'mar', 'apr', 'maj', 'jun', 'jul', 'aug', 'sep', 'okt', 'nov', 'dec'];
  const dstr = d => `${+d.slice(8, 10)} ${MON[+d.slice(5, 7) - 1]} ${d.slice(0, 4)}`;
  const today = () => { const d = new Date(); return new Date(d - d.getTimezoneOffset() * 60000).toISOString().slice(0, 10); };
  const mkr = x => Math.abs(x) >= 1e6 ? nf2.format(x / 1e6) + ' Mkr' : Math.abs(x) >= 1e3 ? nf0.format(x / 1e3) + ' tkr' : nf0.format(x);

  // "1 234,50", "1064+2242-300", "−500" → tal
  function parseNum(s) {
    s = String(s ?? '').replace(/[\s ]/g, '').replace(/kr/gi, '').replace(/−/g, '-').replace(/,/g, '.');
    if (s === '') return 0;
    if (!/^[-+]?\d*\.?\d+([-+]\d*\.?\d+)*$/.test(s)) return NaN;
    return Math.round((s.match(/[-+]?\d*\.?\d+/g) || []).reduce((a, b) => a + parseFloat(b), 0) * 100) / 100;
  }

  const bankNamn = () => (S.data && S.data.settings && S.data.settings.bank) || 'Bank';
  const S = { data: null, etag: null, out: [], busy: false, range: '3', lastErr: null, syncedAt: null, offline: false };

  function toast(msg, kind) {
    const t = h('div', { class: 'toast ' + (kind || '') }, msg);
    $('#toasts').append(t); setTimeout(() => t.classList.add('show'), 10);
    setTimeout(() => { t.classList.remove('show'); setTimeout(() => t.remove(), 400); }, kind === 'err' ? 6000 : 3000);
  }
  function recompute() { S.out = S.data && S.data.rows.length ? Calc.compute(S.data.rows) : []; }

  // ---------------------------------------------------------------- spara
  async function commit(mutator, okMsg) {
    if (S.busy) return false;
    const next = JSON.parse(JSON.stringify(S.data));
    mutator(next);
    next.rows = Calc.sortRows(next.rows); next.updated = new Date().toISOString();
    S.busy = true; render();
    try {
      S.etag = await Store.save(next, S.etag);
      S.data = next; recompute(); S.syncedAt = Date.now();
      if (okMsg) toast(okMsg, 'ok');
      backup(false);
      return true;
    } catch (e) {
      toast(e.message, 'err');
      if (e.conflict) await reload();
      if (e.needLogin) render();
      return false;
    } finally { S.busy = false; render(); }
  }
  async function reload() {
    const r = await Store.load(); S.data = r.data; S.etag = r.etag; S.syncedAt = Date.now();
    if (S.data && Calc.migrate(S.data)) { recompute(); await commit(() => { }, null); }
    recompute();
  }

  // ---------------------------------------------------------------- Excel-backup
  function workbook() {
    const rows = Calc.sortRows(S.data.rows), out = S.out;
    const P = v => ({ v, f: 'pct' }), K = v => ({ v, f: 'kr' });
    const B = bankNamn();
    const head = ['Datum', 'Kassa', 'Avanza insatt aktier', 'Avanza insatt fonder', 'Avanza insatt ETF', 'Avanza värde aktier', 'Avanza värde fonder',
      'Avanza värde ETF', 'Utdelning', 'Avanza realiserad vinst', B + ' insatt', B + ' värde', B + ' realiserad vinst', 'SAVR insatt (egna pengar)', 'SAVR värde',
      'SAVR återbäring', 'Nettolön', 'Totalt värde', 'Egna insatta pengar', 'Avkastat totalt', 'Avkastning veckan', 'Utveckling veckan',
      'Utveckling i år (tidsviktad)', 'Utveckling i år (pengaviktad)', 'Avkastat i år', 'Totalt på egna pengar', 'Förmögenhet', 'Sparat i år',
      'Avanza i år', B + ' i år', 'SAVR i år', 'Sparat sedan förra lönen', 'Sparkvot (månad)', 'Snittspar per månad i år'];
    const wk = [head].concat(rows.map((r, i) => { const o = out[i], a = r.av || {}, d = r.db || {}, s = r.sv || {};
      return [{ date: r.d }, K(r.kassa), K(a.ai), K(a.fi), K(a.ei), K(a.av), K(a.fv), K(a.ev), K(a.utd), K(a.vinst), K(d.inv), K(d.val), K(d.vinst),
        K(s.inv), K(s.val), K(s.ater), K(r.lon), K(o.N), K(o.K), K(o.V), K(o.W), P(o.X), P(o.Y), P(o.O), K(o.U), P(o.P), K(o.formogenhet),
        K(o.sparatAr), P(o.avY), P(o.dbY), P(o.svY), K(o.manSparat), P(o.manKvot), K(o.lon !== null ? o.snittMan : null)]; }));
    const yrs = Calc.years(out);
    const yr = [['År', 'Utveckling (tidsviktad)', 'Utveckling (pengaviktad)', 'Avkastat kr', 'Sparat', 'Utdelningar', 'Nettolön', 'Portföljvärde', 'Förmögenhet',
      'Avanza', B, 'SAVR', 'Sparkvot', 'Snittspar per månad']].concat(yrs.map(o => [o.year, P(o.Y), P(o.O), K(o.U), K(o.sparatAr), K(o.utdAr), K(o.lonAr), K(o.N), K(o.formogenhet), P(o.avY), P(o.dbY), P(o.svY), P(o.sparkvotAr), K(o.snittMan)]));
    const info = [['Ekonomi – backup'], ['Skapad: ' + new Date().toLocaleString('sv-SE')],
      ['Bladet "Veckor" innehåller all rådata (kolumn A–Q) och de uträknade värdena.'],
      ['Rådatan räcker för att återskapa allt. Originalfilen Joel_Ekonomi.xlsm innehåller historiken med formler.']];
    return XLSX.build([{ name: 'Veckor', rows: wk, widths: [11].concat(Array(33).fill(14)) }, { name: 'År', rows: yr, widths: [8].concat(Array(13).fill(16)) },
      ...(AV.an ? avSheets() : []), { name: 'Info', rows: info, widths: [100] }]);
  }
  function avSheets() {
    const K = v => ({ v, f: 'kr' }), P = v => ({ v, f: 'pct' }), an = AV.an;
    const hold = [['Namn', 'Typ', 'Antal', 'GAV', 'Senaste kurs', 'Värde', 'Anskaffningsvärde', 'Sedan köp', 'Sedan köp %', 'Utdelningar', 'Köpt totalt', 'Sålt totalt', 'Total avkastning']]
      .concat(an.current.map(c => [c.name, c.typ || '', c.qty, c.gav, c.senast, K(c.varde), K(c.ansk), K(c.sedanKop), P(c.sedanKopPct), K(c.utd), K(c.kopt), K(c.salt), P(c.total)]));
    const cl = [['Namn', 'Sålt', 'Köpt för', 'Sålt för', 'Resultat', 'Utdelningar', 'Total avkastning']].concat(an.closed.map(c => [c.name, { date: c.last }, K(c.kopt), K(c.salt), K(c.resultat), K(c.utd), P(c.total)]));
    const yr = [['År', 'Utdelningar', 'Källskatt', 'Realiserat resultat', 'Skatt och avgifter']].concat(an.years.map(y => [y.year, K(y.utd), K(y.kallskatt), K(y.realiserat), K(y.skatt)]));
    return [{ name: 'Innehav', rows: hold, widths: [34].concat(Array(12).fill(14)) }, { name: 'Avslutade', rows: cl, widths: [34, 11].concat(Array(5).fill(14)) }, { name: 'Utdelningar', rows: yr, widths: [8, 14, 14, 18, 18] }];
  }
  async function backup(manual) {
    if (!S.data || !S.data.rows.length) return;
    try {
      const blob = workbook();
      if (Store.isDemo()) { if (manual) download(blob); return; }
      await Store.saveFile('Ekonomi-backup.xlsx', blob);
      await Store.saveFile(`Ekonomi-backup-${today().slice(0, 7)}.xlsx`, blob);
      localStorage.setItem('ek.lastBackup', new Date().toISOString());
      if (manual) toast('Excel-backup sparad i OneDrive', 'ok');
      if (location.hash.startsWith('#mer')) render();
    } catch (e) { if (manual) toast('Backup misslyckades: ' + e.message, 'err'); }
  }
  function download(blob, name) {
    const a = h('a', { href: URL.createObjectURL(blob), download: name || `Ekonomi-backup-${today()}.xlsx` });
    document.body.append(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }

  // ---------------------------------------------------------------- grafer (SVG)
  function lineChart(series, dates, opts = {}) {
    const W = chartW(), H = 240, L = 58, R = 10, T = 12, B = 26;
    const all = series.flatMap(s => s.values).filter(isFinite);
    let lo = Math.min(0, ...all), hi = Math.max(...all); if (hi === lo) hi = lo + 1;
    const step = niceStep((hi - lo) / 4); lo = Math.floor(lo / step) * step; hi = Math.ceil(hi / step) * step;
    const x = i => L + (W - L - R) * (dates.length > 1 ? i / (dates.length - 1) : 0.5);
    const y = v => T + (H - T - B) * (1 - (v - lo) / (hi - lo));
    let g = '';
    for (let v = lo; v <= hi + step / 2; v += step) g += `<line x1="${L}" x2="${W - R}" y1="${y(v)}" y2="${y(v)}" class="grid"/><text x="${L - 6}" y="${y(v) + 4}" class="ax" text-anchor="end">${opts.pct ? nf0.format(v * 100) + ' %' : mkr(v)}</text>`;
    let lastY = null;
    dates.forEach((d, i) => { const yy = d.slice(0, 4); if (yy !== lastY && (i > 0 || dates.length < 60)) { if (i > 0) g += `<line x1="${x(i)}" x2="${x(i)}" y1="${T}" y2="${H - B}" class="grid v"/>`; g += `<text x="${x(i) + 3}" y="${H - 8}" class="ax">${yy}</text>`; } lastY = yy; });
    series.forEach(s => {
      const pts = s.values.map((v, i) => isFinite(v) ? `${x(i).toFixed(1)},${y(v).toFixed(1)}` : null).filter(Boolean);
      if (s.fill) g += `<polygon points="${x(0)},${y(Math.max(lo, 0))} ${pts.join(' ')} ${x(dates.length - 1)},${y(Math.max(lo, 0))}" class="fill ${s.cls}"/>`;
      g += `<polyline points="${pts.join(' ')}" class="line ${s.cls}"/>`;
    });
    const svg = h('div', { class: 'chart', html: `<svg viewBox="0 0 ${W} ${H}" role="img">${g}<line class="cursor" x1="0" x2="0" y1="${T}" y2="${H - B}" style="display:none"/></svg>` });
    const tip = h('div', { class: 'tip' }); svg.append(tip);
    const move = ev => {
      const rect = svg.getBoundingClientRect(); const px = ((ev.touches ? ev.touches[0].clientX : ev.clientX) - rect.left) / rect.width * W;
      const i = Math.max(0, Math.min(dates.length - 1, Math.round((px - L) / (W - L - R) * (dates.length - 1))));
      const c = $('.cursor', svg); c.style.display = ''; c.setAttribute('x1', x(i)); c.setAttribute('x2', x(i));
      tip.innerHTML = `<b>${dstr(dates[i])}</b>` + series.map(s => `<div><i class="dot ${s.cls}"></i>${s.name}: ${opts.pct ? pct(s.values[i]) : kr(s.values[i])}</div>`).join('');
      tip.style.display = 'block'; tip.style.left = Math.min(rect.width - 170, Math.max(0, x(i) / W * rect.width - 80)) + 'px';
    };
    svg.addEventListener('pointermove', move); svg.addEventListener('touchmove', move, { passive: true });
    svg.addEventListener('pointerleave', () => { tip.style.display = 'none'; $('.cursor', svg).style.display = 'none'; });
    return svg;
  }
  function chartW() { const m = document.getElementById('main'); return Math.max(300, Math.min(736, (m ? m.clientWidth : 640) - 26)); }
  function niceStep(r) { const p = Math.pow(10, Math.floor(Math.log10(r))); const m = r / p; return (m <= 1 ? 1 : m <= 2 ? 2 : m <= 2.5 ? 2.5 : m <= 5 ? 5 : 10) * p; }
  function barChart(labels, values) {
    const W = chartW(), H = 200, L = 44, R = 8, T = 18, B = 24;
    let lo = Math.min(0, ...values), hi = Math.max(0, ...values); if (hi === lo) hi = 1;
    const step = niceStep((hi - lo) / 4); lo = Math.floor(lo / step) * step; hi = Math.ceil(hi / step) * step;
    const y = v => T + (H - T - B) * (1 - (v - lo) / (hi - lo)); const bw = (W - L - R) / labels.length;
    let g = '';
    for (let v = lo; v <= hi + step / 2; v += step) g += `<line x1="${L}" x2="${W - R}" y1="${y(v)}" y2="${y(v)}" class="grid"/><text x="${L - 6}" y="${y(v) + 4}" class="ax" text-anchor="end">${nf0.format(v * 100)} %</text>`;
    labels.forEach((lab, i) => { const v = values[i], x0 = L + i * bw + bw * 0.18, w = bw * 0.64;
      g += `<rect x="${x0}" y="${Math.min(y(v), y(0))}" width="${w}" height="${Math.abs(y(v) - y(0))}" class="bar ${v >= 0 ? 'pos' : 'neg'}" rx="3"/>` +
        `<text x="${x0 + w / 2}" y="${v >= 0 ? y(v) - 4 : y(0) + 13}" class="ax val${v < 0 ? ' inbar' : ''}" text-anchor="middle">${nf1.format(v * 100)}</text>` +
        `<text x="${x0 + w / 2}" y="${H - 7}" class="ax" text-anchor="middle">${lab}</text>`; });
    return h('div', { class: 'chart', html: `<svg viewBox="0 0 ${W} ${H}">${g}</svg>` });
  }

  // ---------------------------------------------------------------- vyer
  function viewOverview() {
    const out = S.out, o = out[out.length - 1];
    const yrs = Calc.years(out), prevY = yrs.length > 1 ? yrs[yrs.length - 2] : null;
    const mal = (S.data.settings && S.data.settings.mal) || 0;
    const stat = (label, main, sub, c) => h('div', { class: 'stat' }, h('div', { class: 'lbl' }, label), h('div', { class: 'big ' + (c || '') }, main), sub ? h('div', { class: 'sub' }, sub) : null);
    const n = out.length;
    const from = S.range === 'all' ? 0 : Math.max(0, out.findIndex(x => x.d >= shiftYears(o.d, -(+S.range))));
    const sl = out.slice(from);
    return h('div', { class: 'view' },
      paminnelser(),
      h('section', { class: 'hero' },
        h('div', { class: 'lbl' }, 'Total förmögenhet · ' + dstr(o.d)),
        h('div', { class: 'huge' }, kr(o.formogenhet)),
        h('div', { class: 'sub' }, h('span', { class: cls(o.formAr) }, krS(o.formAr)), ' i år'),
        mal ? h('div', { class: 'goal' }, h('div', { class: 'bar' }, h('i', { style: `width:${Math.min(100, o.formogenhet / mal * 100)}%` })),
          h('div', { class: 'sub' }, `${nf1.format(o.formogenhet / mal * 100)} % av målet ${mkr(mal)}`)) : null),
      h('section', { class: 'card' }, h('h2', {}, 'Avkastning'),
        h('div', { class: 'grid3' },
          stat(periodLabel(n - 1), pct(o.X), krS(o.W), cls(o.X)),
          stat('I år', pct(o.Y), krS(o.U), cls(o.Y)),
          stat('Totalt', pct(o.P), krS(o.V), cls(o.P))),
        h('p', { class: 'note' }, `I år på insatt kapital (pengaviktat): ${pct(o.O)}. "Totalt" är avkastat belopp i förhållande till egna insatta pengar.`)),
      h('section', { class: 'card' }, h('h2', {}, 'Konton'),
        h('table', { class: 'tbl' }, h('thead', {}, h('tr', {}, h('th', {}, ''), h('th', {}, 'Värde kr'), h('th', {}, 'I år'), h('th', {}, 'Totalt'))),
          h('tbody', {},
            [['Avanza', o.Aval, o.avY, o.avTot], ['SAVR', o.Sval, o.svY, o.svTot], [bankNamn(), o.Dval, o.dbY, o.dbTot], ['Kassa', o.kassa, null, null]]
              .map(([nm, v, y, t]) => h('tr', {}, h('td', {}, nm), h('td', {}, nf0.format(v)), h('td', { class: cls(y) }, y === null ? '' : pct(y)), h('td', { class: cls(t) }, t === null ? '' : pct(t))))),
          h('tfoot', {}, h('tr', {}, h('td', {}, 'Avanza: aktier / fonder / ETF i år'), h('td', { colspan: 3, class: 'small' }, `${pct(o.aktY)} / ${pct(o.fondY)} / ${pct(o.etfY)}`))))),
      sparandeKort(o, prevY),
      h('section', { class: 'card' }, h('div', { class: 'row' }, h('h2', {}, 'Portföljvärde och insatta pengar'),
        h('div', { class: 'seg' }, [['1', '1 år'], ['3', '3 år'], ['all', 'Allt']].map(([k, l]) => h('button', { class: S.range === k ? 'on' : '', onclick: () => { S.range = k; render(); } }, l)))),
        lineChart([{ name: 'Portföljvärde', values: sl.map(x => x.N), cls: 'c1', fill: true }, { name: 'Egna insatta pengar', values: sl.map(x => x.K), cls: 'c2' }], sl.map(x => x.d)),
        h('div', { class: 'legend' }, h('span', {}, h('i', { class: 'dot c1' }), 'Portföljvärde'), h('span', {}, h('i', { class: 'dot c2' }), 'Egna insatta pengar'))),
      h('section', { class: 'card' }, h('h2', {}, 'Utveckling per år'),
        barChart(yrs.map(y => y.year + (y === yrs[yrs.length - 1] && o.d.slice(5) < '12-28' ? '*' : '')), yrs.map(y => y.Y)),
        h('table', { class: 'tbl' }, h('thead', {}, h('tr', {}, ['År', 'Utv.', 'Avkastat kr', 'Sparat kr', 'Utdeln. kr'].map(x => h('th', {}, x)))),
          h('tbody', {}, yrs.slice().reverse().map(y => h('tr', {}, h('td', {}, y.year), h('td', { class: cls(y.Y) }, pct(y.Y)), h('td', {}, nf0.format(y.U)), h('td', {}, nf0.format(y.sparatAr)), h('td', {}, nf0.format(y.utdAr)))))),
        h('p', { class: 'note' }, '* pågående år. Utvecklingen är tidsviktad – den påverkas inte av när du sätter in pengar.')),
      n < 2 ? h('p', { class: 'note' }, 'Lägg in fler veckor för att se utveckling.') : null);
  }
  function sparandeKort(o, prevY) {
    const stat = (label, main, sub, c) => h('div', { class: 'stat' }, h('div', { class: 'lbl' }, label), h('div', { class: 'big ' + (c || '') }, main), sub ? h('div', { class: 'sub' }, sub) : null);
    const lonRader = S.out.filter(x => x.year === o.year && x.lon !== null).reverse();
    return h('section', { class: 'card' }, h('h2', {}, 'Sparande ' + o.year),
      h('div', { class: 'grid2' },
        stat('Sparat i år', kr(o.sparatAr), o.lonAntal ? `${o.lonAntal} löner` : ''),
        stat('Snittspar per månad', o.snittMan === null ? '–' : kr(o.snittMan), 'sparat i år / antal löner'),
        stat('Sparkvot i år', o.sparkvotAr === null ? '–' : pct(o.sparkvotAr, false), o.lonAntal ? `av ${kr(o.lonAr)} nettolön` : 'ingen lön inlagd i år'),
        stat('Utdelningar + återbäring', kr(o.utdAr + o.aterAr), (o.aterAr ? `varav SAVR ${kr(o.aterAr)}` : '') + (prevY ? `${o.aterAr ? ' · ' : ''}i fjol ${kr(prevY.utdAr + prevY.aterAr)}` : ''))),
      lonRader.length ? h('table', { class: 'tbl' }, h('thead', {}, h('tr', {}, ['Lönevecka', 'Nettolön', 'Sparat', 'Sparkvot'].map(x => h('th', {}, x)))),
        h('tbody', {}, lonRader.map(x => h('tr', {}, h('td', {}, dstr(x.d).replace(/ \d{4}$/, '')), h('td', {}, nf0.format(x.lon)),
          h('td', { class: cls(x.manSparat) }, x.manSparat === null ? '–' : nf0.format(x.manSparat)), h('td', {}, x.manKvot === null ? '–' : pct(x.manKvot, false)))))) : null,
      h('p', { class: 'note' }, 'Sparat = ökningen av egna insatta pengar + kassa sedan förra lönen. Sparkvot = sparat / nettolön.'));
  }
  // påminnelse om lön som kommit efter senaste inlagda vecka
  function periodLabel(i) {
    if (i < 1) return 'Senaste veckan';
    const dd = daysBetween(S.out[i - 1].d, S.out[i].d);
    return dd > 9 ? `Sedan förra (${dd} dagar)` : 'Senaste veckan';
  }
  function paminnelser() {
    const last = S.out[S.out.length - 1]; if (!last) return null;
    const t = today(), dd = daysBetween(last.d, t), list = [];
    if (dd >= 8) list.push(`Det är ${dd} dagar sedan senaste avläsningen (${dstr(last.d)}). Lägg in dagens värden – avkastningen räknas från förra avläsningen, så inget går förlorat.`);
    const y = +t.slice(0, 4), md = t.slice(5);
    if (md >= '12-24' && last.d < y + '-12-28') list.push('Gör gärna en avläsning den 31 december (eller så nära som möjligt) – den blir startvärdet för nästa år.');
    const ld = Calc.lonedagarMellan(last.d, t);
    if (ld.length) list.push(`Lönen kom ${ld.map(dstr).join(' och ')}. Kom ihåg nettolönen när du lägger in veckan.`);
    return list.length ? h('div', { class: 'card hint' }, list.map(x => h('div', {}, x))) : null;
  }
  function lonPaminnelse() {
    const last = S.out[S.out.length - 1]; if (!last) return null;
    const ld = Calc.lonedagarMellan(last.d, today());
    return ld.length ? h('div', { class: 'card hint' }, `Lönen kom ${ld.map(dstr).join(' och ')}. Kom ihåg nettolönen när du lägger in veckan.`) : null;
  }
  const daysBetween = (a, b) => Math.round((Date.parse(b) - Date.parse(a)) / 86400000);
  const shiftDays = (d, n) => new Date(Date.parse(d + 'T00:00:00Z') + n * 86400000).toISOString().slice(0, 10);
  const shiftYears = (d, y) => (+d.slice(0, 4) + y) + d.slice(4);

  const FIELDS = () => [
    ['Kassa', [['kassa', 'Lönekonto + buffert', 'value']]],
    ['Avanza – värde', [['av.av', 'Aktier', 'value'], ['av.fv', 'Fonder', 'value'], ['av.ev', 'ETF', 'value']]],
    ['Avanza – insatt denna vecka', [['av.ai', 'Aktier', 'flow'], ['av.fi', 'Fonder', 'flow'], ['av.ei', 'ETF', 'flow'], ['av.utd', 'Utdelning som kommit in', 'flow']]],
    ['SAVR', [['sv.val', 'Fondvärde', 'value'], ['sv.inv', 'Insatt denna vecka (egna pengar)', 'flow'], ['sv.ater', 'Återbäring', 'flow']]],
    [bankNamn(), [['db.val', 'Fondvärde', 'value'], ['db.inv', 'Insatt denna vecka', 'flow']]],
  ];
  const ADV = () => [['av.vinst', 'Avanza – realiserad vinst vid försäljning'], ['db.vinst', bankNamn() + ' – realiserad vinst vid försäljning']];
  const getP = (o, p) => p.split('.').reduce((a, k) => a == null ? undefined : a[k], o);
  const setP = (o, p, v) => { const k = p.split('.'); let a = o; for (let i = 0; i < k.length - 1; i++) a = a[k[i]] = a[k[i]] || {}; a[k[k.length - 1]] = v; };

  function viewForm(editDate) {
    const rows = S.data ? Calc.sortRows(S.data.rows) : [];
    const editing = editDate ? rows.find(r => r.d === editDate) : null;
    const prev = editing ? rows[rows.indexOf(editing) - 1] : rows[rows.length - 1];
    const form = h('form', { class: 'view form', onsubmit: ev => { ev.preventDefault(); submit(); } });
    const inputs = {};
    const fmtIn = v => (v === undefined || v === null || v === 0) ? '' : String(v).replace('.', ',');
    const field = (path, label, kind) => {
      const ph = kind === 'value' && prev ? 'förra: ' + nf0.format(getP(prev, path) || 0) : '0';
      const inp = h('input', { inputmode: 'decimal', autocomplete: 'off', placeholder: ph, value: editing ? fmtIn(getP(editing, path)) : '', 'data-kind': kind, oninput: preview });
      inputs[path] = inp;
      return h('label', { class: 'fld' }, h('span', {}, label), inp, kind === 'value' && prev ? h('button', { type: 'button', class: 'mini', title: 'Samma som förra veckan', onclick: () => { inp.value = fmtIn(getP(prev, path)); preview(); } }, '↺') : h('span', { class: 'mini' }));
    };
    const date = h('input', { type: 'date', value: editing ? editing.d : today(), max: today(), required: true, oninput: preview });
    const yearEnd = h('input', { type: 'checkbox', onchange: () => preview() });
    const yearEndBox = h('label', { class: 'chk', style: 'display:none' }, yearEnd, h('span', {}, ''));
    form.append(h('section', { class: 'card' }, h('h2', {}, editing ? 'Ändra vecka' : 'Ny vecka'), h('label', { class: 'fld' }, h('span', {}, 'Datum för avläsningen'), date),
      h('p', { class: 'note' }, 'Använd dagens datum, även om du är sen – värdena gäller den dag du läser av dem.'), yearEndBox));
    FIELDS().forEach(([title, fl]) => {
      const sec = h('section', { class: 'card' }, h('h3', {}, title), fl.map(f => field(...f)));
      if (title === 'Avanza – värde') {
        let box = null;
        const btn = h('button', { type: 'button', class: 'linkbtn', onclick: () => {
          if (box) { box.remove(); box = null; return; }
          box = pasteBox(res => { const t = res.totals; inputs['av.av'].value = fmtIn(Math.round(t.aktie)); inputs['av.fv'].value = fmtIn(Math.round(t.fond)); inputs['av.ev'].value = fmtIn(Math.round(t.etf)); box.remove(); box = null; preview(); },
            'Värdena för aktier, fonder och ETF fylls i, och kurserna under Innehav uppdateras.');
          sec.append(box); } }, 'Klistra in innehavslistan från Avanza');
        sec.append(btn);
      }
      form.append(sec);
    });
    // Lön: visas när en lönedag (25:e, lör -> 24:e, sön -> 26:e) infaller sedan förra veckan
    const lonInfo = h('p', { class: 'note' });
    const lonVis = h('button', { type: 'button', class: 'linkbtn', onclick: () => { lonForce = true; preview(); } }, 'Lägg till lön den här veckan ändå');
    let lonForce = !!(editing && typeof editing.lon === 'number');
    const lonSec = h('section', { class: 'card lon' }, h('h3', {}, 'Lön'), field('lon', 'Nettolön', 'opt'), lonInfo);
    form.append(lonSec, lonVis);
    const adv = h('details', { class: 'card', open: editing && ADV().some(([p]) => getP(editing, p)) ? true : null }, h('summary', {}, 'Avancerat: försäljningar'),
      h('p', { class: 'note' }, 'Sålde du och tog ut pengarna: ange anskaffningsvärdet som negativ insättning ovan och vinsten här (förlust som minus). Säljer du och köper nytt inom kontot behövs inget här.'),
      ADV().map(([p, l]) => field(p, l, 'flow')));
    form.append(adv);
    const pv = h('div', { class: 'preview' });
    form.append(h('section', { class: 'card sticky' }, pv, h('div', { class: 'btns' },
      editing ? h('button', { type: 'button', class: 'danger', onclick: del }, 'Ta bort') : null,
      h('button', { type: 'submit', class: 'primary', disabled: S.busy || null }, S.busy ? 'Sparar…' : 'Spara'))));

    function collect() {
      const r = { d: date.value, av: {}, db: {}, sv: {} }; const bad = [];
      for (const [p, inp] of Object.entries(inputs)) {
        const kind = inp.dataset.kind; let v = parseNum(inp.value);
        if (inp.value.trim() === '' && kind === 'value') { bad.push(inp); continue; }
        if (!isFinite(v)) { bad.push(inp); continue; }
        if (kind === 'opt' && inp.value.trim() === '') continue;
        setP(r, p, v);
      }
      Object.values(inputs).forEach(i => i.classList.toggle('bad', bad.includes(i)));
      return { r, bad };
    }
    function lonDagar() {
      if (!date.value) return [];
      const prevD = rows.filter(x => x.d < date.value && x !== editing).map(x => x.d).pop() || shiftDays(date.value, -7);
      return Calc.lonedagarMellan(prevD, date.value);
    }
    // Årsskifte: ny avläsning tidigt i januari utan avläsning runt 31 december
    function yearEndInfo() {
      if (!date.value || editing) return null;
      const y = +date.value.slice(0, 4), prevD = rows.filter(x => x.d < date.value).map(x => x.d).pop();
      if (!prevD || +prevD.slice(0, 4) !== y - 1 || prevD >= (y - 1) + '-12-28' || daysBetween(y + '-01-01', date.value) > 20) return null;
      return { prevD, ye: (y - 1) + '-12-31' };
    }
    function preview() {
      const ye = yearEndInfo();
      yearEndBox.style.display = ye ? '' : 'none'; if (!ye) yearEnd.checked = false;
      if (ye) yearEndBox.lastChild.textContent = `Ingen avläsning finns runt årsskiftet (senaste är ${dstr(ye.prevD)}). Spara samma värden även som 31 december, så att avkastningen hittills räknas till ${ye.ye.slice(0, 4)}.`;
      const ld = lonDagar(), visa = ld.length > 0 || lonForce;
      lonSec.style.display = visa ? '' : 'none'; lonVis.style.display = visa ? 'none' : '';
      if (!visa) inputs.lon.value = '';
      lonInfo.textContent = ld.length ? `Lönedag ${ld.map(dstr).join(', ')} infaller den här veckan.` : 'Ingen lönedag den här veckan enligt regeln (25:e, lör → 24:e, sön → 26:e).';
      const { r, bad } = collect();
      if (bad.length || !date.value) { pv.textContent = bad.length ? `Fyll i alla värden (${bad.length} saknas eller är felaktiga).` : ''; return; }
      const others = rows.filter(x => x.d !== r.d && x !== editing);
      const out = Calc.compute(others.concat([r])); const o = out.find(x => x.d === r.d);
      pv.innerHTML = '';
      pv.append(...[h('div', {}, 'Förmögenhet ', h('b', {}, kr(o.formogenhet))),
        h('div', {}, 'Veckans avkastning ', h('b', { class: cls(o.W) }, `${krS(o.W)} (${pct(o.X)})`)),
        o.lon !== null && o.manKvot !== null && o.manKvot !== undefined ? h('div', {}, 'Sparkvot sedan förra lönen ', h('b', {}, `${pct(o.manKvot, false)} (${kr(o.manSparat)})`)) : null,
        lonSec.style.display === '' && r.lon === undefined ? h('div', { class: 'warn' }, 'Lönedag den här veckan – fyll i nettolönen.') : null,
        rows.some(x => x.d === r.d && x !== editing) ? h('div', { class: 'warn' }, 'Det finns redan en vecka med detta datum – den ersätts.') : null].filter(Boolean));
    }
    async function submit() {
      const { r, bad } = collect();
      if (bad.length) { toast('Kontrollera de markerade fälten. Värden måste fyllas i; t.ex. 1 234,50 eller 1000+250.', 'err'); bad[0].focus(); return; }
      const ye = yearEnd.checked ? yearEndInfo() : null;
      const ok = await commit(d => {
        d.rows = d.rows.filter(x => x.d !== r.d && (!editing || x.d !== editing.d));
        if (ye && !d.rows.some(x => x.d === ye.ye)) {
          // årsskiftesrad: samma värden och veckans insättningar; den nya raden får inga insättningar
          const yr = JSON.parse(JSON.stringify(r)); yr.d = ye.ye; delete yr.lon;
          const nr = JSON.parse(JSON.stringify(r)); ['ai', 'fi', 'ei', 'utd', 'vinst'].forEach(k => nr.av[k] = 0); nr.db.inv = 0; nr.db.vinst = 0; nr.sv.inv = 0; nr.sv.ater = 0;
          if (r.lon !== undefined && ye.ye >= Calc.lonedag(+ye.ye.slice(0, 4), 12)) { yr.lon = r.lon; delete nr.lon; }
          d.rows.push(yr, nr);
        } else d.rows.push(r);
      }, editing ? 'Ändringen är sparad' : 'Veckan är sparad');
      if (ok) location.hash = '#oversikt';
    }
    async function del() {
      if (!confirm(`Ta bort veckan ${dstr(editing.d)}?`)) return;
      if (await commit(d => { d.rows = d.rows.filter(x => x.d !== editing.d); }, 'Veckan är borttagen')) location.hash = '#historik';
    }
    setTimeout(preview, 0);
    return form;
  }


  // ---------------------------------------------------------------- Avanza: innehav, GAV, utdelningar
  const AV = { data: null, etag: null, loaded: false, loading: false, err: null, an: null, filter: 'alla', open: null, paste: false, msg: null };
  async function avLoad(force) {
    if ((AV.loaded && !force) || AV.loading) return;
    AV.loading = true;
    try { const r = await Store.loadAvanza(); AV.data = r.data; AV.etag = r.etag; AV.loaded = true; AV.err = null; avAnalyze(); }
    catch (e) { AV.err = e.message; }
    AV.loading = false; if (location.hash.startsWith('#innehav')) render();
  }
  function avAnalyze() { AV.an = AV.data && (AV.data.tx || []).length ? Avanza.analyze(AV.data.tx, AV.data.holdings) : null; }
  async function avSave(mut, msg) {
    const next = AV.data ? JSON.parse(JSON.stringify(AV.data)) : { app: 'ekonomi-avanza', version: 1, tx: [], holdings: null };
    mut(next); next.updated = new Date().toISOString();
    try {
      AV.etag = await Store.saveAvanza(next, AV.etag); AV.data = next; avAnalyze();
      if (msg) toast(msg, 'ok'); backup(false); return true;
    } catch (e) { toast(e.message, 'err'); if (e.conflict) await avLoad(true); return false; }
    finally { render(); }
  }
  const slim = r => ({ d: r.d, konto: r.konto, typ: r.typ, namn: r.namn, antal: r.antal, kurs: r.kurs, belopp: r.belopp, courtage: r.courtage, isin: r.isin, resultat: r.resultat, seq: r.seq });
  async function avImportCSV(f) {
    if (!f) return;
    try {
      const rows = Avanza.parseCSV(await f.text()).map(slim);
      if (!AV.loaded) await avLoad();
      const m = Avanza.merge((AV.data && AV.data.tx) || [], rows);
      await avSave(d => { d.tx = m.rows; }, m.added ? `${m.added} nya transaktioner (${m.skipped} fanns redan)` : 'Inga nya transaktioner – allt fanns redan');
    } catch (e) { toast(e.message, 'err'); }
  }
  async function avSaveHoldings(text) {
    const h = Avanza.parseHoldings(text);
    if (!AV.loaded) await avLoad();
    await avSave(d => { d.holdings = { date: today(), items: h.items, totals: h.totals }; }, `Kurser uppdaterade (${h.items.length} innehav)`);
    return h;
  }
  function pasteBox(onDone, hint) {
    const ta = h('textarea', { rows: 6, placeholder: 'Markera hela innehavslistan på Avanza (Min ekonomi › Innehav), kopiera och klistra in här.' });
    return h('div', { class: 'paste' }, hint ? h('p', { class: 'note' }, hint) : null, ta,
      h('div', { class: 'btns' }, h('button', { type: 'button', class: 'primary', onclick: async () => {
        try { const res = await avSaveHoldings(ta.value); onDone && onDone(res); } catch (e) { toast(e.message, 'err'); } } }, 'Läs in')));
  }

  function viewInnehav() {
    if (!AV.loaded) { avLoad(); return h('div', { class: 'view' }, h('p', { class: 'note pad' }, AV.err ? 'Kunde inte hämta: ' + AV.err : 'Hämtar Avanza-data…')); }
    const file = h('input', { type: 'file', accept: '.csv,text/csv', style: 'display:none', onchange: ev => avImportCSV(ev.target.files[0]) });
    const importBtns = h('div', { class: 'btns' }, file,
      h('button', { class: 'ghost', onclick: () => file.click() }, 'Importera transaktioner (CSV)'),
      h('button', { class: 'ghost', onclick: () => { AV.paste = !AV.paste; render(); } }, AV.paste ? 'Stäng' : 'Uppdatera kurser'));
    const an = AV.an, hold = AV.data && AV.data.holdings;
    if (!an) return h('div', { class: 'view' }, h('section', { class: 'card' }, h('h2', {}, 'Innehav på Avanza'),
      h('p', {}, 'Läs in transaktionsexporten från Avanza för att se innehav, GAV, utdelningar och avslutade affärer.'),
      h('ol', { class: 'note' }, h('li', {}, 'Avanza › Min ekonomi › Transaktioner › välj hela perioden › Exportera (CSV).'),
        h('li', {}, 'Avanza › Min ekonomi › Innehav › kopiera listan och klistra in under "Uppdatera kurser".')),
      importBtns, AV.paste ? pasteBox(() => { AV.paste = false; }) : null));

    const cur = an.current.filter(c => AV.filter === 'alla' || c.typ === AV.filter);
    const tot = cur.reduce((s, c) => s + (c.varde || 0), 0), sk = cur.reduce((s, c) => s + (c.sedanKop || 0), 0);
    const yrs = an.years.filter(y => y.utd || y.realiserat);
    const row = c => {
      const isOpen = AV.open === c.name;
      const g = c.group;
      return [h('div', { class: 'li h', onclick: () => { AV.open = isOpen ? null : c.name; render(); } },
        h('span', { class: 'nm' }, c.name, h('small', {}, c.andel ? pct(c.andel, false) + ' av portföljen' : '')),
        h('span', {}, c.varde === null ? '–' : nf0.format(c.varde), h('small', { class: cls(c.sedanKop) }, c.sedanKopPct === null || c.sedanKopPct === undefined ? '' : pct(c.sedanKopPct))),
        h('span', { class: cls(c.total) }, c.total === null || c.total === undefined ? '–' : pct(c.total), h('small', {}, c.utd ? 'utd ' + nf0.format(c.utd) : ''))),
        isOpen ? h('div', { class: 'det' },
          h('table', { class: 'kv' }, h('tbody', {},
            [['Antal', c.qty === null ? '–' : nf2.format(c.qty)], ['GAV', c.gav ? nf2.format(c.gav) + ' kr' : '–'], ['Senaste kurs', c.senast ? nf2.format(c.senast) : '–'],
             ['Anskaffningsvärde', kr(c.ansk)], ['Sedan köp', krS(c.sedanKop)], ['Utdelningar totalt', kr(c.utd)], ['Köpt totalt', kr(c.kopt)], ['Sålt totalt', kr(c.salt)]]
              .map(([k, v]) => h('tr', {}, h('td', {}, k), h('td', {}, v))))),
          g ? txList(g.tx) : null) : null];
    };
    return h('div', { class: 'view' },
      h('section', { class: 'hero' }, h('div', { class: 'lbl' }, 'Innehav Avanza' + (hold ? ' · kurser ' + dstr(hold.date) : '')),
        h('div', { class: 'huge' }, hold ? kr(tot) : '–'), hold ? h('div', { class: 'sub' }, h('span', { class: cls(sk) }, krS(sk)), ' sedan köp') : h('div', { class: 'sub' }, 'Klistra in innehavslistan för att se värden.')),
      h('section', { class: 'card' }, importBtns, AV.paste ? pasteBox(() => { AV.paste = false; }) : null,
        h('p', { class: 'note' }, `${an.count} transaktioner ${an.first ? dstr(an.first) : ''} – ${an.lastDate ? dstr(an.lastDate) : ''}.`)),
      h('section', { class: 'card list' }, h('div', { class: 'row pad2' }, h('h2', {}, 'Innehav'),
        h('div', { class: 'seg' }, [['alla', 'Alla'], ['aktie', 'Aktier'], ['fond', 'Fonder'], ['etf', 'ETF']].map(([k, l]) => h('button', { class: AV.filter === k ? 'on' : '', onclick: () => { AV.filter = k; render(); } }, l)))),
        h('div', { class: 'li head h' }, h('span', {}, 'Namn'), h('span', {}, 'Värde kr'), h('span', {}, 'Total avk.')),
        cur.map(row).flat(),
        h('p', { class: 'note pad2' }, 'Värde och "sedan köp" kommer från Avanzas lista. Total avk. = (värde + sålt + utdelningar − köpt) / köpt, för hela tiden du ägt aktien.')),
      h('section', { class: 'card' }, h('h2', {}, 'Utdelningar och resultat per år'),
        yrs.length ? barChartKr(yrs.map(y => String(y.year)), yrs.map(y => y.utdNetto)) : null,
        h('table', { class: 'tbl' }, h('thead', {}, h('tr', {}, ['År', 'Utdelning', 'Källskatt', 'Realiserat', 'Skatt/avg.'].map(x => h('th', {}, x)))),
          h('tbody', {}, an.years.slice().reverse().map(y => h('tr', {}, h('td', {}, y.year), h('td', {}, nf0.format(y.utd)), h('td', {}, nf0.format(y.kallskatt)),
            h('td', { class: cls(y.realiserat) }, nf0.format(y.realiserat)), h('td', {}, nf0.format(y.skatt)))))),
        h('p', { class: 'note' }, 'Realiserat = försäljningspris − anskaffningsvärde (genomsnittsmetoden per konto). Skatt/avg. = avkastningsskatt (ISK/KF), riskkostnad, räntor och avgifter.')),
      h('section', { class: 'card list' }, h('h2', { class: 'pad2' }, `Avslutade innehav (${an.closed.length})`),
        h('div', { class: 'li head h' }, h('span', {}, 'Namn'), h('span', {}, 'Resultat kr'), h('span', {}, 'Total avk.')),
        an.closed.map(c => { const isOpen = AV.open === 'x:' + c.name + c.last; return [h('div', { class: 'li h', onclick: () => { AV.open = isOpen ? null : 'x:' + c.name + c.last; render(); } },
          h('span', { class: 'nm' }, c.name, h('small', {}, 'sålt ' + dstr(c.last))), h('span', { class: cls(c.resultat) }, nf0.format(c.resultat), h('small', {}, c.utd ? '+ utd ' + nf0.format(c.utd) : '')),
          h('span', { class: cls(c.total) }, c.total === null ? '–' : pct(c.total))),
          isOpen ? h('div', { class: 'det' }, h('table', { class: 'kv' }, h('tbody', {}, [['Köpt för', kr(c.kopt)], ['Sålt för', kr(c.salt)], ['Utdelningar', kr(c.utd)]].map(([k, v]) => h('tr', {}, h('td', {}, k), h('td', {}, v))))), txList(c.group.tx)) : null]; }).flat()));
  }
  function txList(tx) {
    const list = tx.slice().reverse();
    return h('details', { class: 'txl' }, h('summary', {}, `Transaktioner (${list.length})`),
      h('table', { class: 'tbl small' }, h('tbody', {}, list.slice(0, 150).map(t => h('tr', {}, h('td', {}, t.d), h('td', {}, t.typ), h('td', {}, t.antal === null ? '' : nf2.format(t.antal)), h('td', {}, t.belopp === null ? '' : nf0.format(t.belopp)))))));
  }
  function barChartKr(labels, values) {
    const W = chartW(), H = 170, L = 50, R = 8, T = 16, B = 22, hi = Math.max(1, ...values), step = niceStep(hi / 4), top = Math.ceil(hi / step) * step;
    const y = v => T + (H - T - B) * (1 - v / top), bw = (W - L - R) / labels.length; let g = '';
    for (let v = 0; v <= top + step / 2; v += step) g += `<line x1="${L}" x2="${W - R}" y1="${y(v)}" y2="${y(v)}" class="grid"/><text x="${L - 6}" y="${y(v) + 4}" class="ax" text-anchor="end">${mkr(v)}</text>`;
    labels.forEach((l, i) => { const x0 = L + i * bw + bw * .18, w = bw * .64, v = Math.max(0, values[i]);
      g += `<rect x="${x0}" y="${y(v)}" width="${w}" height="${y(0) - y(v)}" class="bar pos" rx="3"/><text x="${x0 + w / 2}" y="${y(v) - 4}" class="ax val" text-anchor="middle">${mkr(values[i])}</text><text x="${x0 + w / 2}" y="${H - 6}" class="ax" text-anchor="middle">${l}</text>`; });
    return h('div', { class: 'chart', html: `<svg viewBox="0 0 ${W} ${H}">${g}</svg>` });
  }

  function viewHistory() {
    const out = S.out.slice().reverse();
    let shown = 60;
    const idx = new Map(S.out.map((o, i) => [o, i]));
    const gap = o => { const i = idx.get(o); if (!i) return 0; const d = daysBetween(S.out[i - 1].d, o.d); return d > 9 ? d : 0; };
    const list = h('div', { class: 'card list' });
    const more = h('button', { class: 'ghost wide', onclick: () => { shown += 100; fill(); } }, 'Visa fler');
    function fill() {
      list.innerHTML = '';
      list.append(h('div', { class: 'li head' }, h('span', {}, 'Datum'), h('span', {}, 'Förmögenhet'), h('span', {}, 'Veckan')));
      let yr = null;
      out.slice(0, shown).forEach(o => {
        if (o.year !== yr) { yr = o.year; list.append(h('div', { class: 'yr' }, String(yr))); }
        list.append(h('a', { class: 'li', href: '#ny?d=' + o.d }, h('span', {}, dstr(o.d), o.lon !== null ? h('i', { class: 'tag' }, 'lön') : null), h('span', {}, kr(o.formogenhet)),
          h('span', { class: cls(o.W) }, o.X === null ? '–' : `${krS(o.W)}  ${pct(o.X)}`, gap(o) ? h('small', { class: 'gap' }, `${gap(o)} dagar`) : null)));
      });
      more.style.display = shown < out.length ? '' : 'none';
    }
    fill();
    return h('div', { class: 'view' }, h('p', { class: 'note pad' }, 'Tryck på en vecka för att ändra eller ta bort den.'), list, more);
  }

  function viewMore() {
    const lb = localStorage.getItem('ek.lastBackup');
    const malIn = h('input', { inputmode: 'numeric', value: nf0.format((S.data.settings || {}).mal || 0) });
    const bankIn = h('input', { type: 'text', value: (S.data.settings || {}).bank || '', placeholder: 'Bank' });
    const file = h('input', { type: 'file', accept: '.json,application/json', style: 'display:none', onchange: ev => importFile(ev.target.files[0]) });
    return h('div', { class: 'view' },
      h('section', { class: 'card' }, h('h2', {}, 'Konto och lagring'),
        Store.isDemo() ? h('p', { class: 'warn' }, 'Demoläge – datan sparas bara i den här webbläsaren.')
          : h('p', {}, 'Inloggad som ', h('b', {}, Store.account() || 'Microsoft-konto'), '. Datan ligger i din OneDrive under ', h('b', {}, 'Appar › (appens namn)'), '.'),
        h('p', { class: 'note' }, S.syncedAt ? 'Senast synkad ' + new Date(S.syncedAt).toLocaleString('sv-SE') : 'Visar sparad kopia – inte synkad ännu.'),
        h('div', { class: 'btns' }, h('button', { class: 'ghost', onclick: async () => { try { await reload(); toast('Uppdaterad', 'ok'); } catch (e) { toast(e.message, 'err'); } render(); } }, 'Hämta senaste'),
          !Store.isDemo() ? h('button', { class: 'ghost', onclick: () => { if (confirm('Logga ut på den här enheten?')) { Store.logout(); location.reload(); } } }, 'Logga ut') : null)),
      h('section', { class: 'card' }, h('h2', {}, 'Excel-backup'),
        h('p', {}, 'Varje gång du sparar en vecka skrivs ', h('b', {}, 'Ekonomi-backup.xlsx'), ' till OneDrive, plus en kopia per månad (Ekonomi-backup-ÅÅÅÅ-MM.xlsx). OneDrive sparar dessutom äldre versioner av datafilen.'),
        h('p', { class: 'note' }, lb ? 'Senaste backup: ' + new Date(lb).toLocaleString('sv-SE') : 'Ingen backup gjord från den här enheten ännu.'),
        h('div', { class: 'btns' }, !Store.isDemo() ? h('button', { class: 'ghost', onclick: () => backup(true) }, 'Spara backup nu') : null,
          h('button', { class: 'ghost', onclick: () => download(workbook()) }, 'Ladda ner Excel'),
          h('button', { class: 'ghost', onclick: () => download(new Blob([JSON.stringify(S.data)], { type: 'application/json' }), `ekonomi-data-${today()}.json`) }, 'Ladda ner rådata (JSON)'))),
      h('section', { class: 'card' }, h('h2', {}, 'Inställningar'),
        h('label', { class: 'fld' }, h('span', {}, 'Mål för förmögenhet (kr)'), malIn),
        h('label', { class: 'fld' }, h('span', {}, 'Namn på banken'), bankIn),
        h('div', { class: 'btns' }, h('button', { class: 'ghost', onclick: () => { const v = parseNum(malIn.value); if (!isFinite(v)) return toast('Ogiltigt belopp', 'err');
          const b = bankIn.value.trim(); commit(d => { d.settings = d.settings || {}; d.settings.mal = v; if (b) d.settings.bank = b; else delete d.settings.bank; }, 'Inställningarna är sparade'); } }, 'Spara'))),
      h('section', { class: 'card' }, h('h2', {}, 'Importera'),
        h('p', { class: 'note' }, 'Ersätter all data med innehållet i en JSON-fil (t.ex. startdatan från Excel eller en nedladdad rådatafil).'),
        file, h('div', { class: 'btns' }, h('button', { class: 'ghost', onclick: () => file.click() }, 'Välj fil…'))),
      h('section', { class: 'card' }, h('h2', {}, 'Så räknar appen'),
        h('ul', { class: 'note' },
          h('li', {}, 'Egna insatta pengar = allt du satt in, minus återinvesterade utdelningar och realiserade vinster.'),
          h('li', {}, 'Veckans utveckling = veckans avkastning / (förra veckans värde + veckans insättningar).'),
          h('li', {}, 'I år (tidsviktad) = veckornas utveckling kedjade sedan årsskiftet. Påverkas inte av när du sätter in pengar.'),
          h('li', {}, 'I år (pengaviktad) = årets avkastning / genomsnittligt insatt kapital under året.'),
          h('li', {}, 'Insatt på Avanza ska som tidigare inkludera utdelningar du köper för.'),
          h('li', {}, 'Lönedag: den 25:e, lördag → 24:e, söndag → 26:e. December: den 23:e (helg → fredagen före). Lönefältet visas i den vecka lönen kommer.'),
          h('li', {}, 'SAVR: "Insatt" är bara dina egna pengar. Återbäring räknas som avkastning, precis som utdelningar.'),
          h('li', {}, 'Sparkvot per lön = ökningen av egna insatta pengar + kassa sedan förra lönen / nettolönen.'))),
      h('p', { class: 'note pad' }, 'Version ' + (window.EKONOMI_VERSION || '')));
  }

  async function importFile(f) {
    if (!f) return;
    try {
      const d = JSON.parse(await f.text());
      if (!d || !Array.isArray(d.rows) || !d.rows.every(r => /^\d{4}-\d{2}-\d{2}$/.test(r.d))) throw new Error('Filen ser inte ut som ekonomidata.');
      const n = S.data && S.data.rows.length;
      if (n && !confirm(`Ersätt nuvarande ${n} veckor med ${d.rows.length} veckor från filen?`)) return;
      const keepEtag = S.etag;
      if (!S.data) S.data = { app: 'ekonomi', version: 2, settings: {}, rows: [] };
      S.etag = keepEtag;
      Calc.migrate(d);
      await commit(x => { x.rows = d.rows; x.version = d.version; x.settings = Object.assign({}, x.settings, d.settings || {}); }, `Importerade ${d.rows.length} veckor`);
      location.hash = '#oversikt';
    } catch (e) { toast(e.message, 'err'); }
  }

  function viewEmpty() {
    const file = h('input', { type: 'file', accept: '.json,application/json', style: 'display:none', onchange: ev => importFile(ev.target.files[0]) });
    return h('div', { class: 'view' }, h('section', { class: 'card center' }, h('h2', {}, 'Välkommen'),
      h('p', {}, 'Det finns ingen data ännu. Importera startdatan (ekonomi-startdata.json) som skapades från din Excel-fil, eller börja med en ny vecka.'),
      file, h('div', { class: 'btns col' }, h('button', { class: 'primary', onclick: () => file.click() }, 'Importera startdata…'),
        h('a', { class: 'ghost btn', href: '#ny' }, 'Börja från noll')),
      Store.isDemo() ? h('p', { class: 'warn' }, 'Demoläge – datan sparas bara i den här webbläsaren.') : null));
  }

  function viewLogin(msg) {
    return h('div', { class: 'view' }, h('section', { class: 'card center' }, h('div', { class: 'logo' }, h('img', { src: 'icons/icon-192.png', alt: '' })),
      h('h2', {}, 'Ekonomi'), h('p', {}, 'Logga in med ditt Microsoft-konto. Datan sparas i din egen OneDrive.'),
      msg ? h('p', { class: 'warn' }, msg) : null,
      h('div', { class: 'btns col' }, h('button', { class: 'primary', onclick: () => Store.login() }, 'Logga in med Microsoft'))));
  }

  // ---------------------------------------------------------------- render
  function render() {
    const main = $('#main'); const hash = location.hash || '#oversikt';
    const [route, qs] = hash.slice(1).split('?');
    const nav = $('#nav'); nav.style.display = '';
    document.querySelectorAll('#nav a').forEach(a => a.classList.toggle('on', a.getAttribute('href') === '#' + route));
    $('#status').textContent = S.busy ? 'Sparar…' : S.offline ? 'Offline – visar sparad kopia' : Store.isDemo() ? 'Demoläge' : '';
    let v;
    if (S.needLogin) { v = viewLogin(S.loginMsg); nav.style.display = 'none'; }
    else if (!S.data) { v = S.loading ? h('div', { class: 'view' }, h('p', { class: 'note pad' }, 'Hämtar data från OneDrive…')) : viewEmpty(); }
    else if (route === 'ny') v = viewForm(new URLSearchParams(qs || '').get('d'));
    else if (!S.data.rows.length) v = viewEmpty();
    else if (route === 'historik') v = viewHistory();
    else if (route === 'innehav') v = viewInnehav();
    else if (route === 'mer') v = viewMore();
    else v = viewOverview();
    // behåll formulär vid omritning medan man skriver
    if (route === 'ny' && main.firstChild && main.firstChild.tagName === 'FORM' && main.dataset.route === hash) {
      const btn = main.querySelector('button[type=submit]'); if (btn) { btn.disabled = S.busy; btn.textContent = S.busy ? 'Sparar…' : 'Spara'; } return;
    }
    main.dataset.route = hash; main.replaceChildren(v);
  }
  let rz; window.addEventListener('resize', () => { clearTimeout(rz); rz = setTimeout(() => { if (!location.hash.startsWith('#ny')) render(); }, 250); });
  window.addEventListener('hashchange', () => { render(); window.scrollTo(0, 0); });

  async function boot() {
    try {
      const r = await Store.handleRedirect();
      if (r && r.needLogin) { S.needLogin = true; return render(); }
    } catch (e) { S.needLogin = true; S.loginMsg = 'Inloggningen misslyckades: ' + e.message; return render(); }
    if (!Store.isDemo() && !Store.signedIn()) { S.needLogin = true; return render(); }
    const c = Store.cached();
    if (c) { S.data = c.data; S.etag = c.etag; Calc.migrate(S.data); recompute(); }
    S.loading = true; render();
    try { await reload(); S.offline = false; }
    catch (e) {
      if (e.needLogin) { S.needLogin = true; S.loginMsg = 'Inloggningen har gått ut. Logga in igen.'; }
      else { S.offline = true; if (!c) toast(e.message, 'err'); }
    }
    S.loading = false; render();
  }
  window.addEventListener('online', () => { if (S.offline) boot(); });
  boot();
  if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js');
})();
