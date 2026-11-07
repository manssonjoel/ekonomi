/* Ekonomi – appens gränssnitt */
window.EKONOMI_VERSION = '1.8.1';
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
  const S = { data: null, etag: null, out: [], busy: false, range: '3', chart: 'formogenhet', lastErr: null, syncedAt: null, offline: false };

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
    const head = ['Datum', 'Kassa', 'Avanza sparkonto', 'Avanza insatt aktier', 'Avanza insatt fonder', 'Avanza insatt ETF', 'Avanza värde aktier', 'Avanza värde fonder',
      'Avanza värde ETF', 'Utdelning', 'Avanza skatt & avgifter', 'Avanza realiserad vinst', B + ' insatt', B + ' värde', B + ' realiserad vinst', 'SAVR insatt (egna pengar)', 'SAVR värde',
      'SAVR återbäring', 'Nettolön', 'Totalt värde', 'Egna insatta pengar', 'Avkastat totalt', 'Avkastning veckan', 'Utveckling veckan',
      'Utveckling i år (tidsviktad)', 'Utveckling i år (pengaviktad)', 'Avkastat i år', 'Totalt på egna pengar', 'Förmögenhet', 'Sparat i år',
      'Avanza i år', B + ' i år', 'SAVR i år', 'Sparat sedan förra lönen', 'Sparkvot (månad)', 'Snittspar per månad i år'];
    const wk = [head].concat(rows.map((r, i) => { const o = out[i], a = r.av || {}, d = r.db || {}, s = r.sv || {};
      return [{ date: r.d }, K(r.kassa), K(r.kassa2), K(a.ai), K(a.fi), K(a.ei), K(a.av), K(a.fv), K(a.ev), K(a.utd), K(a.avg), K(a.vinst), K(d.inv), K(d.val), K(d.vinst),
        K(s.inv), K(s.val), K(s.ater), K(r.lon), K(o.N), K(o.K), K(o.V), K(o.W), P(o.X), P(o.Y), P(o.O), K(o.U), P(o.P), K(o.formogenhet),
        K(o.sparatAr), P(o.avY), P(o.dbY), P(o.svY), K(o.manSparat), P(o.manKvot), K(o.lon !== null ? o.snittMan : null)]; }));
    const yrs = Calc.years(out);
    const yr = [['År', 'Utveckling (tidsviktad)', 'Utveckling (pengaviktad)', 'Avkastat kr', 'Sparat', 'Utdelningar', 'Nettolön', 'Portföljvärde', 'Förmögenhet',
      'Avanza', B, 'SAVR', 'Sparkvot', 'Snittspar per månad']].concat(yrs.map(o => [o.year, P(o.Y), P(o.O), K(o.U), K(o.sparatAr), K(o.utdAr), K(o.lonAr), K(o.N), K(o.formogenhet), P(o.avY), P(o.dbY), P(o.svY), P(o.sparkvotAr), K(o.snittMan)]));
    const info = [['Ekonomi – backup'], ['Skapad: ' + new Date().toLocaleString('sv-SE')],
      ['Bladet "Veckor" innehåller all rådata (kolumn A–Q) och de uträknade värdena.'],
      ['Rådatan räcker för att återskapa allt. Originalfilen Joel_Ekonomi.xlsm innehåller historiken med formler.']];
    return XLSX.build([{ name: 'Veckor', rows: wk, widths: [11].concat(Array(35).fill(14)) }, { name: 'År', rows: yr, widths: [8].concat(Array(13).fill(16)) },
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
  function chartW() { const m = document.getElementById('main'); return Math.max(300, Math.min(736, (m ? m.clientWidth : 640) - 26)); }
  function niceStep(r) { const p = Math.pow(10, Math.floor(Math.log10(r))); const m = r / p; return (m <= 1 ? 1 : m <= 2 ? 2 : m <= 2.5 ? 2.5 : m <= 5 ? 5 : 10) * p; }
  // ---------------------------------------------------------------- vyer
  // ---------------------------------------------------------------- gemensamma bitar för översikt och detaljsidor
  const C = window.Charts;
  const card = (title, href, ...kids) => h(href ? 'a' : 'section', { class: 'card' + (href ? ' link' : ''), href }, title ? h('div', { class: 'ch' }, h('h2', {}, title), href ? h('span', { class: 'more' }, '›') : null) : null, ...kids);
  const stat = (label, main, sub, c) => h('div', { class: 'stat' }, h('div', { class: 'lbl' }, label), h('div', { class: 'big ' + (c || '') }, main), sub ? h('div', { class: 'sub' }, sub) : null);
  const back = () => h('a', { class: 'back', href: '#oversikt' }, '‹ Översikt');
  const rangeSeg = () => h('div', { class: 'seg' }, [['1', '1 år'], ['3', '3 år'], ['all', 'Allt']].map(([k, l]) => h('button', { class: S.range === k ? 'on' : '', onclick: ev => { ev.preventDefault(); S.range = k; render(); } }, l)));
  function slice() { const out = S.out, o = out[out.length - 1]; if (S.range === 'all') return out; const from = shiftYears(o.d, -(+S.range)); return out.filter(x => x.d >= from); }
  const legend = items => h('div', { class: 'legend' }, items.map(([cl, t]) => h('span', {}, h('i', { class: 'dot ' + cl }), t)));
  const yearsDone = () => { const yrs = Calc.years(S.out), o = S.out[S.out.length - 1]; return yrs.map(y => ({ ...y, label: y.year + (y === yrs[yrs.length - 1] && o.d.slice(5) < '12-28' ? '*' : '') })); };
  const utdYtd = o => o.utdAr + o.aterAr;
  function sameTimeLastYear(field) {
    const o = S.out[S.out.length - 1], day = C.dayOfYear(o.d); let v = null;
    S.out.forEach(x => { if (x.year === o.year - 1 && C.dayOfYear(x.d) <= day) v = typeof field === 'function' ? field(x) : x[field]; });
    return v;
  }
  // År mot år: i år och i fjol som standard, övriga år väljs med knappar
  const YCLS = ['c1', 'c2', 'c3', 'c4', 'c5', 'c7', 'c8'];
  function yearCompare(key, fn) {
    const ys = [...new Set(S.out.map(x => x.year))], cur = ys[ys.length - 1];
    S.ysel = S.ysel || {}; const sel = S.ysel[key] || (S.ysel[key] = [cur, cur - 1].filter(y => ys.includes(y)));
    const color = y => YCLS[(cur - y) % YCLS.length];
    const series = sel.slice().sort().map(y => ({ name: String(y), cls: color(y), pts: S.out.filter(x => x.year === y).map(x => [C.dayOfYear(x.d), fn(x)]) }));
    return [h('div', { class: 'chips small' }, ys.slice().reverse().map(y => h('button', { class: sel.includes(y) ? 'on ' + color(y) : '', onclick: ev => { ev.preventDefault();
        const k = sel.indexOf(y); if (k >= 0) { if (sel.length > 1) sel.splice(k, 1); } else sel.push(y); render(); } }, String(y)))),
      C.yearLines(series, { labels: false }),
      h('div', { class: 'ytable' }, series.slice().reverse().map(s => { const last = s.pts.length ? s.pts[s.pts.length - 1][1] : 0;
        return h('div', {}, h('i', { class: 'dot ' + s.cls }), h('span', {}, s.name), h('b', {}, kr(last)), h('small', {}, s.name == cur ? 'hittills i år' : 'helår')); }))];
  }
  // diagram (återanvänds på flera sidor)
  const G = {
    formogenhet: sl => [C.line(sl.map(x => x.d), [{ name: 'Förmögenhet', values: sl.map(x => x.formogenhet), cls: 'c1', fill: true }, { name: 'Portföljvärde', values: sl.map(x => x.N), cls: 'c3' }, { name: 'Egna insatta pengar', values: sl.map(x => x.K), cls: 'c2', dash: true }]),
      legend([['c1', 'Förmögenhet'], ['c3', 'Portföljvärde'], ['c2', 'Egna insatta pengar']])],
    perAr: () => { const y = yearsDone(); return [C.bars(y.map(x => x.label), [{ name: 'Utveckling', values: y.map(x => x.Y) }], { kind: 'pct', labels: true })]; },
    inkomst: () => { const y = yearsDone().filter(x => x.lonAr > 0);
      return [C.bars(y.map(x => x.label), [{ name: 'Nettoinkomst', values: y.map(x => x.lonAr), cls: 'c3' }, { name: 'Sparat', values: y.map(x => x.sparatAr), cls: 'c1' }, { name: 'Utdelningar', values: y.map(x => utdYtd(x)), cls: 'c2' }],
        { line: { name: 'Sparkvot', values: y.map(x => x.sparkvotAr), cls: 'c4', kind: 'pct' } }), legend([['c3', 'Nettoinkomst'], ['c1', 'Sparat'], ['c2', 'Utdelningar'], ['c4', 'Sparkvot (höger axel)']])]; },
    vecka: sl => [C.timeBars(sl.map(x => x.d), sl.map(x => x.W), { name: 'Avkastning', extra: i => `<div>${pct(sl[i].X)}</div>` })],
    total: sl => [C.timeBars(sl.map(x => x.d), sl.map(x => x.V), { name: 'Avkastat totalt', cls: 'c1' })],
    utv: sl => [C.line(sl.map(x => x.d), [{ name: 'Totalt på egna pengar', values: sl.map(x => x.P), cls: 'c1', fill: true }], { kind: 'pct' })],
    iAr: () => { const o = S.out[S.out.length - 1], r = S.out.filter(x => x.year === o.year); return [C.line(r.map(x => x.d), [{ name: 'Tidsviktad', values: r.map(x => x.Y), cls: 'c1', fill: true }, { name: 'Pengaviktad', values: r.map(x => x.O), cls: 'c2', dash: true }], { kind: 'pct' }), legend([['c1', 'Tidsviktad'], ['c2', 'Pengaviktad (på insatt kapital)']])]; },
  };
  const CHARTS = [['formogenhet', 'Förmögenhet', true], ['perAr', 'Per år', false], ['inkomst', 'Inkomst & sparkvot', false], ['vecka', 'Per vecka', true], ['total', 'Total avkastning', true], ['utv', 'Utveckling %', true]];
  function fordelning(o) {
    const items = [{ name: 'Aktier', value: o.aav, cls: 'c1' }, { name: 'Fonder', value: o.afv + o.Sval + o.Dval, cls: 'c3' }, { name: 'ETF', value: o.aev, cls: 'c2' }];
    return C.donut(items, { center: [C.mkr(o.N), 'investerat'] });
  }

  // ---------------------------------------------------------------- översikt
  function viewOverview() {
    const out = S.out, o = out[out.length - 1], n = out.length;
    const mal = (S.data.settings && S.data.settings.mal) || 0;
    const ch = CHARTS.find(c => c[0] === S.chart) || CHARTS[0];
    const fjol = sameTimeLastYear(utdYtd);
    const lastLon = out.filter(x => x.lon !== null && x.manKvot !== null && x.manKvot !== undefined).pop();
    return h('div', { class: 'view' },
      paminnelser(),
      h('section', { class: 'hero' },
        h('div', { class: 'lbl' }, 'Total förmögenhet · ' + dstr(o.d)),
        h('div', { class: 'huge' }, kr(o.formogenhet)),
        h('div', { class: 'sub' }, h('span', { class: cls(o.formAr) }, krS(o.formAr)), ' i år'),
        mal ? h('a', { class: 'goal link', href: '#prognos' }, h('div', { class: 'bar' }, h('i', { style: `width:${Math.min(100, o.formogenhet / mal * 100)}%` })),
          h('div', { class: 'sub' }, `${nf1.format(o.formogenhet / mal * 100)} % av målet ${mkr(mal)}`, prognosKort(o, mal))) : null),
      card('Avkastning', '#avkastning', h('div', { class: 'grid3' },
        stat(periodLabel(n - 1), pct(o.X), krS(o.W), cls(o.X)), stat('I år', pct(o.Y), krS(o.U), cls(o.Y)), stat('Totalt', pct(o.P), krS(o.V), cls(o.P)))),
      card('Konton', null, h('div', { class: 'acclist' },
        [['avanza', 'Avanza', o.Aval, o.avY], ['savr', 'SAVR', o.Sval, o.svY], ['bank', bankNamn(), o.Dval, o.dbY], ['kassa', 'Kassa', o.kassa, null]]
          .map(([k, nm, v, y]) => h('a', { class: 'accrow', href: '#konto?k=' + k }, h('span', {}, nm), h('span', {}, kr(v)), h('span', { class: cls(y) }, y === null ? '' : pct(y)), h('span', { class: 'more' }, '›'))))),
      card('Sparande ' + o.year, '#sparande', h('div', { class: 'grid3' },
        stat('Sparat i år', kr(o.sparatAr), o.snittMan !== null ? `≈ ${kr(o.snittMan)}/mån` : ''),
        stat('Sparkvot i år', o.sparkvotAr === null ? '–' : pct(o.sparkvotAr, false), lastLon ? `senaste lönen ${pct(lastLon.manKvot, false)}` : ''),
        stat('Utdelningar i år', kr(utdYtd(o)), fjol !== null ? `i fjol vid denna tid ${kr(fjol)}` : ''))),
      card('Fördelning', null, fordelning(o)),
      h('section', { class: 'card' }, h('div', { class: 'chips' }, CHARTS.map(([k, l]) => h('button', { class: ch[0] === k ? 'on' : '', onclick: () => { S.chart = k; render(); } }, l))),
        ch[2] ? h('div', { class: 'row end' }, rangeSeg()) : null,
        G[ch[0]](slice())));
  }


  // ---------------------------------------------------------------- prognos för målet
  // Historisk avkastning: genomsnitt per år (geometriskt) för hela kalenderår
  function histAvk() {
    const first = S.out[0].d; // första året räknas bara om det börjar i januari
    const y = Calc.years(S.out).filter((x, i, a) => i < a.length - 1 && x.Y !== null && (i > 0 || first.slice(5, 7) === '01'));
    if (!y.length) return null;
    const g = y.reduce((p, x) => p * (1 + x.Y), 1);
    return { r: Math.pow(g, 1 / y.length) - 1, from: y[0].year, to: y[y.length - 1].year };
  }
  // Snittsparande per månad senaste 12 månaderna (ökning av egna insatta pengar + kassa)
  function snittSpar12() {
    const out = S.out, o = out[out.length - 1], from = shiftYears(o.d, -1);
    const a = out.find(x => x.d >= from) || out[0];
    const months = Math.max(1, daysBetween(a.d, o.d) / 30.44);
    return ((o.K + o.kassa) - (a.K + a.kassa)) / months;
  }
  function progAnt() {
    let st = {}; try { st = JSON.parse(localStorage.getItem('ek.prognos') || '{}'); } catch (e) { }
    const hs = histAvk();
    return { r: st.r ?? 0.07, s: st.s ?? Math.max(0, Math.round(snittSpar12() / 500) * 500), ar: st.ar ?? (new Date().getFullYear() + 3), hist: hs };
  }
  // Månadsvis framskrivning: portföljen växer med r per år, kassan står still, sparandet går in varje månad
  function framskriv(o, r, s, maxMan = 480, mal = Infinity) {
    const rm = Math.pow(1 + r, 1 / 12) - 1; let N = o.N; const Q = o.kassa, pts = [o.formogenhet]; let hit = null;
    for (let m = 1; m <= maxMan; m++) { N = N * (1 + rm) + s; pts.push(N + Q); if (hit === null && N + Q >= mal) hit = m; }
    return { pts, hit };
  }
  const manadStr = (d0, m) => { const d = new Date(d0 + 'T12:00:00'); d.setMonth(d.getMonth() + m); return MONL[d.getMonth()] + ' ' + d.getFullYear(); };
  const MONL = ['januari', 'februari', 'mars', 'april', 'maj', 'juni', 'juli', 'augusti', 'september', 'oktober', 'november', 'december'];
  const tidStr = m => { const y = Math.floor(m / 12), mm = m % 12; return [y ? y + ' år' : '', mm ? mm + ' mån' : ''].filter(Boolean).join(' ') || 'nu'; };
  function prognosKort(o, mal) {
    if (o.formogenhet >= mal) return ' · målet är nått!';
    const a = progAnt(), f = framskriv(o, a.r, a.s, 600, mal);
    return f.hit ? ` · nås ${manadStr(o.d, f.hit)} ›` : ' ›';
  }
  // Månadssparande som krävs för att nå målet vid en viss tidpunkt (binärsökning)
  function kravSpar(o, r, mal, man) {
    if (man <= 0) return null; let lo = 0, hi = 1e6;
    if (framskriv(o, r, 0, man).pts[man] >= mal) return 0;
    for (let k = 0; k < 50; k++) { const mid = (lo + hi) / 2; if (framskriv(o, r, mid, man).pts[man] >= mal) hi = mid; else lo = mid; }
    return hi;
  }
  function viewPrognos() {
    const out = S.out, o = out[out.length - 1], mal = (S.data.settings && S.data.settings.mal) || 0;
    const a = progAnt(), save = () => localStorage.setItem('ek.prognos', JSON.stringify({ r: a.r, s: a.s, ar: a.ar }));
    if (!mal) return h('div', { class: 'view' }, back(), card('Prognos', null, h('p', {}, 'Ange ett mål för förmögenheten under Mer › Inställningar.')));
    const box = h('div');
    const slider = (label, min, max, step, val, fmtV, on) => {
      const out = h('b', {}, fmtV(val));
      const inp = h('input', { type: 'range', min, max, step, value: val, oninput: ev => { const v = +ev.target.value; out.textContent = fmtV(v); on(v); save(); draw(); } });
      return h('label', { class: 'sld' }, h('div', {}, h('span', {}, label), out), inp);
    };
    const pctF = v => nf1.format(v * 100) + ' %';
    function draw() {
      const scen = [['Försiktig', a.r - 0.03, 'c6'], ['Ditt antagande', a.r, 'c1'], ['Optimistisk', a.r + 0.03, 'c2']]
        .map(([n, r, cl]) => ({ n, r, cl, f: framskriv(o, r, a.s, 600, mal) }));
      const base = scen[1].f, horizon = Math.min(600, Math.max(36, (base.hit || 240) + 24));
      const dates = Array.from({ length: horizon + 1 }, (_, m) => { const d = new Date(o.d + 'T12:00:00'); d.setMonth(d.getMonth() + m); return d.toISOString().slice(0, 10); });
      const krav = kravSpar(o, a.r, mal, Math.round(daysBetween(o.d, a.ar + '-12-31') / 30.44));
      const steg = []; for (let v = Math.ceil(o.formogenhet / 5e5) * 5e5; v <= mal + 1 && steg.length < 8; v += 5e5) steg.push(v);
      box.replaceChildren(
        h('section', { class: 'hero' }, h('div', { class: 'lbl' }, `Målet ${mkr(mal)} nås`),
          h('div', { class: 'huge' }, o.formogenhet >= mal ? 'Nu!' : base.hit ? manadStr(o.d, base.hit) : 'inte inom 50 år'),
          h('div', { class: 'sub' }, base.hit ? `om ${tidStr(base.hit)} · med ${kr(a.s)}/mån och ${pctF(a.r)} per år` : 'Öka sparandet eller avkastningen.')),
        ctrl,
        card('Framskrivning', null, C.line(dates, scen.map(x => ({ name: `${x.n} (${pctF(x.r)})`, values: x.f.pts.slice(0, horizon + 1), cls: x.cl, dash: x.cl !== 'c1' }))
          .concat([{ name: 'Mål', values: dates.map(() => mal), cls: 'c4', dash: true }])),
          legend(scen.map(x => [x.cl, `${x.n}: ${x.f.hit ? manadStr(o.d, x.f.hit) : '–'}`]))),
        card('Delmål', null, h('table', { class: 'tbl' }, h('tbody', {}, steg.map(v => { const m = base.pts.findIndex(p => p >= v);
          return h('tr', {}, h('td', {}, mkr(v)), h('td', {}, m < 0 ? '–' : m === 0 ? 'nått' : manadStr(o.d, m)), h('td', { class: 'small' }, m > 0 ? 'om ' + tidStr(m) : '')); })))),
        card(`Nå målet till slutet av ${a.ar}`, null, h('p', {}, krav === null ? 'Välj ett senare år.' : krav === 0 ? 'Det klarar du utan att spara mer, om avkastningen håller.' :
          h('span', {}, 'Då behöver du spara ', h('b', {}, kr(krav) + '/mån'), ` med ${pctF(a.r)} avkastning per år.`)),
          slider('År', new Date().getFullYear() + 1, new Date().getFullYear() + 25, 1, a.ar, v => String(v), v => a.ar = v)));
    }
    const ref = a.hist ? `Din historiska avkastning ${a.hist.from}–${a.hist.to}: ${pctF(a.hist.r)} per år. Snittsparande senaste 12 mån: ${kr(snittSpar12())}/mån.` : '';
    const ctrl = card('Antaganden', null,
      slider('Avkastning per år', 0, 0.15, 0.005, a.r, pctF, v => a.r = v),
      slider('Sparande per månad', 0, 60000, 500, a.s, v => kr(v), v => a.s = v),
      h('p', { class: 'note' }, ref, ' Kassan antas stå still och sparandet läggs i portföljen. Prognosen är en räkneövning, inte ett löfte.'));
    draw();
    return h('div', { class: 'view' }, back(), box);
  }


  // ---------------------------------------------------------------- ISK-skatt och jämförelse med vanlig depå
  // Schablonränta (statslåneränta 30 nov året innan + 1 %, golv 1,25 % t.o.m. 2023) och skattefri grundnivå per inkomstår
  const ISK_RATE = { 2019: .0151, 2020: .0125, 2021: .0125, 2022: .0125, 2023: .0294, 2024: .0362, 2025: .0296, 2026: .0355 };
  const ISK_FRI = { 2025: 150000, 2026: 300000 };
  const iskRate = y => ISK_RATE[y] ?? ISK_RATE[Math.max(...Object.keys(ISK_RATE).map(Number))];
  const iskFri = y => y < 2025 ? 0 : (ISK_FRI[y] ?? 300000);
  const iskKnown = y => y in ISK_RATE;
  function iskPerAr() {
    const st = S.data.settings || {}, inclS = !!st.iskSavr, inclB = !!st.iskBank;
    const val = o => o.Aval + (inclS ? o.Sval : 0) + (inclB ? o.Dval : 0);
    const own = o => o.Aown + (inclS ? o.SinvG : 0) + (inclB ? o.Dinv : 0);
    const out = S.out, years = [...new Set(out.map(o => o.year))];
    let ackSkatt = 0, ackVinst = 0;
    return years.map(y => {
      const rows = out.filter(o => o.year === y), last = rows[rows.length - 1], prevRow = out[out.indexOf(rows[0]) - 1];
      const at = d => { let v = null; for (const o of out) { if (o.d <= d) v = o; else break; } return v || rows[0]; };
      const q = ['01-01', '04-01', '07-01', '10-01'].map(m => { const d = y + '-' + m; return d <= last.d ? val(at(d)) : val(last); });
      let ins = 0; rows.forEach(o => { const i = out.indexOf(o); if (i > 0) ins += Math.max(0, own(o) - own(out[i - 1])); });
      const ku = (q.reduce((a, b) => a + b, 0) + ins) / 4, rate = iskRate(y), fri = iskFri(y);
      const skatt = Math.max(0, ku - fri) * rate * 0.3;
      const start = prevRow || rows[0];
      const vinst = (val(last) - val(start)) - (own(last) - own(start));        // årets avkastning (inkl. utdelningar)
      const af = Math.max(0, vinst) * 0.3;                                         // skatt om vinsten tas ut på vanlig depå
      const pctAvk = ku ? vinst / ku : null, brytpunkt = ku ? skatt / 0.3 / ku : null;
      ackSkatt += skatt; ackVinst += vinst;
      // vanlig depå: skatt först vid försäljning – 30 % av hela vinsten sedan start (förluster räknas av)
      const afAck = Math.max(0, ackVinst) * 0.3;
      return { y, ku, rate, fri, skatt, vinst, af, pctAvk, brytpunkt, delar: rows.length && last.d.slice(5) < '12-20', known: iskKnown(y), last,
        ackSkatt, ackVinst, afAck, diffAck: afAck - ackSkatt };
    });
  }
  function viewIsk() {
    const ar = iskPerAr(), cur = ar[ar.length - 1], st = S.data.settings || {}, first = ar[0];
    const tog = (key, label) => h('label', { class: 'acc' }, h('input', { type: 'checkbox', checked: st[key] ? true : null, onchange: ev => commit(d => { d.settings = d.settings || {}; d.settings[key] = ev.target.checked; }, null) }), h('span', {}, label));
    return h('div', { class: 'view' }, h('a', { class: 'back', href: '#avkastning' }, '‹ Avkastning'),
      h('section', { class: 'hero' }, h('div', { class: 'lbl' }, `ISK-skatt ${cur.y}${cur.delar ? ' (beräknad hittills)' : ''}`), h('div', { class: 'huge' }, kr(cur.skatt)),
        h('div', { class: 'sub' }, `kapitalunderlag ${kr(cur.ku)} · ${nf2.format(cur.rate * 100)} % × 30 %` + (cur.fri ? ` · ${mkr(cur.fri)} skattefritt` : '') + (cur.known ? '' : ' · prognos'))),
      card('Om du sålde allt idag', null,
        h('div', { class: 'grid2' },
          stat('Vinst sedan ' + first.y, krS(cur.ackVinst), null, cls(cur.ackVinst)),
          stat('ISK-skatt betald totalt', kr(cur.ackSkatt), `${ar.length} år, varje år oavsett resultat`),
          stat('Skatt på vanlig depå', kr(cur.afAck), '30 % av vinsten, en gång vid försäljning'),
          stat('Brytpunkt ' + cur.y, pct(cur.brytpunkt, false), `i år: ${pct(cur.pctAvk)}`)),
        h('div', { class: 'verdict ' + (cur.diffAck >= 0 ? 'pos' : 'neg') }, cur.diffAck >= 0 ? `ISK har sparat dig ${kr(cur.diffAck)}` : `Vanlig depå hade varit ${kr(-cur.diffAck)} billigare`),
        h('p', { class: 'note' }, `Brytpunkten är den avkastning där årets ISK-skatt blir lika stor som 30 % av vinsten. Över den lönar sig ISK det året.`)),
      card('Över tid', null, h('div', { class: 'scrollx' }, h('table', { class: 'tbl tight' }, h('thead', {}, h('tr', {}, ['År', 'Avk.', 'ISK-skatt', 'ISK tot.', 'Depå*', 'ISK ger'].map(x => h('th', {}, x)))),
        h('tbody', {}, ar.slice().reverse().map(x => h('tr', {}, h('td', {}, x.y + (x.delar ? '·' : '')), h('td', { class: cls(x.pctAvk) }, pct(x.pctAvk)), h('td', {}, nf0.format(x.skatt)),
          h('td', {}, nf0.format(x.ackSkatt)), h('td', {}, nf0.format(x.afAck)), h('td', { class: cls(x.diffAck) }, (x.diffAck > 0 ? '+' : '') + nf0.format(x.diffAck))))))),
        h('p', { class: 'note' }, 'Avk. = årets avkastning. ISK tot. = summan av alla års ISK-skatt. * Depå = skatten om du hade sålt allt vid årets slut: 30 % av all vinst sedan start, förluster avräknade. · = pågående år.')),
      card('Vilka konton är ISK eller kapitalförsäkring?', null, h('p', { class: 'note' }, 'Avanza räknas alltid med (ISK och kapitalförsäkring). Konton du stängt av ingår inte.'),
        tog('iskSavr', 'SAVR är ISK'), tog('iskBank', `${bankNamn()} är ISK`)),
      card('Så räknas det', null, h('ul', { class: 'note' },
        h('li', {}, 'ISK: skatt varje år på kapitalunderlaget, oavsett om det gått upp eller ned. Kapitalunderlag = (värdet vid varje kvartals början + årets insättningar) / 4.'),
        h('li', {}, 'Skatt = (kapitalunderlag − skattefri grundnivå) × schablonränta × 30 %. Schablonräntan är statslåneräntan 30 november året innan + 1 %.'),
        h('li', {}, 'Vanlig depå: 30 % skatt på vinsten när du säljer. Förluster räknas av mot vinster. Utdelningar beskattas egentligen varje år på en depå, och pengar som betalats i ISK-skatt har inte kunnat växa vidare – båda effekterna är små och räknas inte med.'),
        h('li', {}, 'Vinsten räknas från första veckan i appen. Den skattefria nivån gäller per person för ISK och kapitalförsäkring tillsammans. Allt är uppskattningar – den riktiga skatten står i deklarationen.'))));
  }



  // ---------------------------------------------------------------- årssammanfattning (utskrivbar / PDF)
  function viewAr(yq) {
    if (!AV.loaded) avLoad().then(() => { if (location.hash.startsWith('#ar')) render(); });
    const out = S.out, years = [...new Set(out.map(o => o.year))], y = +(yq || years[years.length - 1]);
    const rows = out.filter(o => o.year === y); if (!rows.length) return h('div', { class: 'view' }, back(), card('Årssammanfattning', null, h('p', {}, 'Inga veckor finns för ' + y + '.')));
    const o = rows[rows.length - 1], i0 = out.indexOf(rows[0]), start = i0 > 0 ? out[i0 - 1] : rows[0];
    const yrs = Calc.years(out), prev = yrs.find(x => x.year === y - 1), mal = (S.data.settings || {}).mal || 0;
    const pagaende = o.d.slice(5) < '12-20';
    const forandr = o.formogenhet - start.formogenhet;
    // avkastning per månad
    const man = Array(12).fill(null); rows.forEach(r => { const m = +r.d.slice(5, 7) - 1; man[m] = (man[m] || 0) + r.W; });
    const isk = iskPerAr().find(x => x.y === y);
    const big = (l, v, sub, c) => h('div', { class: 'arstat' }, h('div', { class: 'lbl' }, l), h('div', { class: 'big ' + (c || '') }, v), sub ? h('div', { class: 'sub' }, sub) : null);
    const jmf = (v, p, f = kr) => p === undefined || p === null ? null : `${y - 1}: ${f(p)}`;
    // innehav (från Avanza-transaktionerna)
    let innehav = null;
    if (AV.an) {
      const g = AV.an.groups;
      const utd = g.map(x => ({ n: x.name, v: x.dy[y] || 0 })).filter(x => x.v > 0).sort((a, b) => b.v - a.v);
      const real = g.map(x => ({ n: x.name, v: x.ry[y] || 0 })).filter(x => Math.abs(x.v) > 1).sort((a, b) => b.v - a.v);
      const sn = avSnaps().filter(x => x.date.slice(0, 4) === String(y)), snPrev = avSnaps().filter(x => x.date < y + '-01-08').pop();
      let bs = null;
      if (sn.length && (snPrev || sn.length > 1)) { const A = snPrev || sn[0], B = sn[sn.length - 1]; if (A !== B) bs = { A, B, l: Avanza.perf(AV.data.tx, AV.an, A, B) }; }
      const lst = (title, arr, n = 5) => arr.length ? h('div', {}, h('h3', {}, title), h('table', { class: 'tbl small' }, h('tbody', {}, arr.slice(0, n).map(x => h('tr', {}, h('td', {}, x.n), h('td', { class: cls(x.v) }, krS(x.v))))))) : null;
      innehav = h('section', { class: 'card' }, h('h2', {}, 'Innehav på Avanza'),
        bs ? h('div', { class: 'cols2' }, lst('Bäst', bs.l.filter(x => x.kr > 0).map(x => ({ n: x.name, v: x.kr }))), lst('Sämst', bs.l.filter(x => x.kr < 0).reverse().map(x => ({ n: x.name, v: x.kr })))) : null,
        bs ? h('p', { class: 'note' }, `Bäst/sämst räknat ${dstr(bs.A.date)} – ${dstr(bs.B.date)}.`) : null,
        h('div', { class: 'cols2' }, lst('Största utdelningarna', utd), lst('Realiserade resultat', real.filter(x => x.v > 0).concat(real.filter(x => x.v < 0).reverse().slice(0, 2)), 6)));
    }
    const knappar = h('div', { class: 'noprint arbar' },
      h('div', { class: 'chips small' }, years.slice().reverse().map(yy => h('a', { class: 'chipa' + (yy === y ? ' on' : ''), href: '#ar?y=' + yy }, String(yy)))),
      h('button', { class: 'primary', onclick: () => window.print() }, 'Spara som PDF / skriv ut'));
    return h('div', { class: 'view arsida' }, h('a', { class: 'back noprint', href: '#avkastning' }, '‹ Avkastning'), knappar,
      h('section', { class: 'arhead' }, h('div', { class: 'lbl' }, 'Årssammanfattning'), h('h1', {}, String(y) + (pagaende ? ' (hittills)' : '')),
        h('div', { class: 'sub' }, `${dstr(rows[0].d)} – ${dstr(o.d)} · skapad ${dstr(today())}`)),
      h('section', { class: 'card' }, h('div', { class: 'argrid' },
        big('Förmögenhet vid ' + (pagaende ? 'senaste avläsning' : 'årets slut'), kr(o.formogenhet), `${krS(forandr)} under året`),
        big('Avkastning', pct(o.Y), `${krS(o.U)} · pengaviktad ${pct(o.O)}`, cls(o.Y)),
        big('Sparat', kr(o.sparatAr), prev ? jmf(o.sparatAr, prev.sparatAr) : '', cls(o.sparatAr)),
        big('Sparkvot', o.sparkvotAr === null ? '–' : pct(o.sparkvotAr, false), o.lonAr ? `av ${kr(o.lonAr)} nettolön` : ''),
        big('Utdelningar', kr(o.utdAr + o.aterAr), prev ? jmf(0, prev.utdAr + prev.aterAr) : ''),
        mal ? big('Mot målet ' + mkr(mal), pct(o.formogenhet / mal, false), `${kr(Math.max(0, mal - o.formogenhet))} kvar`) : null)),
      h('section', { class: 'card' }, h('h2', {}, 'Konton'), h('table', { class: 'tbl' }, h('thead', {}, h('tr', {}, ['', 'Värde', 'Utveckling', 'Avkastat'].map(x => h('th', {}, x)))),
        h('tbody', {}, [['Avanza', o.Aval, o.avY, (o.Aval - o.Aown) - (start.Aval - start.Aown)], ['SAVR', o.Sval, o.svY, (o.Sval - o.SinvG) - (start.Sval - start.SinvG)],
          [bankNamn(), o.Dval, o.dbY, (o.Dval - o.Dinv) - (start.Dval - start.Dinv)], ['Kassa', o.kassa, null, null]]
          .map(([n, v, p, a]) => h('tr', {}, h('td', {}, n), h('td', {}, kr(v)), h('td', { class: cls(p) }, p === null ? '' : pct(p)), h('td', { class: cls(a) }, a === null ? '' : krS(a))))))),
      h('section', { class: 'card' }, h('h2', {}, 'Förmögenhet under året'), C.line(rows.map(r => r.d), [{ name: 'Förmögenhet', values: rows.map(r => r.formogenhet), cls: 'c1', fill: true }, { name: 'Egna insatta pengar', values: rows.map(r => r.K + r.kassa), cls: 'c2', dash: true }], { zero: false }),
        legend([['c1', 'Förmögenhet'], ['c2', 'Egna pengar + kassa']])),
      h('section', { class: 'card' }, h('h2', {}, 'Avkastning per månad'), C.bars(MONL.map(m => m.slice(0, 3)), [{ name: 'Avkastning', values: man }], { labels: false })),
      innehav,
      isk ? h('section', { class: 'card' }, h('h2', {}, 'ISK-skatt'), h('p', {}, `Beräknad skatt ${kr(isk.skatt)} på kapitalunderlag ${kr(isk.ku)}. `,
        `ISK-skatt totalt sedan ${iskPerAr()[0].y}: ${kr(isk.ackSkatt)}. Hade allt sålts vid årets slut på en vanlig depå hade skatten blivit ${kr(isk.afAck)} – `,
        isk.diffAck >= 0 ? `ISK har sparat ${kr(isk.diffAck)}.` : `vanlig depå hade varit ${kr(-isk.diffAck)} billigare.`)) : null,
      prev ? h('section', { class: 'card' }, h('h2', {}, `Jämfört med ${y - 1}`), h('table', { class: 'tbl' }, h('thead', {}, h('tr', {}, ['', String(y - 1), String(y)].map(x => h('th', {}, x)))),
        h('tbody', {}, [['Avkastning', pct(prev.Y), pct(o.Y)], ['Avkastat', kr(prev.U), kr(o.U)], ['Sparat', kr(prev.sparatAr), kr(o.sparatAr)], ['Sparkvot', prev.sparkvotAr === null ? '–' : pct(prev.sparkvotAr, false), o.sparkvotAr === null ? '–' : pct(o.sparkvotAr, false)],
          ['Utdelningar', kr(prev.utdAr + prev.aterAr), kr(o.utdAr + o.aterAr)], ['Förmögenhet', kr(prev.formogenhet), kr(o.formogenhet)]].map(r => h('tr', {}, r.map(c => h('td', {}, c))))))) : null);
  }

  // ---------------------------------------------------------------- detaljsida: avkastning
  function viewAvkastning() {
    const out = S.out, o = out[out.length - 1], sl = slice(), y = yearsDone();
    return h('div', { class: 'view' }, back(),
      h('section', { class: 'hero' }, h('div', { class: 'lbl' }, 'Avkastning'), h('div', { class: 'huge ' + cls(o.V) }, krS(o.V)), h('div', { class: 'sub' }, `totalt sedan start · ${pct(o.P)} på egna pengar`)),
      card(null, null, h('div', { class: 'grid2' },
        stat(periodLabel(out.length - 1), pct(o.X), krS(o.W), cls(o.X)), stat('I år (tidsviktad)', pct(o.Y), krS(o.U), cls(o.Y)),
        stat('I år (pengaviktad)', pct(o.O), 'på insatt kapital', cls(o.O)), stat('Totalt', pct(o.P), krS(o.V), cls(o.P)))),
      card('Utveckling i år', null, G.iAr()),
      card('Avkastning per vecka', null, h('div', { class: 'row end' }, rangeSeg()), G.vecka(sl)),
      card('Total avkastning', null, G.total(sl)),
      card('Utveckling på egna pengar', null, G.utv(sl)),
      card('Per år', null, G.perAr(), h('table', { class: 'tbl' }, h('thead', {}, h('tr', {}, ['År', 'Tidsviktad', 'Pengaviktad', 'Avkastat'].map(x => h('th', {}, x)))),
        h('tbody', {}, y.slice().reverse().map(x => h('tr', {}, h('td', {}, x.label), h('td', { class: cls(x.Y) }, pct(x.Y)), h('td', { class: cls(x.O) }, pct(x.O)), h('td', {}, kr(x.U))))))),
      h('div', { class: 'links' }, h('a', { href: '#isk' }, 'ISK-skatt och jämförelse med vanlig depå ›'), h('a', { href: '#ar' }, 'Årssammanfattning ›')));
  }

  // ---------------------------------------------------------------- detaljsida: konto
  const KONTO = {
    avanza: { name: () => 'Avanza', val: o => o.Aval, own: o => o.Aown, y: o => o.avY, tot: o => o.avTot, w: o => o.avW, x: o => o.avX },
    savr: { name: () => 'SAVR', val: o => o.Sval, own: o => o.SinvG, y: o => o.svY, tot: o => o.svTot, w: o => o.svW, x: o => o.svX },
    bank: { name: () => bankNamn(), val: o => o.Dval, own: o => o.Dinv, y: o => o.dbY, tot: o => o.dbTot, w: o => o.dbW, x: o => o.dbX },
    kassa: { name: () => 'Kassa', val: o => o.kassa },
  };
  function viewKonto(k) {
    const K = KONTO[k] || KONTO.avanza, out = S.out, o = out[out.length - 1], sl = slice(), nm = K.name();
    if (k === 'kassa') return h('div', { class: 'view' }, back(),
      h('section', { class: 'hero' }, h('div', { class: 'lbl' }, 'Kassa'), h('div', { class: 'huge' }, kr(o.kassa)), h('div', { class: 'sub' }, 'lönekonto, buffert och sparkonto')),
      card('Kassa över tid', null, h('div', { class: 'row end' }, rangeSeg()), C.line(sl.map(x => x.d), [{ name: 'Kassa', values: sl.map(x => x.kassa), cls: 'c3', fill: true }])),
      card('Andel av förmögenheten', null, C.line(sl.map(x => x.d), [{ name: 'Andel', values: sl.map(x => x.formogenhet ? x.kassa / x.formogenhet : null), cls: 'c3' }], { kind: 'pct' })));
    const yrs = yearsDone().filter(y => (K.val(y) || 0) > 0 || K.y(y));
    const avk = K.val(o) - K.own(o);
    return h('div', { class: 'view' }, back(),
      h('section', { class: 'hero' }, h('div', { class: 'lbl' }, nm), h('div', { class: 'huge' }, kr(K.val(o))), h('div', { class: 'sub' }, h('span', { class: cls(K.y(o)) }, pct(K.y(o))), ' i år')),
      card(null, null, h('div', { class: 'grid2' },
        stat(periodLabel(out.length - 1), pct(K.x(o)), krS(K.w(o)), cls(K.x(o))), stat('I år', pct(K.y(o)), null, cls(K.y(o))),
        stat('Totalt', pct(K.tot(o)), null, cls(K.tot(o))), stat('Avkastat totalt', krS(avk), null, cls(avk)),
        stat('Egna insatta pengar', kr(K.own(o))),
        k === 'avanza' ? stat('Veckan exkl. utdelningar', krS(o.avW - o.utd - o.avg), 'jämförbar med Avanzas "Utveckling 1 vecka"', cls(o.avW - o.utd - o.avg)) : null)),
      k === 'avanza' ? card('Fördelning', null, C.donut([{ name: 'Aktier', value: o.aav, cls: 'c1' }, { name: 'Fonder', value: o.afv, cls: 'c3' }, { name: 'ETF', value: o.aev, cls: 'c2' }], { center: [C.mkr(o.Aval), 'Avanza'] }),
        h('table', { class: 'tbl' }, h('thead', {}, h('tr', {}, ['', 'Värde', 'I år'].map(x => h('th', {}, x)))), h('tbody', {},
          [['Aktier', o.aav, o.aktY], ['Fonder', o.afv, o.fondY], ['ETF', o.aev, o.etfY]].map(([a, v, y]) => h('tr', {}, h('td', {}, a), h('td', {}, kr(v)), h('td', { class: cls(y) }, pct(y)))))),
        h('div', { class: 'btns' }, h('a', { class: 'ghost btn', href: '#innehav' }, 'Visa innehav, GAV och utdelningar'))) : null,
      card('Värde och insatta pengar', null, h('div', { class: 'row end' }, rangeSeg()),
        C.line(sl.map(x => x.d), [{ name: 'Värde', values: sl.map(K.val), cls: 'c1', fill: true }, { name: 'Egna insatta pengar', values: sl.map(K.own), cls: 'c2', dash: true }]),
        legend([['c1', 'Värde'], ['c2', 'Egna insatta pengar']])),
      card('Avkastat över tid', null, C.timeBars(sl.map(x => x.d), sl.map(x => K.val(x) - K.own(x)), { name: 'Avkastat', cls: 'c1' })),
      card('Utveckling per år', null, C.bars(yrs.map(y => y.label), [{ name: nm, values: yrs.map(K.y) }], { kind: 'pct', labels: true })));
  }

  // ---------------------------------------------------------------- detaljsida: sparande
  function viewSparande() {
    const out = S.out, o = out[out.length - 1], y = yearsDone();
    const lastLon = out.filter(x => x.lon !== null && x.manKvot !== null && x.manKvot !== undefined).pop();
    const fjolSp = sameTimeLastYear('sparatAr'), fjolUtd = sameTimeLastYear(utdYtd);
    const lonRader = out.filter(x => x.year === o.year && x.lon !== null).reverse();
    return h('div', { class: 'view' }, back(),
      h('section', { class: 'hero' }, h('div', { class: 'lbl' }, 'Sparat ' + o.year), h('div', { class: 'huge' }, kr(o.sparatAr)),
        h('div', { class: 'sub' }, fjolSp !== null ? `i fjol vid denna tid ${kr(fjolSp)}` : '')),
      card(null, null, h('div', { class: 'grid2' },
        stat('Snittspar per månad', o.snittMan === null ? '–' : kr(o.snittMan), o.lonAntal ? `${o.lonAntal} löner i år` : ''),
        stat('Sparkvot i år', o.sparkvotAr === null ? '–' : pct(o.sparkvotAr, false), o.lonAntal ? `av ${kr(o.lonAr)} nettolön` : ''),
        stat('Senaste lönen', lastLon ? pct(lastLon.manKvot, false) : '–', lastLon ? `${kr(lastLon.manSparat)} av ${kr(lastLon.lon)}` : ''),
        stat('Utdelningar i år', kr(utdYtd(o)), fjolUtd !== null ? `i fjol vid denna tid ${kr(fjolUtd)}` : ''))),
      card('Nettoårsinkomst och sparkvot', null, G.inkomst()),
      card('Sparande år för år', null, yearCompare('spar', x => x.sparatAr)),
      card('Utdelningar per år', null, C.bars(y.map(x => x.label), [{ name: 'Utdelningar', values: y.map(utdYtd), cls: 'c2' }], { labels: true })),
      card('Utdelningar år för år', null, yearCompare('utd', utdYtd)),
      lonRader.length ? card('Per lön ' + o.year, null, h('table', { class: 'tbl' }, h('thead', {}, h('tr', {}, ['Lönevecka', 'Nettolön', 'Sparat', 'Sparkvot'].map(x => h('th', {}, x)))),
        h('tbody', {}, lonRader.map(x => h('tr', {}, h('td', {}, dstr(x.d).replace(/ \d{4}$/, '')), h('td', {}, nf0.format(x.lon)),
          h('td', { class: cls(x.manSparat) }, x.manSparat === null ? '–' : nf0.format(x.manSparat)), h('td', {}, x.manKvot === null ? '–' : pct(x.manKvot, false))))))) : null,
      card('Per år', null, h('table', { class: 'tbl' }, h('thead', {}, h('tr', {}, ['År', 'Nettoinkomst', 'Sparat', 'Kvot', 'Utdeln.'].map(x => h('th', {}, x)))),
        h('tbody', {}, y.slice().reverse().map(x => h('tr', {}, h('td', {}, x.label), h('td', {}, x.lonAr ? nf0.format(x.lonAr) : '–'), h('td', { class: cls(x.sparatAr) }, nf0.format(x.sparatAr)),
          h('td', {}, x.sparkvotAr === null ? '–' : pct(x.sparkvotAr, false)), h('td', {}, nf0.format(utdYtd(x)))))))));
  }

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
  function senastInlagd(rows) {
    const last = rows.length ? rows[rows.length - 1].d : null; if (!last) return null;
    const dd = daysBetween(last, today()), WD = ['sön', 'mån', 'tis', 'ons', 'tors', 'fre', 'lör'];
    return h('div', { class: 'lastin' + (dd >= 8 ? ' late' : '') }, h('span', {}, 'Senast inlagd'),
      h('b', {}, `${WD[new Date(last + 'T12:00:00').getDay()]} ${dstr(last)}`), h('small', {}, dd === 0 ? 'idag' : dd === 1 ? 'igår' : `${dd} dagar sedan`));
  }
  const shiftYears = (d, y) => (+d.slice(0, 4) + y) + d.slice(4);

  const FIELDS = () => [
    ['Kassa', [['kassa', 'Lönekonto + buffert', 'value'], ['kassa2', 'Avanza sparkonto', 'opt']]],
    ['Avanza – värde', [['av.av', 'Aktier', 'value'], ['av.fv', 'Fonder', 'value'], ['av.ev', 'ETF', 'value']]],
    ['Avanza – insatt denna vecka', [['av.ai', 'Aktier', 'flow'], ['av.fi', 'Fonder', 'flow'], ['av.ei', 'ETF', 'flow'], ['av.utd', 'Utdelningar (efter källskatt)', 'flow'], ['av.avg', 'Skatt, avgifter & ränta', 'flow']]],
    ['SAVR', [['sv.val', 'Fondvärde', 'value'], ['sv.inv', 'Insatt denna vecka (egna pengar)', 'flow'], ['sv.ater', 'Återbäring', 'flow']]],
    [bankNamn(), [['db.val', 'Fondvärde', 'value'], ['db.inv', 'Insatt denna vecka', 'flow']]],
  ];
  const ADV = () => [['av.vinst', 'Avanza – realiserad vinst (behövs inte om insatt = köp − sälj)'], ['db.vinst', bankNamn() + ' – realiserad vinst vid försäljning']];
  const getP = (o, p) => p.split('.').reduce((a, k) => a == null ? undefined : a[k], o);
  const setP = (o, p, v) => { const k = p.split('.'); let a = o; for (let i = 0; i < k.length - 1; i++) a = a[k[i]] = a[k[i]] || {}; a[k[k.length - 1]] = v; };

  function viewForm(editDate) {
    const rows = S.data ? Calc.sortRows(S.data.rows) : [];
    const editing = editDate ? rows.find(r => r.d === editDate) : null;
    const prev = editing ? rows[rows.indexOf(editing) - 1] : rows[rows.length - 1];
    const form = h('form', { class: 'view form', onsubmit: ev => { ev.preventDefault(); submit(); } });
    const inputs = {};
    const fmtIn = v => (v === undefined || v === null || v === 0) ? '' : nf2.format(v);
    const field = (path, label, kind) => {
      const inp = h('input', { inputmode: 'decimal', enterkeyhint: 'next', autocomplete: 'off', placeholder: kind === 'value' ? '' : '0', value: editing ? fmtIn(getP(editing, path)) : '', 'data-kind': kind,
        oninput: ev => { delete ev.target.dataset.auto; preview(); }, onkeydown: ev => { if (ev.key === 'Enter') { ev.preventDefault(); nextEmpty(inp); } } });
      inputs[path] = inp; labels[path] = label;
      return h('label', { class: 'fld' }, h('span', {}, label), inp);
    };
    const labels = {}; let tried = false;
    function applyHoldings(res) {
      const t = res.totals, ex = ((S.data.settings || {}).avanzaEx) || [], x = AV.an ? Avanza.excludedValues(AV.an, AV.data.holdings, ex) : { aktie: 0, fond: 0, etf: 0 };
      inputs['av.av'].value = fmtIn(Math.round(t.aktie - x.aktie)); inputs['av.fv'].value = fmtIn(Math.round(t.fond - x.fond)); inputs['av.ev'].value = fmtIn(Math.round(t.etf - x.etf));
      ['av.av', 'av.fv', 'av.ev'].forEach(p => delete inputs[p].dataset.auto);
      if (x.aktie + x.fond + x.etf) toast(`${kr(x.aktie + x.fond + x.etf)} på avstängda konton är borträknat`, 'ok');
    }
    // Enter / Nästa: hoppa till nästa synliga tomma fält (eller Spara)
    function nextEmpty(from) {
      const all = [...form.querySelectorAll('input[inputmode=decimal]')].filter(i => i.offsetParent !== null);
      const k = all.indexOf(from), nx = all.slice(k + 1).find(i => i.value.trim() === '') || null;
      if (nx) { nx.focus(); nx.scrollIntoView({ block: 'center', behavior: 'smooth' }); } else { from.blur(); form.querySelector('button[type=submit]').scrollIntoView({ block: 'center', behavior: 'smooth' }); }
    }
    const date = h('input', { type: 'date', value: editing ? editing.d : today(), max: today(), required: true, oninput: () => { autoFill(false); preview(); } });
    let kassa2Row = null;
    const autoInfo = h('div', { class: 'auto' });
    const csvIn = h('input', { type: 'file', accept: '.csv,text/csv', style: 'display:none', onchange: async ev => { await avImportCSV(ev.target.files[0]); autoFill(false); preview(); } });
    // Avanza: fyll i insatt/utdelningar från importerade transaktioner (fält du själv ändrat lämnas orörda)
    function autoFill(force) {
      if (!AV.loaded || !AV.an || !date.value) { renderAuto(null); return; }
      const prevD = rows.filter(x => x.d < date.value && x !== editing).map(x => x.d).pop() || shiftDays(date.value, -7);
      const ex = ((S.data.settings || {}).avanzaEx) || [];
      const f = Avanza.weekFlows(AV.data.tx, AV.an, AV.data.holdings, prevD, date.value, ex);
      const covered = !!(AV.data.coverTo && AV.data.coverTo >= date.value);
      if (!editing || force) {
        [['av.ai', f.aktie], ['av.fi', f.fond], ['av.ei', f.etf], ['av.utd', f.utd], ['av.avg', f.avg], ['kassa2', f.sparkonto]].forEach(([p, v]) => {
          const inp = inputs[p]; if (!inp) return;
          if (force || inp.value.trim() === '' || inp.dataset.auto !== undefined) { inp.value = v ? fmtIn(v) : ''; inp.dataset.auto = covered ? '1' : '0'; }
        });
      }
      renderAuto({ f, prevD, covered });
    }
    function renderAuto(x) {
      autoInfo.innerHTML = '';
      const last = AV.an ? AV.an.lastDate : null;
      if (x) autoInfo.append(h('p', { class: 'note' }, `Ifyllt från Avanza-transaktionerna ${dstr(x.prevD)} – ${dstr(date.value)} (${x.f.n} st). Senaste transaktion i filen: ${last ? dstr(last) : '–'}. Du kan ändra värdena.`),
        ...(!x.covered ? [h('p', { class: 'warn' }, `Transaktionsfilen gäller till ${AV.data.coverTo ? dstr(AV.data.coverTo) : 'okänt datum'} – importera en ny export för att få med hela veckan.`)] : []));
      else autoInfo.append(h('p', { class: 'note' }, AV.loaded ? 'Fyll i själv, eller importera Avanzas transaktionsexport så fylls fälten i automatiskt.' : 'Hämtar Avanza-transaktioner…'));
      autoInfo.append(csvIn, h('div', { class: 'btns' }, h('button', { type: 'button', class: 'ghost small', onclick: () => csvIn.click() }, 'Importera transaktioner (CSV)'),
        x && editing ? h('button', { type: 'button', class: 'ghost small', onclick: () => { autoFill(true); preview(); } }, 'Fyll i från transaktionerna') : null));
    }
    const yearEnd = h('input', { type: 'checkbox', onchange: () => preview() });
    const yearEndBox = h('label', { class: 'chk', style: 'display:none' }, yearEnd, h('span', {}, ''));
    form.append(h('section', { class: 'card' }, h('h2', {}, editing ? 'Ändra vecka' : 'Ny vecka'),
      editing ? null : senastInlagd(rows),
      editing ? null : h('button', { type: 'button', class: 'primary wide wizbtn', onclick: () => wizard() }, 'Lägg in veckan steg för steg'),
      h('label', { class: 'fld' }, h('span', {}, 'Datum för avläsningen'), date),
      h('p', { class: 'note' }, 'Använd dagens datum, även om du är sen – värdena gäller den dag du läser av dem.'), yearEndBox));
    FIELDS().forEach(([title, fl]) => {
      const sec = h('section', { class: 'card' }, h('h3', {}, title), fl.map(f => field(...f)));
      if (title === 'Avanza – värde') {
        let box = null;
        const btn = h('button', { type: 'button', class: 'linkbtn', onclick: () => {
          if (box) { box.remove(); box = null; return; }
          box = pasteBox(res => { applyHoldings(res); box.remove(); box = null; preview(); },
            'Värdena för aktier, fonder och ETF fylls i (konton du stängt av räknas bort), och kurserna under Innehav uppdateras.');
          sec.append(box); } }, 'Klistra in innehavslistan från Avanza');
        sec.append(btn);
      }
      if (title === 'Avanza – insatt denna vecka') sec.append(autoInfo);
      if (title === 'Kassa') kassa2Row = inputs.kassa2.closest('label');
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
      Object.values(inputs).forEach(i => i.classList.toggle('bad', tried && bad.includes(i)));
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
      if (kassa2Row) kassa2Row.style.display = (inputs.kassa2.value.trim() && parseNum(inputs.kassa2.value) !== 0) ? '' : 'none';
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
      tried = true;
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

    // ------------------------------------------------ steg för steg (som makrot): en fråga i taget, hoppar över det som redan är ifyllt
    function wizard() {
      if (document.querySelector('.wiz')) return;            // bara en ruta åt gången
      const lonVisas = lonSec.style.display !== 'none';
      const plan = [
        ['kassa', 'Kassa', 'Hur mycket har du på lönekonto + buffert?'],
        ['av.av', 'Avanza', 'Vad är aktievärdet på Avanza?', 'paste'], ['av.fv', 'Avanza', 'Vad är fondvärdet på Avanza?', 'paste'], ['av.ev', 'Avanza', 'Vad är ETF-värdet på Avanza?', 'paste'],
        ['av.ai', 'Avanza', 'Hur mycket har du investerat i aktier denna vecka?'], ['av.fi', 'Avanza', 'Hur mycket har du investerat i fonder denna vecka?'],
        ['av.ei', 'Avanza', 'Hur mycket har du investerat i ETF:er denna vecka?'], ['av.utd', 'Avanza', 'Hur mycket utdelning har kommit in?'],
        ['sv.val', 'SAVR', 'Vad är fondvärdet i SAVR?'], ['sv.inv', 'SAVR', 'Hur mycket har du satt in i SAVR (egna pengar)?'],
        ['db.val', bankNamn(), `Vad är fondvärdet i ${bankNamn()}?`], ['db.inv', bankNamn(), `Hur mycket har du satt in i fonder i ${bankNamn()}?`],
      ].concat(lonVisas ? [['lon', 'Lön', 'Lönen har kommit – vad blev nettolönen?']] : []);
      // hoppa över fält som redan har ett värde (inklistrat, importerat eller skrivet)
      const steps = plan.filter(([p]) => inputs[p] && inputs[p].value.trim() === '' && inputs[p].dataset.auto !== '1');
      let i = 0, pasteMode = false;
      const boxEl = h('div', { class: 'wizbox', role: 'dialog', 'aria-modal': 'true' }), ov = h('div', { class: 'wiz' }, boxEl);
      const close = () => { ov.remove(); document.body.classList.remove('noscroll'); document.removeEventListener('keydown', esc); preview(); };
      const esc = ev => { if (ev.key === 'Escape') close(); };
      document.addEventListener('keydown', esc);
      ov.addEventListener('click', ev => { if (ev.target === ov) close(); });
      const go = d => { i = Math.max(0, Math.min(steps.length, i + d)); pasteMode = false; show(); };
      function frame(title, body, foot) {
        boxEl.replaceChildren(
          h('div', { class: 'wizhead' },
            h('div', {}, h('div', { class: 'wizstep' }, i < steps.length ? `Steg ${i + 1} av ${steps.length}` : 'Klart'), h('div', { class: 'wiztitle' }, title)),
            h('button', { type: 'button', class: 'wizx', onclick: close, 'aria-label': 'Stäng' }, '✕')),
          h('div', { class: 'wizbar' }, h('i', { style: `width:${(steps.length ? Math.min(i, steps.length) / steps.length : 1) * 100}%` })),
          h('div', { class: 'wizbody' }, ...body.filter(Boolean)),
          h('div', { class: 'wizfoot' }, ...foot.filter(Boolean)));
      }
      function show() {
        if (i >= steps.length) {
          preview();
          const filled = plan.filter(([p]) => inputs[p] && inputs[p].value.trim() !== '');
          frame('Sammanfattning', [
            h('div', { class: 'preview big' }, ...[...pv.childNodes].map(n => n.cloneNode(true))),
            filled.length ? h('table', { class: 'kv wizsum' }, h('tbody', {}, filled.map(([p, sec, q]) => h('tr', {}, h('td', {}, `${sec}: ${labels[p] || q}`), h('td', {}, inputs[p].value))))) : null,
            steps.length ? null : h('p', { class: 'note' }, 'Allt var redan ifyllt.')],
            [steps.length ? h('button', { type: 'button', class: 'ghost', onclick: () => go(-1) }, '‹ Tillbaka') : null,
             h('button', { type: 'button', class: 'ghost', onclick: close }, 'Granska'),
             h('button', { type: 'button', class: 'primary', onclick: () => { close(); form.requestSubmit(); } }, 'Spara veckan')]);
          return;
        }
        const [p, sec, q, extra] = steps[i], src = inputs[p], kind = src.dataset.kind;
        if (pasteMode) {
          const ta = h('textarea', { rows: 7, placeholder: 'Markera hela innehavslistan på Avanza (Min ekonomi › Innehav), kopiera och klistra in här.' });
          const read = async () => {
            try { const res = await avSaveHoldings(ta.value); applyHoldings(res); pasteMode = false;
              const left = steps.slice(i).filter(([pp]) => inputs[pp].value.trim() === ''); steps.splice(i, steps.length - i, ...left); show(); }
            catch (e) { toast(e.message, 'err'); }
          };
          frame('Avanza', [h('div', { class: 'wizq' }, 'Klistra in innehavslistan'), h('p', { class: 'note' }, 'Aktier, fonder och ETF fylls i på en gång. Konton du stängt av räknas bort.'), ta],
            [h('button', { type: 'button', class: 'ghost', onclick: () => { pasteMode = false; show(); } }, '‹ Avbryt'), h('button', { type: 'button', class: 'primary', onclick: read }, 'Läs in')]);
          setTimeout(() => ta.focus(), 30);
          return;
        }
        const inp = h('input', { inputmode: 'decimal', enterkeyhint: 'next', autocomplete: 'off', value: src.value, placeholder: kind === 'value' ? '' : '0' });
        const err = h('p', { class: 'warn' });
        const next = () => {
          const v = inp.value.trim();
          if (kind === 'value' && v === '') { err.textContent = 'Fyll i ett värde.'; inp.focus(); return; }
          if (v !== '' && !isFinite(parseNum(v))) { err.textContent = 'Skriv ett belopp, t.ex. 1 234,50 eller 1000+250.'; inp.focus(); return; }
          src.value = v === '' ? '' : fmtIn(parseNum(v)); delete src.dataset.auto; go(1);
        };
        inp.addEventListener('keydown', ev => { if (ev.key === 'Enter') { ev.preventDefault(); next(); } });
        frame(sec, [h('div', { class: 'wizq' }, q), inp, err,
          kind !== 'value' ? h('p', { class: 'note' }, 'Lämna tomt om det är 0.') : null,
          extra === 'paste' ? h('button', { type: 'button', class: 'ghost wide', onclick: () => { pasteMode = true; show(); } }, 'Klistra in innehavslistan från Avanza i stället') : null],
          [h('button', { type: 'button', class: 'ghost', disabled: i === 0 ? true : null, onclick: () => go(-1) }, '‹ Tillbaka'),
           h('button', { type: 'button', class: 'primary', onclick: next }, i === steps.length - 1 ? 'Klar ›' : 'Nästa ›')]);
        setTimeout(() => inp.focus(), 30);
      }
      document.body.append(ov); document.body.classList.add('noscroll'); show();
    }
    setTimeout(() => { if (AV.loaded) autoFill(false); else avLoad().then(() => { autoFill(false); preview(); }); preview(); }, 0);
    return form;
  }


  // ---------------------------------------------------------------- Avanza: innehav, GAV, utdelningar
  const AV = { data: null, etag: null, loaded: false, loading: false, err: null, an: null, filter: 'alla', open: null, paste: false, msg: null };
  async function avLoad(force) {
    if ((AV.loaded && !force) || AV.loading) return;
    AV.loading = true;
    try { const r = await Store.loadAvanza(); AV.data = r.data; AV.etag = r.etag; AV.loaded = true; AV.err = null; avAnalyze(); }
    catch (e) { AV.err = e.message; }
    AV.loading = false; if (/^#(innehav|ar|mer)/.test(location.hash)) render();
  }
  function avSnaps() {
    const d = AV.data; if (!d) return [];
    const s = (d.snaps || []).slice();
    if (d.holdings && !s.some(x => x.date === d.holdings.date)) s.push({ date: d.holdings.date, items: d.holdings.items });
    return s.sort((a, b) => a.date < b.date ? -1 : 1);
  }
  // Bäst/sämst: 'v' = senaste veckan, 'y' = i år (sedan första listan i år, annars äldsta)
  function bastSamst(kind) {
    const s = avSnaps(); if (s.length < 2 || !AV.an) return null;
    const B = s[s.length - 1]; let A;
    if (kind === 'v') { A = s.filter(x => daysBetween(x.date, B.date) >= 5).pop() || s[s.length - 2]; }
    else { A = s.filter(x => x.date < B.date.slice(0, 4) + '-01-08').pop() || s.find(x => x.date.slice(0, 4) === B.date.slice(0, 4)) || s[0]; if (A === B) A = s[s.length - 2]; }
    return { A, B, list: Avanza.perf(AV.data.tx, AV.an, A, B) };
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
      // exportens slutdatum står i filnamnet (transaktioner_ÅÅÅÅ-MM-DD_ÅÅÅÅ-MM-DD.csv), annars dagens datum
      const fm = (f.name || '').match(/(\d{4}-\d{2}-\d{2})\D*\.csv$/i), end = fm ? fm[1] : today();
      await avSave(d => { d.tx = m.rows; if (!d.coverTo || end > d.coverTo) d.coverTo = end; }, m.added ? `${m.added} nya transaktioner (${m.skipped} fanns redan)` : 'Inga nya transaktioner – allt fanns redan');
    } catch (e) { toast(e.message, 'err'); }
  }
  async function avSaveHoldings(text) {
    const h = Avanza.parseHoldings(text);
    if (!AV.loaded) await avLoad();
    await avSave(d => {
      d.holdings = { date: today(), items: h.items, totals: h.totals };
      // ögonblicksbild per dag (används för bäst/sämst och årssammanfattning)
      const slim = h.items.map(i => ({ name: i.name, typ: i.typ, varde: i.varde }));
      d.snaps = (d.snaps || []).filter(x => x.date !== d.holdings.date).concat([{ date: d.holdings.date, items: slim }]).sort((a, b) => a.date < b.date ? -1 : 1).slice(-200);
    }, `Kurser uppdaterade (${h.items.length} innehav)`);
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
      bastSamstKort(),
      h('section', { class: 'card list' }, h('div', { class: 'row pad2' }, h('h2', {}, 'Innehav'),
        h('div', { class: 'seg' }, [['alla', 'Alla'], ['aktie', 'Aktier'], ['fond', 'Fonder'], ['etf', 'ETF']].map(([k, l]) => h('button', { class: AV.filter === k ? 'on' : '', onclick: () => { AV.filter = k; render(); } }, l)))),
        h('div', { class: 'li head h' }, h('span', {}, 'Namn'), h('span', {}, 'Värde kr'), h('span', {}, 'Total avk.')),
        cur.map(row).flat(),
        h('p', { class: 'note pad2' }, 'Värde och "sedan köp" kommer från Avanzas lista. Total avk. = (värde + sålt + utdelningar − köpt) / köpt, för hela tiden du ägt aktien.')),
      h('section', { class: 'card' }, h('h2', {}, 'Utdelningar och resultat per år'),
        yrs.length ? C.bars(yrs.map(y => String(y.year)), [{ name: 'Utdelningar netto', values: yrs.map(y => y.utdNetto), cls: 'c2' }], { labels: true }) : null,
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
  function bastSamstKort() {
    AV.bsKind = AV.bsKind || 'v';
    const r = bastSamst(AV.bsKind);
    const seg = h('div', { class: 'seg' }, [['v', 'Veckan'], ['y', 'I år']].map(([k, l]) => h('button', { class: AV.bsKind === k ? 'on' : '', onclick: () => { AV.bsKind = k; render(); } }, l)));
    const head = h('div', { class: 'row' }, h('h2', {}, 'Bäst och sämst'), seg);
    if (!r) return h('section', { class: 'card' }, head, h('p', { class: 'note' }, 'Visas när du har klistrat in innehavslistan minst två gånger (t.ex. två veckor i rad).'));
    const rad = x => h('div', { class: 'bsrow' }, h('span', {}, x.name), h('b', { class: cls(x.kr) }, krS(x.kr)), h('small', { class: cls(x.pct) }, x.pct === null ? '' : pct(x.pct)));
    const list = r.list, best = list.filter(x => x.kr > 0).slice(0, 3), worst = list.filter(x => x.kr < 0).slice(-3).reverse();
    const tot = list.reduce((a, x) => a + x.kr, 0);
    const all = h('details', { class: 'txl' }, h('summary', {}, `Alla innehav (${list.length})`), h('div', {}, list.map(rad)));
    return h('section', { class: 'card' }, head,
      h('p', { class: 'note' }, `${dstr(r.A.date)} – ${dstr(r.B.date)} · totalt ${krS(tot)}` + (AV.bsKind === 'y' && r.A.date > r.B.date.slice(0, 4) + '-01-08' ? ' · räknat från första listan i år' : '')),
      h('div', { class: 'bsgrid' }, h('div', {}, h('h3', {}, 'Bäst'), best.length ? best.map(rad) : h('p', { class: 'note' }, '–')),
        h('div', {}, h('h3', {}, 'Sämst'), worst.length ? worst.map(rad) : h('p', { class: 'note' }, '–'))), all);
  }
  function txList(tx) {
    const list = tx.slice().reverse();
    return h('details', { class: 'txl' }, h('summary', {}, `Transaktioner (${list.length})`),
      h('table', { class: 'tbl small' }, h('tbody', {}, list.slice(0, 150).map(t => h('tr', {}, h('td', {}, t.d), h('td', {}, t.typ), h('td', {}, t.antal === null ? '' : nf2.format(t.antal)), h('td', {}, t.belopp === null ? '' : nf0.format(t.belopp)))))));
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
      h('section', { class: 'card' }, h('h2', {}, 'Inställningar'), h('a', { class: 'small', href: '#prognos' }, 'Se prognos för målet ›'),
        h('label', { class: 'fld' }, h('span', {}, 'Mål för förmögenhet (kr)'), malIn),
        h('label', { class: 'fld' }, h('span', {}, 'Namn på banken'), bankIn),
        h('div', { class: 'btns' }, h('button', { class: 'ghost', onclick: () => { const v = parseNum(malIn.value); if (!isFinite(v)) return toast('Ogiltigt belopp', 'err');
          const b = bankIn.value.trim(); commit(d => { d.settings = d.settings || {}; d.settings.mal = v; if (b) d.settings.bank = b; else delete d.settings.bank; }, 'Inställningarna är sparade'); } }, 'Spara'))),
      h('section', { class: 'card' }, h('h2', {}, 'Årssammanfattning'), h('p', { class: 'note' }, 'En sammanfattning per år som du kan spara som PDF.'), h('div', { class: 'btns' }, h('a', { class: 'ghost btn', href: '#ar' }, 'Öppna årssammanfattning'))),
      kontoKort(),
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

  function kontoKort() {
    if (!AV.loaded) { avLoad().then(() => { if (location.hash.startsWith('#mer')) render(); }); return null; }
    if (!AV.data || !(AV.data.tx || []).length) return null;
    const ex = new Set(((S.data.settings || {}).avanzaEx) || []);
    const accs = Avanza.accounts(AV.data.tx), boxes = {};
    return h('section', { class: 'card' }, h('h2', {}, 'Avanza-konton'),
      h('p', { class: 'note' }, 'Välj vilka konton som räknas med i förmögenheten. Avstängda konton räknas bort när veckans värden och insättningar fylls i automatiskt (de syns fortfarande under Innehav).'),
      accs.map(a => { const cb = h('input', { type: 'checkbox', checked: ex.has(a.konto) ? null : true }); boxes[a.konto] = cb;
        const val = a.kind === 'sparkonto' ? `sparkonto · saldo ${kr(a.cash)}` : `värdepapper${AV.data.holdings && AV.an ? ' · värde ' + kr(Object.values(Avanza.excludedValues(AV.an, AV.data.holdings, [a.konto])).reduce((x, y) => x + y, 0)) : ''}`;
        return h('label', { class: 'acc' }, cb, h('span', {}, h('b', {}, a.konto), h('small', {}, `${val} · senast ${dstr(a.last)}`))); }),
      h('div', { class: 'btns' }, h('button', { class: 'ghost', onclick: () => {
        const list = Object.entries(boxes).filter(([k, cb]) => !cb.checked).map(([k]) => k);
        commit(d => { d.settings = d.settings || {}; d.settings.avanzaEx = list; }, 'Kontoinställningen är sparad'); } }, 'Spara')));
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
    const navR = ['avkastning', 'sparande', 'konto', 'prognos', 'isk', 'ar'].includes(route) ? 'oversikt' : route;
    document.querySelectorAll('#nav a').forEach(a => a.classList.toggle('on', a.getAttribute('href') === '#' + navR));
    $('#status').textContent = S.busy ? 'Sparar…' : S.offline ? 'Offline – visar sparad kopia' : Store.isDemo() ? 'Demoläge' : '';
    let v;
    if (S.needLogin) { v = viewLogin(S.loginMsg); nav.style.display = 'none'; }
    else if (!S.data) { v = S.loading ? h('div', { class: 'view' }, h('p', { class: 'note pad' }, 'Hämtar data från OneDrive…')) : viewEmpty(); }
    else if (route === 'ny') v = viewForm(new URLSearchParams(qs || '').get('d'));
    else if (!S.data.rows.length) v = viewEmpty();
    else if (route === 'historik') v = viewHistory();
    else if (route === 'innehav') v = viewInnehav();
    else if (route === 'avkastning') v = viewAvkastning();
    else if (route === 'sparande') v = viewSparande();
    else if (route === 'prognos') v = viewPrognos();
    else if (route === 'isk') v = viewIsk();
    else if (route === 'ar') v = viewAr(new URLSearchParams(qs || '').get('y'));
    else if (route === 'konto') v = viewKonto(new URLSearchParams(qs || '').get('k'));
    else if (route === 'mer') v = viewMore();
    else v = viewOverview();
    // behåll formulär vid omritning medan man skriver
    if (route === 'ny' && main.firstChild && main.firstChild.tagName === 'FORM' && main.dataset.route === hash) {
      const btn = main.querySelector('button[type=submit]'); if (btn) { btn.disabled = S.busy; btn.textContent = S.busy ? 'Sparar…' : 'Spara'; } return;
    }
    main.dataset.route = hash; main.replaceChildren(v);
    const on = main.querySelector('.chips button.on'); if (on) on.scrollIntoView({ block: 'nearest', inline: 'center' });
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
  // Om webbläsaren har en äldre index.html än app.js: ladda om en gång så att allt kommer från samma version
  const mv = document.querySelector('meta[name=app-version]');
  if ((!mv || mv.content !== window.EKONOMI_VERSION) && !sessionStorage.getItem('ek.reloaded')) { sessionStorage.setItem('ek.reloaded', '1'); location.reload(); return; }
  sessionStorage.removeItem('ek.reloaded');
  boot();
  if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js');
})();
