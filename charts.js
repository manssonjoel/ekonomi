/* Grafer (SVG, inga bibliotek): linjer, tidsstaplar, staplar med linje, donut, år-mot-år */
(function (root) {
  'use strict';
  const nf0 = new Intl.NumberFormat('sv-SE', { maximumFractionDigits: 0 });
  const nf1 = new Intl.NumberFormat('sv-SE', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  const nf2 = new Intl.NumberFormat('sv-SE', { maximumFractionDigits: 2 });
  const MON = ['jan', 'feb', 'mar', 'apr', 'maj', 'jun', 'jul', 'aug', 'sep', 'okt', 'nov', 'dec'];
  const dstr = d => `${+d.slice(8, 10)} ${MON[+d.slice(5, 7) - 1]} ${d.slice(0, 4)}`;
  const mkr = x => Math.abs(x) >= 1e6 ? nf2.format(x / 1e6) + ' Mkr' : Math.abs(x) >= 1e3 ? nf0.format(x / 1e3) + ' tkr' : nf0.format(x);
  const fmt = (v, kind) => v === null || v === undefined || !isFinite(v) ? '–' : kind === 'pct' ? (v > 0 ? '+' : '') + nf1.format(v * 100) + ' %' : nf0.format(Math.round(v)) + ' kr';
  const axis = (v, kind) => kind === 'pct' ? nf0.format(v * 100) + ' %' : mkr(v);
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;');
  const W = () => { const m = document.getElementById('main'); return Math.max(300, Math.min(736, (m ? m.clientWidth : 640) - 26)); };
  function niceStep(r) { if (!(r > 0)) return 1; const p = Math.pow(10, Math.floor(Math.log10(r))); const m = r / p; return (m <= 1 ? 1 : m <= 2 ? 2 : m <= 2.5 ? 2.5 : m <= 5 ? 5 : 10) * p; }
  function scale(vals, incZero = true) {
    const v = vals.filter(x => x !== null && isFinite(x)); let lo = Math.min(...v, incZero ? 0 : Infinity), hi = Math.max(...v, incZero ? 0 : -Infinity);
    if (!v.length) { lo = 0; hi = 1; } if (hi === lo) hi = lo + (Math.abs(lo) || 1);
    const st = niceStep((hi - lo) / 4); return { lo: Math.floor(lo / st) * st, hi: Math.ceil(hi / st) * st, st };
  }
  function grid(sc, y, L, Wd, kind, right) {
    let g = '';
    for (let v = sc.lo; v <= sc.hi + sc.st / 2; v += sc.st) g += `<line x1="${L}" x2="${Wd}" y1="${y(v)}" y2="${y(v)}" class="grid${Math.abs(v) < 1e-9 ? ' zero' : ''}"/>` +
      `<text x="${right ? Wd + 4 : L - 6}" y="${y(v) + 4}" class="ax" text-anchor="${right ? 'start' : 'end'}">${axis(v, kind)}</text>`;
    return g;
  }
  // Behållare med tooltip: tipAt(px) -> {x, html} eller null
  function box(svg, Wd, H, tipAt) {
    const d = document.createElement('div'); d.className = 'chart';
    d.innerHTML = `<svg viewBox="0 0 ${Wd} ${H}" role="img">${svg}<line class="cursor" x1="0" x2="0" y1="8" y2="${H - 22}" style="display:none"/></svg><div class="tip"></div>`;
    if (!tipAt) return d;
    const tip = d.querySelector('.tip'), cur = d.querySelector('.cursor');
    const move = ev => {
      const r = d.getBoundingClientRect(), cx = (ev.touches ? ev.touches[0].clientX : ev.clientX) - r.left;
      const t = tipAt(cx / r.width * Wd); if (!t) return;
      cur.style.display = ''; cur.setAttribute('x1', t.x); cur.setAttribute('x2', t.x);
      tip.innerHTML = t.html; tip.style.display = 'block';
      tip.style.left = Math.min(r.width - 180, Math.max(0, t.x / Wd * r.width - 90)) + 'px';
    };
    d.addEventListener('pointermove', move); d.addEventListener('pointerdown', move); d.addEventListener('touchmove', move, { passive: true });
    d.addEventListener('pointerleave', () => { tip.style.display = 'none'; cur.style.display = 'none'; });
    return d;
  }
  function yearTicks(dates, x, H, B) {
    let g = '', last = null; const n = dates.length, labs = [];
    const span = n > 1 ? (Date.parse(dates[n - 1]) - Date.parse(dates[0])) / 864e5 : 0;
    dates.forEach((d, i) => {
      const key = span < 400 ? d.slice(0, 7) : d.slice(0, 4);
      if (key !== last) {
        if (i > 0) g += `<line x1="${x(i)}" x2="${x(i)}" y1="8" y2="${H - B}" class="grid v"/>`;
        if (span >= 400 || +d.slice(5, 7) % (span < 200 ? 1 : 2) === 1) labs.push([x(i), span < 400 ? MON[+d.slice(5, 7) - 1] : d.slice(0, 4)]);
      }
      last = key;
    });
    labs.forEach((l, j) => { if (j + 1 < labs.length && labs[j + 1][0] - l[0] < 34) return; g += `<text x="${l[0] + 3}" y="${H - 7}" class="ax">${l[1]}</text>`; });
    return g;
  }

  // Linjer över tid. series: [{name, values, cls, fill, dash}]
  function line(dates, series, o = {}) {
    const Wd = W(), H = o.h || 220, L = 56, R = 10, T = 10, B = 24, kind = o.kind || 'kr';
    const sc = scale(series.flatMap(s => s.values), o.zero !== false), n = dates.length;
    const x = i => L + (Wd - L - R) * (n > 1 ? i / (n - 1) : .5), y = v => T + (H - T - B) * (1 - (v - sc.lo) / (sc.hi - sc.lo));
    let g = grid(sc, y, L, Wd - R, kind) + yearTicks(dates, x, H, B);
    series.forEach(s => {
      const pts = s.values.map((v, i) => v !== null && isFinite(v) ? `${x(i).toFixed(1)},${y(v).toFixed(1)}` : null).filter(Boolean);
      if (!pts.length) return;
      if (s.fill) g += `<polygon points="${x(0)},${y(Math.max(sc.lo, 0))} ${pts.join(' ')} ${x(n - 1)},${y(Math.max(sc.lo, 0))}" class="fill ${s.cls}"/>`;
      g += `<polyline points="${pts.join(' ')}" class="line ${s.cls}${s.dash ? ' dash' : ''}"/>`;
    });
    return box(g, Wd, H, px => { const i = Math.max(0, Math.min(n - 1, Math.round((px - L) / (Wd - L - R) * (n - 1))));
      return { x: x(i), html: `<b>${dstr(dates[i])}</b>` + series.map(s => `<div><i class="dot ${s.cls}"></i>${esc(s.name)}: ${fmt(s.values[i], kind)}</div>`).join('') }; });
  }
  // Staplar över tid (en per vecka), positiva/negativa färger
  function timeBars(dates, values, o = {}) {
    const Wd = W(), H = o.h || 200, L = 56, R = 10, T = 10, B = 24, kind = o.kind || 'kr', n = dates.length;
    const sc = scale(values), bw = (Wd - L - R) / Math.max(1, n);
    const x = i => L + i * bw + bw / 2, y = v => T + (H - T - B) * (1 - (v - sc.lo) / (sc.hi - sc.lo));
    let g = grid(sc, y, L, Wd - R, kind) + yearTicks(dates, i => L + i * bw, H, B);
    values.forEach((v, i) => { if (v === null || !isFinite(v)) return; const w = Math.max(.6, bw * (n > 120 ? .9 : .7));
      g += `<rect x="${(x(i) - w / 2).toFixed(1)}" y="${Math.min(y(v), y(0)).toFixed(1)}" width="${w.toFixed(1)}" height="${Math.max(.5, Math.abs(y(v) - y(0))).toFixed(1)}" class="bar ${o.cls || (v >= 0 ? 'pos' : 'neg')}"/>`; });
    return box(g, Wd, H, px => { const i = Math.max(0, Math.min(n - 1, Math.floor((px - L) / bw)));
      return { x: x(i), html: `<b>${dstr(dates[i])}</b><div>${esc(o.name || '')}${o.name ? ': ' : ''}${fmt(values[i], kind)}</div>` + (o.extra ? o.extra(i) : '') }; });
  }
  // Staplar per kategori (t.ex. år), flera serier sida vid sida, valfri linje på höger axel
  function bars(labels, series, o = {}) {
    const Wd = W(), H = o.h || 220, kind = o.kind || 'kr', ln = o.line, L = 56, R = ln ? 46 : 10, T = 18, B = 24, n = labels.length;
    const sc = scale(series.flatMap(s => s.values)), bw = (Wd - L - R) / Math.max(1, n), k = series.length, inner = bw * .74 / k;
    const y = v => T + (H - T - B) * (1 - (v - sc.lo) / (sc.hi - sc.lo));
    let g = grid(sc, y, L, Wd - R, kind);
    labels.forEach((lab, i) => {
      series.forEach((s, j) => { const v = s.values[i]; if (v === null || !isFinite(v)) return; const x0 = L + i * bw + bw * .13 + j * inner;
        g += `<rect x="${x0.toFixed(1)}" y="${Math.min(y(v), y(0)).toFixed(1)}" width="${(inner * .92).toFixed(1)}" height="${Math.max(.5, Math.abs(y(v) - y(0))).toFixed(1)}" class="bar ${s.cls || (v >= 0 ? 'pos' : 'neg')}" rx="2"/>`;
        if (o.labels && k === 1) g += `<text x="${x0 + inner * .46}" y="${v >= 0 ? y(v) - 4 : y(0) + 12}" class="ax val${v < 0 ? ' inbar' : ''}" text-anchor="middle">${kind === 'pct' ? nf1.format(v * 100) : mkr(v)}</text>`; });
      g += `<text x="${L + i * bw + bw / 2}" y="${H - 7}" class="ax" text-anchor="middle">${lab}</text>`;
    });
    let ys = null;
    if (ln) {
      const s2 = scale(ln.values); ys = v => T + (H - T - B) * (1 - (v - s2.lo) / (s2.hi - s2.lo));
      for (let v = s2.lo; v <= s2.hi + s2.st / 2; v += s2.st) g += `<text x="${Wd - R + 6}" y="${ys(v) + 4}" class="ax">${axis(v, ln.kind || 'pct')}</text>`;
      const pts = ln.values.map((v, i) => v !== null && isFinite(v) ? `${L + i * bw + bw / 2},${ys(v)}` : null).filter(Boolean);
      g += `<polyline points="${pts.join(' ')}" class="line ${ln.cls}"/>` + ln.values.map((v, i) => v !== null && isFinite(v) ? `<circle cx="${L + i * bw + bw / 2}" cy="${ys(v)}" r="3" class="pt ${ln.cls}"/>` : '').join('');
    }
    return box(g, Wd, H, px => { const i = Math.max(0, Math.min(n - 1, Math.floor((px - L) / bw)));
      return { x: L + i * bw + bw / 2, html: `<b>${labels[i]}</b>` + series.map(s => `<div><i class="dot ${s.cls || 'pos'}"></i>${esc(s.name || '')}: ${fmt(s.values[i], kind)}</div>`).join('') +
        (ln ? `<div><i class="dot ${ln.cls}"></i>${esc(ln.name)}: ${fmt(ln.values[i], ln.kind || 'pct').replace('+', '')}</div>` : '') }; });
  }
  // Donut: items [{name, value, cls}]
  function donut(items, o = {}) {
    const tot = items.reduce((s, i) => s + Math.max(0, i.value), 0) || 1, R = 70, r = 46, c = 85;
    let a = -Math.PI / 2, g = '';
    items.forEach(it => { const f = Math.max(0, it.value) / tot; if (f <= 0) return; const a2 = a + f * Math.PI * 2 - 1e-6, large = f > .5 ? 1 : 0;
      const p = (ang, rad) => `${(c + rad * Math.cos(ang)).toFixed(2)},${(c + rad * Math.sin(ang)).toFixed(2)}`;
      g += `<path d="M${p(a, R)} A${R},${R} 0 ${large} 1 ${p(a2, R)} L${p(a2, r)} A${r},${r} 0 ${large} 0 ${p(a, r)} Z" class="slice ${it.cls}"/>`; a = a2; });
    if (o.center) g += `<text x="${c}" y="${c - 2}" text-anchor="middle" class="dc1">${esc(o.center[0])}</text><text x="${c}" y="${c + 15}" text-anchor="middle" class="dc2">${esc(o.center[1] || '')}</text>`;
    const d = document.createElement('div'); d.className = 'donut';
    d.innerHTML = `<svg viewBox="0 0 170 170">${g}</svg><div class="dlegend">` + items.map(it => `<div><i class="dot ${it.cls}"></i><span>${esc(it.name)}</span><b>${nf1.format(Math.max(0, it.value) / tot * 100)} %</b><small>${mkr(it.value)}</small></div>`).join('') + '</div>';
    return d;
  }
  // År mot år: kumulativt värde per dag på året. series: [{name, pts:[[dag, värde]], cls}]
  function yearLines(series, o = {}) {
    const Wd = W(), H = o.h || 220, L = 56, R = 10, T = 10, B = 24, kind = o.kind || 'kr';
    const sc = scale(series.flatMap(s => s.pts.map(p => p[1])));
    const x = d => L + (Wd - L - R) * d / 365, y = v => T + (H - T - B) * (1 - (v - sc.lo) / (sc.hi - sc.lo));
    let g = grid(sc, y, L, Wd - R, kind);
    [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334].forEach((d, m) => { if (m) g += `<line x1="${x(d)}" x2="${x(d)}" y1="${T}" y2="${H - B}" class="grid v"/>`; if (m % 2 === 0 || Wd > 500) g += `<text x="${x(d) + 2}" y="${H - 7}" class="ax">${MON[m]}</text>`; });
    series.forEach(s => { if (!s.pts.length) return; const pts = [[0, 0]].concat(s.pts);
      g += `<polyline points="${pts.map(p => `${x(p[0]).toFixed(1)},${y(p[1]).toFixed(1)}`).join(' ')}" class="line step ${s.cls}"/>`;
      if (o.labels !== false) { const lp = pts[pts.length - 1]; g += `<text x="${Math.min(x(lp[0]) + 3, Wd - 32)}" y="${y(lp[1]) - 3}" class="ax ylab ${s.cls}">${esc(s.name)}</text>`; } });
    return box(g, Wd, H, px => { const day = Math.max(0, Math.min(365, (px - L) / (Wd - L - R) * 365)); const dt = new Date(Date.UTC(2001, 0, 1 + Math.round(day)));
      return { x: x(day), html: `<b>${dt.getUTCDate()} ${MON[dt.getUTCMonth()]}</b>` + series.map(s => { let v = 0; for (const p of s.pts) if (p[0] <= day) v = p[1]; return `<div><i class="dot ${s.cls}"></i>${esc(s.name)}: ${fmt(v, kind)}</div>`; }).join('') }; });
  }
  const dayOfYear = d => (Date.parse(d) - Date.parse(d.slice(0, 4) + '-01-01')) / 864e5;
  root.Charts = { line, timeBars, bars, donut, yearLines, dayOfYear, mkr };
})(this);
