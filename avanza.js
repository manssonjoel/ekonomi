/* Avanza: läser transaktionsexporten (CSV) och innehavslistan (inklistrad text),
   räknar innehav, GAV, utdelningar och realiserade resultat. */
(function (root) {
  'use strict';
  const num = s => {
    if (s === null || s === undefined) return null;
    s = String(s).trim(); if (!s || s === '-') return null;
    s = s.replace(/[\s  ]/g, '').replace(/kr$/i, '').replace(/%$/, '').replace(/−/g, '-').replace(',', '.');
    const v = parseFloat(s); return isFinite(v) ? v : null;
  };
  const norm = s => String(s || '').toLowerCase().replace(/[\s ]+/g, ' ').trim();

  // ---------------------------------------------------------------- CSV
  function splitCSV(line, sep) {
    const out = []; let cur = '', q = false;
    for (let i = 0; i < line.length; i++) {
      const c = line[i];
      if (q) { if (c === '"') { if (line[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += c; }
      else if (c === '"') q = true; else if (c === sep) { out.push(cur); cur = ''; } else cur += c;
    }
    out.push(cur); return out;
  }
  const COLS = { 'datum': 'd', 'konto': 'konto', 'typ av transaktion': 'typ', 'värdepapper/beskrivning': 'namn', 'antal': 'antal', 'kurs': 'kurs',
    'belopp': 'belopp', 'transaktionsvaluta': 'valuta', 'courtage': 'courtage', 'valutakurs': 'vkurs', 'instrumentvaluta': 'ivaluta', 'isin': 'isin', 'resultat': 'resultat' };
  function parseCSV(text) {
    text = text.replace(/^﻿/, '');
    const lines = text.split(/\r?\n/).filter(l => l.trim());
    if (!lines.length) throw new Error('Filen är tom.');
    const sep = lines[0].includes(';') ? ';' : lines[0].includes('\t') ? '\t' : ',';
    const head = splitCSV(lines[0], sep).map(h => COLS[norm(h)] || null);
    if (!head.includes('d') || !head.includes('typ') || !head.includes('belopp')) throw new Error('Känner inte igen filen – välj transaktionsexporten från Avanza (CSV).');
    const rows = [];
    lines.slice(1).forEach((l, i) => {
      const f = splitCSV(l, sep), r = { seq: i };
      head.forEach((k, j) => { if (k) r[k] = (f[j] || '').trim(); });
      if (!/^\d{4}-\d{2}-\d{2}$/.test(r.d)) return;
      for (const k of ['antal', 'kurs', 'belopp', 'courtage', 'vkurs', 'resultat']) r[k] = num(r[k]);
      rows.push(r);
    });
    return rows;
  }
  const key = r => [r.d, r.konto, r.typ, r.namn, r.antal, r.kurs, r.belopp, r.isin].join('|');
  // Slår ihop exporter: samma rad räknas bara så många gånger som den förekommer i en enskild fil.
  function merge(existing, incoming) {
    const cnt = (list) => { const m = new Map(); list.forEach(r => m.set(key(r), (m.get(key(r)) || 0) + 1)); return m; };
    const have = cnt(existing), inc = cnt(incoming), used = new Map(); let added = 0;
    const base = existing.length ? Math.max(...existing.map(r => r.seq || 0)) + 1 : 0;
    const out = existing.slice();
    incoming.forEach(r => {
      const k = key(r), u = (used.get(k) || 0) + 1; used.set(k, u);
      if (u > (have.get(k) || 0)) { out.push(Object.assign({}, r, { seq: base + r.seq })); added++; }
    });
    return { rows: out, added, total: out.length, skipped: incoming.length - added };
  }

  // ---------------------------------------------------------------- innehavslista (inklistrad från Avanza)
  // forced = 'aktie' | 'fond' | 'etf' när man klistrar in en grupp utan rubrik.
  // Klarar datorns tabellformat (namnrad + tabbrad) och mobilappens format (namn, sedan rubrik/värde på varsin rad).
  const LABEL = /^(idag|igår|i går|i dag|\d{1,2} [a-zåäö]{3}\.?( \d{4})?|\d+ ?(d|dag|dagar|v|vecka|veckor|m|mån|månader|år)|i år|ytd|sedan köp( %| kr)?|värde( kr)?|utveckling( %)?|senast|kurs|antal|gav|anskaffningsvärde|marknadsvärde|1 dag( %)?|idag %|avkastning|%|kr)$/i;
  function parseHoldings(text, forced) {
    const lines = text.split(/\r?\n/).map(l => l.replace(/[  ]/g, ' ')).filter(l => l.trim());
    const items = [], sections = new Set(), shown = []; let typ = forced || null, lastName = null, blk = null, totNext = false;
    const SKIP = /^(inköpsinfo|utdelningar|historik|handelsinfo|land|namn|nav|verktyg|konto|innehav|min ekonomi|översikt|ordrar|transaktioner|analys|kalender|köp|sälj|fyll på ditt innehav|visa (fler|alla)|sortera|just nu|marknaden|upptäck|meny)$/i;
    const push = (name, o) => { const t = typ || 'aktie'; sections.add(t); items.push(Object.assign({ name, typ: t, konto: null }, o)); };
    // mobilblock: { name, pairs: [[label, value, raw]] }
    const flushBlk = () => {
      if (blk && blk.pairs.length) {
        const P = blk.pairs, get = re => { const x = P.find(p => p[0] && re.test(p[0])); return x ? x : null; };
        const v = get(/^(värde|marknadsvärde)/i), krs = P.filter(p => /kr\s*$/i.test(p[2]));
        const varde = v ? v[1] : krs.length ? krs[krs.length - 1][1] : P[P.length - 1][1];
        const sk = get(/^sedan köp/i), skKr = P.find(p => /^sedan köp/i.test(p[0] || '') && /kr\s*$/i.test(p[2]));
        const pctv = sk && /%\s*$/.test(sk[2]) ? sk[1] / 100 : null;
        const sedanKop = skKr ? skKr[1] : pctv !== null && pctv > -1 ? varde - varde / (1 + pctv) : 0;
        push(blk.name, { senast: null, sedanKop, sedanKopPct: pctv !== null ? pctv : (varde - sedanKop ? sedanKop / (varde - sedanKop) : 0), varde });
      }
      blk = null;
    };
    for (const raw of lines) {
      const l = raw.trim();
      const hdr = /^aktier$/i.test(l) ? 'aktie' : /^fonder$/i.test(l) ? 'fond' : /^(etf:er|etf)$/i.test(l) ? 'etf' : /^(certifikat|obligationer|övriga|optioner|warranter)$/i.test(l) ? 'övrigt' : null;
      if (hdr) { flushBlk(); lastName = null; typ = hdr; continue; }
      if (/^totalt( värde)?\b/i.test(l)) { flushBlk(); lastName = null; const n = l.split(/\t/).map(num).filter(x => x !== null); if (n.length) shown.push(n[n.length - 1]); else totNext = true; continue; }
      if (SKIP.test(l)) continue;
      const parts = raw.split('\t').map(x => x.trim()).filter(x => x !== '');
      if (parts.length >= 4) {   // datorns tabbrad
        const nums = parts.map(num);
        if (nums.slice(-4).every(x => x !== null) && lastName) {
          const n = nums.filter(x => x !== null), [senast, sedanKop, sedanKopPct, varde] = n.slice(-4);
          push(lastName, { konto: parts.length > 5 && /^\d{5,}$/.test(parts[0]) ? parts[0] : null, senast, sedanKop, sedanKopPct: sedanKopPct / 100, varde });
          lastName = null; blk = null;
        }
        continue;
      }
      if (LABEL.test(l)) { if (blk) blk.label = l; continue; }
      const n = parts.length === 1 ? num(l) : null;
      if (n !== null && /\d/.test(l) && !/[a-zåäö]{3,}/i.test(l.replace(/kr|sek|usd|eur/ig, ''))) {
        if (totNext) { shown.push(n); totNext = false; continue; }
        if (blk) { blk.pairs.push([blk.label || null, n, l]); blk.label = null; }
        continue;
      }
      totNext = false;
      flushBlk(); lastName = l; blk = { name: l, pairs: [], label: null };
    }
    flushBlk();
    if (!items.length) throw new Error('Hittade inga innehav. Kopiera listan från Avanza (Min ekonomi › Innehav) – hela listan eller en grupp i taget.');
    const totals = { aktie: 0, fond: 0, etf: 0, övrigt: 0 };
    items.forEach(i => { totals[i.typ] = (totals[i.typ] || 0) + i.varde; });
    totals.alla = items.reduce((s, i) => s + i.varde, 0);
    // kontroll mot Avanzas egen "Totalt"-rad (bara när en enda grupp klistrats in)
    let varning = null;
    if (shown.length === 1 && sections.size === 1) { const sum = totals.alla; if (Math.abs(sum - shown[0]) > Math.max(2, shown[0] * 0.001)) varning = { summa: sum, avanza: shown[0] }; }
    return { items, totals, sections: [...sections], varning };
  }

  // ---------------------------------------------------------------- analys
  const DIV = new Set(['Utdelning', 'Utländsk källskatt']);
  // Ändrar antal men inte anskaffningsvärdet (splittar, byten, emissioner m.m.)
  const CONV = new Set(['Split värdepapper', 'Split nytt värdepapper', 'Byte', 'Byte värdepapper', 'Teckning rätter', 'Teckning inlägg',
    'Tilldelning nyemission', 'Utdelning värdepapper', 'Avskiljning', 'Värdepappersinsättning']);
  const LINK = new Set(['Split värdepapper', 'Split nytt värdepapper', 'Byte', 'Byte värdepapper', 'Teckning rätter', 'Teckning inlägg']);
  const TAXFEE = new Set(['Avkastningsskatt', 'Uttag av riskkostnad', 'ADR-avgift', 'Utlåningsränta', 'Preliminärskatt kapitalränta', 'Justering ränta']);
  const INCOME = new Set(['Inlåningsränta', 'Utbetalning aktielån']);

  function analyze(tx, holdings) {
    const rows = tx.slice().sort((a, b) => a.d < b.d ? -1 : a.d > b.d ? 1 : (b.seq || 0) - (a.seq || 0));
    // koppla ihop ISIN som byts mot varandra (split till ny ISIN, byte, teckning)
    const par = new Map(); const find = x => { while (par.get(x) !== x) { par.set(x, par.get(par.get(x))); x = par.get(x); } return x; };
    const idOf = r => r.isin ? 'I:' + r.isin : (r.namn && (r.antal !== null || DIV.has(r.typ)) ? 'N:' + norm(r.namn) : null);
    rows.forEach(r => { const id = idOf(r); if (id && !par.has(id)) par.set(id, id); });
    const ev = new Map();
    rows.filter(r => LINK.has(r.typ) && idOf(r)).forEach(r => { const k = r.d + '|' + r.konto; (ev.get(k) || ev.set(k, []).get(k)).push(idOf(r)); });
    ev.forEach(ids => ids.forEach(id => { const a = find(ids[0]), b = find(id); if (a !== b) par.set(b, a); }));

    const G = new Map(), year = d => +d.slice(0, 4);
    const perYear = {}; const Y = y => perYear[y] || (perYear[y] = { utd: 0, kallskatt: 0, realiserat: 0, skatt: 0, ranta: 0 });
    for (const r of rows) {
      const y = year(r.d), cash = r.belopp || 0;
      if (r.typ === 'Återbetalning utländsk källskatt') { Y(y).kallskatt += cash; continue; }
      if (TAXFEE.has(r.typ)) { Y(y).skatt += cash; continue; }
      if (INCOME.has(r.typ)) { Y(y).ranta += cash; continue; }
      const id = idOf(r); if (!id) continue;
      const gid = find(id);
      const g = G.get(gid) || G.set(gid, { id: gid, acc: new Map(), ry: {}, dy: {}, bought: 0, proceeds: 0, removed: 0, div: 0, names: [], tx: [], first: r.d, last: r.d, isins: new Set() }).get(gid);
      g.tx.push(r); g.last = r.d; if (r.isin) g.isins.add(r.isin);
      if (r.namn && (r.typ === 'Köp' || r.typ === 'Sälj')) g.names.push(r.namn);
      if (DIV.has(r.typ)) { g.dy[y] = (g.dy[y] || 0) + cash; g.div += cash; if (r.typ === 'Utdelning') Y(y).utd += cash; else Y(y).kallskatt += cash; continue; }
      // genomsnittsmetoden räknas per konto, som hos Avanza
      const q = r.antal || 0, A = g.acc.get(r.konto) || g.acc.set(r.konto, { qty: 0, cost: 0 }).get(r.konto);
      if (q > 0) { if (cash < 0 && !CONV.has(r.typ)) { A.cost += -cash; g.bought += -cash; } A.qty += q; }
      else if (q < 0) {
        if (CONV.has(r.typ)) A.qty += q;
        else {
          const frac = A.qty > 1e-9 ? Math.min(1, -q / A.qty) : 1, rem = A.cost * frac;
          A.cost -= rem; g.removed += rem; A.qty += q; Y(y).realiserat -= rem; g.ry[y] = (g.ry[y] || 0) - rem;
          if (cash > 0) { g.proceeds += cash; Y(y).realiserat += cash; g.ry[y] += cash; }
          if (Math.abs(A.qty) < 1e-6) { A.qty = 0; A.cost = 0; }
        }
      } else {
        if (cash < 0) { A.cost += -cash; g.bought += -cash; }
        else if (cash > 0) { g.proceeds += cash; Y(y).realiserat += cash; g.ry[y] = (g.ry[y] || 0) + cash; }
      }
    }
    const groups = [...G.values()].map(g => {
      g.name = g.names.length ? g.names[g.names.length - 1] : g.tx[g.tx.length - 1].namn;
      g.qty = 0; g.cost = 0; g.acc.forEach(a => { g.qty += a.qty; g.cost += a.cost; });
      if (Math.abs(g.qty) < 1e-6) g.qty = 0;
      g.realized = g.proceeds - g.removed;
      g.allNames = new Set(g.tx.map(t => norm(t.namn)));
      return g;
    });

    // matcha mot innehavslistan
    const current = [], used = new Set();
    if (holdings && holdings.items) {
      const byName = new Map();
      holdings.items.forEach(h => { const k = norm(h.name); const a = byName.get(k); if (a) { a.varde += h.varde; a.sedanKop += h.sedanKop; } else byName.set(k, Object.assign({}, h)); });
      for (const [k, h] of byName) {
        const cand = groups.filter(x => !used.has(x) && (norm(x.name) === k || x.allNames.has(k)));
        const g = cand.sort((a, b) => (b.qty - a.qty) || ((norm(b.name) === k) - (norm(a.name) === k)))[0] || null;
        if (g) used.add(g);
        const ansk = h.varde - h.sedanKop;
        current.push({ name: h.name, typ: h.typ, varde: h.varde, senast: h.senast, sedanKop: h.sedanKop, sedanKopPct: ansk ? h.sedanKop / ansk : null, ansk,
          qty: g ? g.qty : null, gav: g && g.qty > 1e-9 ? ansk / g.qty : null, gavBeraknad: g && g.qty > 1e-9 ? g.cost / g.qty : null,
          utd: g ? g.div : 0, kopt: g ? g.bought : null, salt: g ? g.proceeds : 0,
          total: g && g.bought > 0 ? (h.varde + g.proceeds + g.div - g.bought) / g.bought : null, group: g });
      }
      const tot = current.reduce((s, c) => s + c.varde, 0); current.forEach(c => c.andel = tot ? c.varde / tot : null);
      current.sort((a, b) => b.varde - a.varde);
    } else {
      groups.filter(g => g.qty > 1e-6).forEach(g => { used.add(g); current.push({ name: g.name, typ: null, varde: null, qty: g.qty, gavBeraknad: g.cost / g.qty, gav: g.cost / g.qty, ansk: g.cost, utd: g.div, kopt: g.bought, salt: g.proceeds, group: g }); });
      current.sort((a, b) => (b.ansk || 0) - (a.ansk || 0));
    }
    const closed = groups.filter(g => !used.has(g) && g.qty <= 1e-6 && (g.bought > 0 || g.proceeds > 0))
      .map(g => ({ name: g.name, kopt: g.bought, salt: g.proceeds, resultat: g.proceeds - g.bought, utd: g.div, total: g.bought > 0 ? (g.proceeds + g.div - g.bought) / g.bought : null, last: g.last, group: g }))
      .sort((a, b) => a.last < b.last ? 1 : -1);
    const years = Object.keys(perYear).sort().map(y => Object.assign({ year: +y }, perYear[y], { utdNetto: perYear[y].utd + perYear[y].kallskatt }));
    return { current, closed, years, groups, count: tx.length, first: rows.length ? rows[0].d : null, lastDate: rows.length ? rows[rows.length - 1].d : null };
  }


  // ---------------------------------------------------------------- veckans Avanza-flöden
  const SECCASH = new Set(['Köp', 'Sälj', 'Rättelse sälj', 'Teckning likvid', 'Bokning fraktionslikvider', 'Inlösen']);
  const RETURN = new Set(['Utdelning', 'Utländsk källskatt', 'Återbetalning utländsk källskatt', 'Avkastningsskatt', 'Uttag av riskkostnad', 'ADR-avgift',
    'Utlåningsränta', 'Preliminärskatt kapitalränta', 'Justering ränta', 'Inlåningsränta', 'Utbetalning aktielån']);
  const SECTYPES = new Set([...SECCASH, ...DIV, ...CONV, 'Utbokning fraktioner', 'Värdepappersuttag', 'Split värdepapper']);
  function typeOf(g, holdings) {
    if (holdings && holdings.items) for (const i of holdings.items) if (g.allNames.has(norm(i.name))) return i.typ;
    const s = [...g.allNames].join(' ');
    if (/etf|ucits|ishares|xact|xtrackers|amundi|vanguard|spdr/.test(s)) return 'etf';
    if (/fond|robur|avanza (zero|global|usa|emerging|auto|sweden|tech|fastighet|europa|asien)|handelsbanken|spiltan aktie|lannebo|didner|skagen|plus |indexnära/.test(s)) return 'fond';
    return 'aktie';
  }
  // Konton i exporten: värdepapperskonton och rena sparkonton (t.ex. "Buffert")
  function accounts(tx) {
    const m = new Map();
    tx.forEach(t => { const a = m.get(t.konto) || m.set(t.konto, { konto: t.konto, sec: false, cash: 0, last: t.d, n: 0 }).get(t.konto);
      a.n++; a.cash += t.belopp || 0; if (t.d > a.last) a.last = t.d; if (SECTYPES.has(t.typ) && t.isin) a.sec = true; });
    return [...m.values()].map(a => Object.assign(a, { cash: Math.round(a.cash * 100) / 100 + 0, kind: a.sec ? 'värdepapper' : 'sparkonto' })).sort((a, b) => b.n - a.n);
  }
  // Köp − försäljningar per typ och avkastningsposter (utdelningar, källskatt, skatt, avgifter) under (from, to]
  function weekFlows(tx, an, holdings, from, to, excluded) {
    const ex = new Set(excluded || []), byTx = new Map(), typ = new Map();
    an.groups.forEach(g => { const t = typeOf(g, holdings); typ.set(g, t); g.tx.forEach(r => byTx.set(r, g)); });
    const acc = new Map(accounts(tx).map(a => [a.konto, a]));
    const f = { aktie: 0, fond: 0, etf: 0, utd: 0, avg: 0, sparkonto: 0, n: 0 };
    for (const t of tx) {
      if (ex.has(t.konto)) continue;
      const a = acc.get(t.konto);
      if (a && a.kind === 'sparkonto') { if (t.d <= to) f.sparkonto += t.belopp || 0; continue; }
      if (!(t.d > from && t.d <= to)) continue;
      const c = t.belopp || 0;
      if (SECCASH.has(t.typ) && c) { const g = byTx.get(t); f[g ? typ.get(g) : 'aktie'] += -c; f.n++; }
      else if (RETURN.has(t.typ) && c) { if (DIV.has(t.typ) || t.typ === 'Återbetalning utländsk källskatt') f.utd += c; else f.avg += c; f.n++; }
    }
    for (const k of ['aktie', 'fond', 'etf', 'utd', 'avg', 'sparkonto']) f[k] = Math.round(f[k] * 100) / 100;
    return f;
  }
  // Värde (kr) per typ på de avstängda kontona, från antal per konto × värde per andel i innehavslistan
  function excludedValues(an, holdings, excluded) {
    const out = { aktie: 0, fond: 0, etf: 0 }; if (!holdings || !excluded || !excluded.length) return out;
    an.current.forEach(c => {
      const g = c.group; if (!g || !g.qty) return;
      const per = c.varde / g.qty; let q = 0;
      excluded.forEach(k => { const a = g.acc.get(k); if (a) q += a.qty; });
      if (q > 1e-9) out[c.typ in out ? c.typ : 'aktie'] += q * per;
    });
    for (const k in out) out[k] = Math.round(out[k]);
    return out;
  }


  // ---------------------------------------------------------------- utveckling per innehav mellan två inklistrade listor
  // kr = värde nu − värde då − (köp − sälj däremellan);  % = kr / (värde då + nettoköp om positivt)
  function perf(tx, an, snapA, snapB) {
    // jämför bara grupper (aktier/fonder/ETF) som finns i båda listorna
    const typs = snap => new Set((snap.items || []).map(i => i.typ));
    const tA = typs(snapA), tB = typs(snapB), both = t => tA.has(t) && tB.has(t);
    // samma innehav kan heta lite olika (mobil/dator, förkortat namn): slå ihop via transaktionsgruppen eller början av namnet
    const keyOf = name => { const k = norm(name); const g = an && an.groups.find(g => g.allNames.has(k)); return g ? 'g:' + g.name : k; };
    const sum = snap => { const m = new Map(); (snap.items || []).filter(i => both(i.typ)).forEach(i => { const k = keyOf(i.name); const o = m.get(k) || { name: i.name, typ: i.typ, v: 0, nk: norm(i.name) }; o.v += i.varde; m.set(k, o); }); return m; };
    const A = sum(snapA), B = sum(snapB);
    [...A.keys()].forEach(ka => { if (B.has(ka)) return; const a = A.get(ka);
      const kb = [...B.keys()].find(k => !A.has(k) && B.get(k).typ === a.typ && a.nk.length >= 6 && B.get(k).nk.length >= 6 && (B.get(k).nk.startsWith(a.nk) || a.nk.startsWith(B.get(k).nk)));
      if (kb) { B.set(ka, B.get(kb)); B.delete(kb); } });
    const keys = new Set([...A.keys(), ...B.keys()]), res = [];
    const flowOf = (name) => { const k = norm(name); let f = 0; an.groups.forEach(g => { if (!g.allNames.has(k)) return;
      g.tx.forEach(t => { if (t.d > snapA.date && t.d <= snapB.date && SECCASH.has(t.typ) && t.belopp) f += -t.belopp; }); }); return f; };
    keys.forEach(k => {
      const a = A.get(k), b = B.get(k), va = a ? a.v : 0, vb = b ? b.v : 0, net = flowOf((b || a).name);
      const kr = vb - va - net, base = va + Math.max(net, 0);
      if (!va && !vb) return;
      res.push({ name: (b || a).name, typ: (b || a).typ, kr, pct: base > 0 ? kr / base : null, va, vb, net });
    });
    return res.sort((x, y) => y.kr - x.kr);
  }

  const api = { parseCSV, merge, parseHoldings, analyze, num, accounts, weekFlows, excludedValues, typeOf, perf };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.Avanza = api;
})(this);
