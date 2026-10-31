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
  function parseHoldings(text) {
    const lines = text.split(/\r?\n/).map(l => l.replace(/ /g, ' ')).filter(l => l.trim());
    const items = []; let typ = null, lastName = null;
    const SKIP = /^(utveckling|inköpsinfo|utdelningar|historik|handelsinfo|land|namn|1 dag %|senast|nav|sedan köp( %)?|värde|verktyg|konto)$/i;
    for (const raw of lines) {
      const l = raw.trim();
      if (/^aktier$/i.test(l)) { typ = 'aktie'; continue; }
      if (/^fonder$/i.test(l)) { typ = 'fond'; continue; }
      if (/^(etf:er|etf)$/i.test(l)) { typ = 'etf'; continue; }
      if (/^(certifikat|obligationer|övriga|optioner|warranter)$/i.test(l)) { typ = 'övrigt'; continue; }
      if (SKIP.test(l)) continue;
      if (/^totalt värde/i.test(l)) { lastName = null; continue; }
      const parts = raw.split('\t').map(x => x.trim()).filter(x => x !== '');
      const nums = parts.map(num);
      if (parts.length >= 4 && nums.slice(-4).every(x => x !== null) && lastName) {
        const n = nums.filter(x => x !== null);
        const konto = parts.length > 5 && /^\d{5,}$/.test(parts[0]) ? parts[0] : null;
        const [senast, sedanKop, sedanKopPct, varde] = n.slice(-4);
        items.push({ name: lastName, typ: typ || 'aktie', konto, senast, sedanKop, sedanKopPct: sedanKopPct / 100, varde });
        lastName = null;
      } else if (parts.length === 1 && nums[0] === null) lastName = l;
    }
    if (!items.length) throw new Error('Hittade inga innehav. Markera och kopiera hela innehavslistan från Avanza (Aktier, Fonder och ETF:er).');
    const totals = { aktie: 0, fond: 0, etf: 0, övrigt: 0 };
    items.forEach(i => { totals[i.typ] = (totals[i.typ] || 0) + i.varde; });
    totals.alla = items.reduce((s, i) => s + i.varde, 0);
    return { items, totals };
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
      const g = G.get(gid) || G.set(gid, { id: gid, acc: new Map(), bought: 0, proceeds: 0, removed: 0, div: 0, names: [], tx: [], first: r.d, last: r.d, isins: new Set() }).get(gid);
      g.tx.push(r); g.last = r.d; if (r.isin) g.isins.add(r.isin);
      if (r.namn && (r.typ === 'Köp' || r.typ === 'Sälj')) g.names.push(r.namn);
      if (DIV.has(r.typ)) { g.div += cash; if (r.typ === 'Utdelning') Y(y).utd += cash; else Y(y).kallskatt += cash; continue; }
      // genomsnittsmetoden räknas per konto, som hos Avanza
      const q = r.antal || 0, A = g.acc.get(r.konto) || g.acc.set(r.konto, { qty: 0, cost: 0 }).get(r.konto);
      if (q > 0) { if (cash < 0 && !CONV.has(r.typ)) { A.cost += -cash; g.bought += -cash; } A.qty += q; }
      else if (q < 0) {
        if (CONV.has(r.typ)) A.qty += q;
        else {
          const frac = A.qty > 1e-9 ? Math.min(1, -q / A.qty) : 1, rem = A.cost * frac;
          A.cost -= rem; g.removed += rem; A.qty += q; Y(y).realiserat -= rem;
          if (cash > 0) { g.proceeds += cash; Y(y).realiserat += cash; }
          if (Math.abs(A.qty) < 1e-6) { A.qty = 0; A.cost = 0; }
        }
      } else {
        if (cash < 0) { A.cost += -cash; g.bought += -cash; }
        else if (cash > 0) { g.proceeds += cash; Y(y).realiserat += cash; }
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

  const api = { parseCSV, merge, parseHoldings, analyze, num };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.Avanza = api;
})(this);
