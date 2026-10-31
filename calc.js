/* Beräkningsmotor – samma (rättade) logik som Excel-filen.
   Rådata per vecka: kassa, Avanza (insatt aktier/fonder/ETF, värden, utdelning, realiserad vinst),
   Danske (insatt, värde, realiserad vinst), SAVR (insatt, värde, återbetald avgift), ev. nettolön. */
(function (root) {
  'use strict';
  const num = x => (typeof x === 'number' && isFinite(x)) ? x : 0;
  const DAY = 86400000;
  const t = d => Date.parse(d + 'T00:00:00Z') / DAY;
  const yr = d => +d.slice(0, 4);

  // Lönedag: den 25:e; lördag -> 24:e, söndag -> 26:e.
  // December: den 23:e; lördag -> 22:e, söndag -> 21:e.
  function lonedag(y, m) {
    const d = new Date(Date.UTC(y, m - 1, m === 12 ? 23 : 25)), w = d.getUTCDay();
    if (m === 12) { if (w === 6) d.setUTCDate(22); else if (w === 0) d.setUTCDate(21); }
    else if (w === 6) d.setUTCDate(24); else if (w === 0) d.setUTCDate(26);
    return d.toISOString().slice(0, 10);
  }

  // Dataformat v2: SAVR "inv" = bara egna pengar, återbäringen ("ater") läggs till av appen.
  // (v1/Excel: "inv" innehöll även återbäringen.)
  function migrate(data) {
    if (!data || !Array.isArray(data.rows)) return false;
    if ((data.version || 1) >= 2) return false;
    data.rows.forEach(r => { if (r.sv && r.sv.ater) r.sv.inv = Math.round(((r.sv.inv || 0) - r.sv.ater) * 100) / 100; });
    data.version = 2;
    return true;
  }
  // Lönedagar i intervallet (from, to] (datumsträngar ÅÅÅÅ-MM-DD)
  function lonedagarMellan(from, to) {
    const out = []; let y = +from.slice(0, 4), m = +from.slice(5, 7);
    for (let k = 0; k < 400; k++) {
      const ld = lonedag(y, m);
      if (ld > to) break;
      if (ld > from) out.push(ld);
      if (++m > 12) { m = 1; y++; }
    }
    return out;
  }

  function sortRows(rows) {
    return rows.slice().sort((a, b) => a.d < b.d ? -1 : a.d > b.d ? 1 : 0);
  }

  // Tidsviktad avkastning för ett konto, kedjad vecka för vecka, nollställs vid årsskifte.
  function chainTWR(n, start, val, inv, inc, flowLess) {
    const out = new Array(n).fill(null), wk = new Array(n).fill(null);
    for (let i = 1; i < n; i++) {
      const dInv = inv[i] - inv[i - 1];
      const den = val[i - 1] + Math.max(dInv - (flowLess ? flowLess[i] : 0), 0);
      let w = (val[i] - val[i - 1] - dInv + (inc ? inc[i] : 0)) / den;
      if (!isFinite(w)) w = 0;
      wk[i] = w;
      out[i] = (start[i] === i - 1 || out[i - 1] === null) ? w : (1 + out[i - 1]) * (1 + w) - 1;
    }
    return { year: out, week: wk };
  }

  function compute(rawRows) {
    const rows = sortRows(rawRows), n = rows.length;
    const R = rows.map(r => ({
      d: r.d, kassa: num(r.kassa) + num(r.kassa2), lon: (typeof r.lon === 'number') ? r.lon : null,
      ai: num(r.av && r.av.ai), fi: num(r.av && r.av.fi), ei: num(r.av && r.av.ei),
      aav: num(r.av && r.av.av), afv: num(r.av && r.av.fv), aev: num(r.av && r.av.ev),
      utd: num(r.av && r.av.utd), avg: num(r.av && r.av.avg), avinst: num(r.av && r.av.vinst),
      dinv: num(r.db && r.db.inv), dval: num(r.db && r.db.val), dvinst: num(r.db && r.db.vinst),
      sinv: num(r.sv && r.sv.inv), sval: num(r.sv && r.sv.val), sater: num(r.sv && r.sv.ater)
    }));
    // kumulativa serier
    const c = { Ainv: [], Aai: [], Afi: [], Aei: [], Autd: [], Aown: [], Dinv: [], Sinv: [], SinvG: [], L: [], M: [] };
    let s = { Ainv: 0, Aai: 0, Afi: 0, Aei: 0, Autd: 0, Aown: 0, Dinv: 0, Sinv: 0, SinvG: 0, L: 0, M: 0 };
    R.forEach((r, i) => {
      s.Aai += r.ai; s.Afi += r.fi; s.Aei += r.ei; s.Ainv += r.ai + r.fi + r.ei; s.Autd += r.utd + r.avg; s.Aown += r.ai + r.fi + r.ei - r.utd - r.avg - r.avinst;
      s.Dinv += r.dinv; s.Sinv += r.sinv + r.sater; s.SinvG += r.sinv;
      s.L += r.avinst + r.dvinst;           // realiserad vinst (återinvesterad vinst)
      s.M += r.utd + r.avg + r.sater;       // utdelningar, skatt/avgifter (minus) och återbäring
      for (const k in s) c[k][i] = s[k];
    });
    let dv = 0; const dvCum = R.map(r => (dv += r.dvinst));
    const Aval = R.map(r => r.aav + r.afv + r.aev), Dval = R.map(r => r.dval), Sval = R.map(r => r.sval);
    const N = R.map((r, i) => Aval[i] + Dval[i] + Sval[i]);
    const K = R.map((r, i) => c.Ainv[i] + c.Dinv[i] + c.Sinv[i] - c.L[i] - c.M[i]); // egna insatta pengar
    const V = N.map((x, i) => x - K[i]);
    const Q = R.map(r => r.kassa);
    // startpunkt för året = sista raden före årsskiftet (första raden om inget finns)
    const start = R.map((r, i) => { let j = i - 1; while (j >= 0 && yr(R[j].d) >= yr(r.d)) j--; return j < 0 ? 0 : j; });

    const out = R.map((r, i) => ({ d: r.d, year: yr(r.d), N: N[i], K: K[i], V: V[i], kassa: Q[i], formogenhet: N[i] + Q[i],
      Aval: Aval[i], Dval: Dval[i], Sval: Sval[i], Ainv: c.Ainv[i], Aown: c.Aown[i], Dinv: c.Dinv[i], SinvG: c.SinvG[i], lon: r.lon, utd: r.utd, avg: r.avg, sater: r.sater,
      aav: r.aav, afv: r.afv, aev: r.aev, start: start[i] }));

    for (let i = 0; i < n; i++) {
      const o = out[i], st = start[i];
      o.W = i ? V[i] - V[i - 1] : 0;                                    // avkastning denna vecka (kr)
      const den = i ? N[i - 1] + Math.max(K[i] - K[i - 1], 0) : 0;
      o.X = i && den ? o.W / den : null;                                // veckans utveckling
      o.Y = !i ? null : (st === i - 1 || out[i - 1].Y === null) ? o.X : (1 + out[i - 1].Y) * (1 + o.X) - 1; // året, tidsviktat
      o.U = V[i] - V[st];                                               // avkastat i år (kr)
      // året, pengaviktat (Modified Dietz) – avkastning på de pengar som faktiskt varit insatta
      if (i > st) {
        let wsum = 0; const span = t(R[i].d) - t(R[st].d);
        for (let j = st + 1; j <= i; j++) wsum += (K[j] - K[j - 1]) * (t(R[i].d) - t(R[j].d)) / span;
        o.O = o.U / (N[st] + wsum);
      } else o.O = null;
      o.P = K[i] ? V[i] / K[i] : null;                                  // totalt på egna pengar
      o.sparatAr = (K[i] - K[st]) + (Q[i] - Q[st]);
      o.formAr = o.formogenhet - (N[st] + Q[st]);
      o.avTot = Aval[i] / (c.Ainv[i] - c.Autd[i]) - 1;
      o.dbTot = c.Dinv[i] ? Dval[i] / c.Dinv[i] - 1 : null;
      o.Down = c.Dinv[i] - dvCum[i];
      o.svTot = c.SinvG[i] ? Sval[i] / c.SinvG[i] - 1 : null;
      let utd = 0, ater = 0, avg = 0, lon = 0, lonN = 0;
      for (let j = (st === i ? i : st + 1); j <= i; j++) { utd += R[j].utd; ater += R[j].sater; avg += R[j].avg; if (R[j].lon !== null) { lon += R[j].lon; lonN++; } }
      if (st === 0 && yr(R[0].d) === o.year) { utd += R[0].utd; ater += R[0].sater; }
      o.utdAr = utd; o.aterAr = ater; o.avgAr = avg; o.lonAr = lon; o.lonAntal = lonN;
      o.sparkvotAr = lon ? o.sparatAr / lon : null;
    }
    // Per lön: sparat sedan förra lönen (egna insatta pengar + kassa) / nettolön
    let prevLon = -1;
    for (let i = 0; i < n; i++) {
      const o = out[i];
      if (o.lon === null) continue;
      if (prevLon >= 0) { o.manSparat = (K[i] + Q[i]) - (K[prevLon] + Q[prevLon]); o.manKvot = o.lon ? o.manSparat / o.lon : null; }
      else { o.manSparat = null; o.manKvot = null; }
      prevLon = i;
    }
    out.forEach(o => { o.snittMan = o.lonAntal ? o.sparatAr / o.lonAntal : null; });
    const tw = (val, inv, inc, fl) => chainTWR(n, start, val, inv, inc, fl);
    const av = tw(Aval, c.Ainv, R.map(r => r.utd + r.avg));
    const db = tw(Dval, c.Dinv, R.map(r => r.dvinst), R.map(r => r.dvinst));
    const sv = tw(Sval, c.SinvG);
    const aa = tw(R.map(r => r.aav), c.Aai, R.map(r => r.utd + r.avg));
    const af = tw(R.map(r => r.afv), c.Afi);
    const ae = tw(R.map(r => r.aev), c.Aei);
    out.forEach((o, i) => { o.avY = av.year[i]; o.dbY = db.year[i]; o.svY = sv.year[i];
      o.aktY = aa.year[i]; o.fondY = af.year[i]; o.etfY = ae.year[i]; });
    return out;
  }

  // Sammanfattning per år (sista raden varje år)
  function years(out) {
    const m = new Map();
    out.forEach(o => m.set(o.year, o));
    return [...m.values()];
  }

  const api = { compute, years, sortRows, lonedag, lonedagarMellan, migrate };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.Calc = api;
})(this);
