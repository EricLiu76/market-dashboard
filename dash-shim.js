// Stand-in for the claude.ai Dashboard runtime (window.dash), so dashboard_page.html runs unchanged as a
// static site on GitHub Pages: datasets are the CSV files in data/, filters live in the URL query string,
// and d3 is loaded by index.html before this file.
(() => {
  'use strict';

  // Dataset ids the page reads, and the file behind each
  const SOURCES = {
    snapshot: { url: 'data/market_snapshot.csv' },
    master: { url: 'data/master_data_file.csv', parse: wideToLong },
    earnings: { url: 'data/earnings_calendar.csv' },
  };
  // Filter defaults, as in the claude.ai dashboard
  const DEFAULTS = { cmp: null, drange: null, ma200: true, ma50: true, open: null, range: '6M', sort: 'default', spark200: false };
  const COLORS = [1, 2, 3, 4, 5, 6, 7, 8].map((i) => `var(--series-${i})`);
  const PAGE = 'dashboard_page.html';

  const root = document.getElementById('dash-root');
  const result = (status, data = [], columns = [], message = '') =>
    ({ status, data, rows: data, value: data, columns, message, refreshing: false, meta: [] });
  const noBom = (text) => text.replace(/^﻿/, '');

  // master_data_file.csv: rows 1-8 describe each column (type, name, code, frequency, field, currency,
  // unit, source), row 10 is "date" plus names, data from row 11. Gives the rows master_data_long.csv holds.
  function wideToLong(text) {
    const rows = d3.csvParseRows(noBom(text));
    const meta = {};
    let i = 0;
    for (; i < rows.length && rows[i][0] !== 'date'; i++) if (rows[i][0]) meta[rows[i][0]] = rows[i];
    const keys = ['type', 'name', 'code', 'frequency', 'field', 'currency', 'unit', 'source'];
    const out = [];
    for (const r of rows.slice(i + 1)) {
      if (!r[0]) continue;
      for (let j = 1; j < r.length; j++) {
        if (r[j] === '') continue;
        const o = { date: r[0], value: r[j] };
        keys.forEach((k) => (o[k] = meta[k] ? meta[k][j] : ''));
        out.push(o);
      }
    }
    return { rows: out, columns: ['date', 'frequency', 'code', 'name', 'type', 'field', 'value', 'unit', 'currency', 'source'] };
  }

  // ---------- datasets ----------
  const store = new Map();
  function load(id) {
    const src = SOURCES[id];
    const prev = store.get(id);
    store.set(id, prev && prev.status === 'ok' ? { ...prev, refreshing: true } : result('loading'));
    fetch(src.url, { cache: 'no-cache' })
      .then((r) => { if (!r.ok) throw new Error(`${src.url}: HTTP ${r.status}`); return r.text(); })
      .then((text) => {
        const t = src.parse ? src.parse(text) : (() => { const rows = d3.csvParse(noBom(text)); return { rows: Array.from(rows), columns: rows.columns }; })();
        store.set(id, result('ok', t.rows, t.columns));
      })
      .catch((err) => { console.error(err); store.set(id, result('error', [], [], String(err.message || err))); })
      .finally(redraw);
  }

  // dash.calc: worked out again only when an input's data changes
  const calcs = new Map();
  function calcData(id) {
    const c = calcs.get(id);
    const ins = c.inputs.map(data);
    const bad = ins.find((s) => s.status !== 'ok');
    if (bad) return result(bad.status === 'loading' ? 'loading' : 'error');
    if (!c.out || ins.some((s, k) => s.data !== c.ins[k].data)) {
      c.ins = ins;
      try {
        const v = c.fn(...ins.map((s) => s.data));
        c.out = result('ok', v, Array.isArray(v) && v[0] ? Object.keys(v[0]) : []);
      } catch (e) {
        console.error(`dash.calc ${id}`, e);
        c.out = result('error', [], [], String(e.message || e));
      }
    }
    return c.out;
  }

  function data(id) {
    if (calcs.has(id)) return calcData(id);
    if (!SOURCES[id]) return result('missing', [], [], `No dataset named ${id}`);
    if (!store.has(id)) load(id);
    return store.get(id);
  }

  // ---------- redraws ----------
  const draws = [];
  let pending = false;
  const run = (fn) => { try { fn(); } catch (e) { console.error('a dash.onData callback', e); } };
  function redraw() {
    if (pending) return;
    pending = true;
    setTimeout(() => { pending = false; draws.forEach(run); }, 0);
  }
  let lastW = 0, resizeTimer = null;
  new ResizeObserver(([e]) => {
    const w = Math.round(e.contentRect.width);
    if (w === lastW) return;
    lastW = w;
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(redraw, 120);
  }).observe(root);

  // ---------- filters (kept in the URL so a link reopens the same view) ----------
  const fromURL = () => {
    const out = {};
    new URLSearchParams(location.search).forEach((v, k) => (out[k] = v === 'true' ? true : v === 'false' ? false : v));
    return out;
  };
  let params = { ...DEFAULTS, ...fromURL() };
  function toURL() {
    const q = new URLSearchParams();
    Object.entries(params).forEach(([k, v]) => { if (v != null && v !== DEFAULTS[k]) q.set(k, String(v)); });
    const s = q.toString();
    history.replaceState(null, '', location.pathname + (s ? '?' + s : '') + location.hash);
  }
  function setParams(obj) {
    if (!obj || typeof obj !== 'object') return;
    let changed = false;
    for (const [k, v] of Object.entries(obj)) {
      const nv = v === undefined ? null : v;
      if (params[k] !== nv) { params[k] = nv; changed = true; }
    }
    if (changed) { toURL(); redraw(); }
  }

  window.dash = {
    data,
    onData(fn) { if (typeof fn !== 'function') return; draws.push(fn); run(fn); },
    calc(id, { inputs, fn }) { calcs.set(id, { inputs: inputs || [], fn, ins: null, out: null }); },
    loader() { console.warn('dash.loader is not used on the static copy'); },
    colors: COLORS,
    params: () => ({ ...params }),
    setParams,
    resetParams() { params = { ...DEFAULTS }; toURL(); redraw(); },
    setLink(name) { if (typeof name === 'string') history.replaceState(null, '', location.pathname + location.search + '#' + name); },
    refresh(id) { if (SOURCES[id]) load(id); },
  };

  // Start every download now, then put the page in place and run its scripts in order
  Object.keys(SOURCES).forEach((id) => data(id));
  fetch(PAGE, { cache: 'no-cache' })
    .then((r) => { if (!r.ok) throw new Error(`${PAGE}: HTTP ${r.status}`); return r.text(); })
    .then((html) => {
      root.innerHTML = html;
      root.querySelectorAll('script').forEach((old) => {
        const s = document.createElement('script');
        s.textContent = old.textContent;
        old.replaceWith(s);
      });
    })
    .catch((err) => {
      console.error(err);
      const p = document.createElement('p');
      p.className = 'dash-fail';
      p.textContent = 'The dashboard failed to load. Reload the page to try again.';
      root.replaceChildren(p);
    });
})();
