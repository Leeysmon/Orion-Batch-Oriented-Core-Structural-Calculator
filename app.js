(function () {
  'use strict';
  if (window.OrionLock && window.OrionLock.expired) return;
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => [...r.querySelectorAll(s)];
  const { planeFromAlphaBeta, planeBasis, buildTrace, norm360, D2R } = window.Calc;

  /* ---------------------------------------------------------------- fields */
  const hasMin = (n) => (h) => new RegExp('min\\w*[\\W_]*' + n + '(?!\\d)').test(h) && !/%|pct|percent/.test(h);
  const hasPct = (n) => (h) => new RegExp('min\\w*[\\W_]*' + n + '(?!\\d)').test(h) && /%|pct|percent/.test(h);
  const HOLE = (h) => /^(hole[\W_]*id|holeid|hole|bhid|dh[\W_]*id|drill[\W_]*hole|name)/.test(h);
  const XC = (h) => /(^x|east|(coord|orig)\w*x$)/.test(h);
  const YC = (h) => /(^y|north|(coord|orig)\w*y$)/.test(h);
  const ZC = (h) => /(^z|elev|^rl|(coord|orig)\w*z$)/.test(h);

  const FIELDS = {
    collar: [
      { k: 'holeId', label: 'Hole ID', req: 1, test: HOLE },
      { k: 'x', label: 'X Coord (GPS)', req: 1, test: XC },
      { k: 'y', label: 'Y Coord (GPS)', req: 1, test: YC },
      { k: 'z', label: 'Z Coord (elevation, m.s.l.)', req: 1, test: ZC },
      { k: 'depth', label: 'mDepth (maximum hole depth)', req: 1, test: (h) => /(depth|eoh|length)/.test(h) },
    ],
    survey: [
      { k: 'holeId', label: 'Hole ID', req: 1, test: HOLE },
      { k: 'x', label: 'X Coord (optional)', test: XC },
      { k: 'y', label: 'Y Coord (optional)', test: YC },
      { k: 'z', label: 'Z Coord (optional)', test: ZC },
      { k: 'az', label: 'Azimuth (drillhole trace)', req: 1, test: (h) => /^(az|azi)/.test(h) },
      { k: 'dip', label: 'Dip / inclination (drillhole trace)', req: 1, test: (h) => /^(dip|incl)/.test(h) },
      { k: 'depth', label: 'mDepth (depth of survey)', req: 1, test: (h) => /(depth|dist|^at$|^from)/.test(h) },
      { k: 'size', label: 'Hole size (optional)', test: (h) => /(size|diam|hole[\W_]*size)/.test(h) },
    ],
    structures: [
      { k: 'holeId', label: 'Hole ID', req: 1, test: HOLE },
      { k: 'depth', label: 'mDepth (top or bottom of structure)', req: 1, test: (h) => /(depth|from)/.test(h) && !/len|mark|pos/.test(h) },
      { k: 'mark', label: 'Top/Bottom marker (per row, optional)', test: (h) => /(mark|position)/.test(h) },
      { k: 'oriLine', label: 'Ori-line (solid / dashed)', test: (h) => /ori/.test(h) },
      { k: 'type', label: 'Type', test: (h) => /^(type|struct|feature)/.test(h) },
      { k: 'coreLength', label: 'Core length (cm)', test: (h) => /(core[\W_]*len|^length)/.test(h) },
      { k: 'coreThick', label: 'Core thickness (cm)', test: (h) => /(thick|width)/.test(h) },
      { k: 'alpha', label: 'Alpha', req: 1, test: (h) => /^(alpha|alp|a$)/.test(h) },
      { k: 'beta', label: 'Beta', req: 1, test: (h) => /^(beta|bet|b$)/.test(h) },
      { k: 'gamma', label: 'Gamma (if present)', test: (h) => /^(gamma|gam)/.test(h) },
      ...[1, 2, 3, 4].flatMap((n) => [
        { k: 'min' + n, label: 'Vein Min. ' + n, test: hasMin(n) },
        { k: 'pct' + n, label: 'Vein Min. ' + n + ' %', test: hasPct(n) },
      ]),
      { k: 'style', label: 'Vein style', test: (h) => /style/.test(h) },
      { k: 'comment', label: 'Comment', test: (h) => /(comment|note|remark)/.test(h) },
    ],
  };
  const TITLES = { collar: 'Collar', survey: 'Survey', structures: 'Structures' };

  const state = {
    tables: {}, // key -> {fileName, wb, sheet, headers, rows, map}
    opts: { dipConv: 'neg', depthMode: 'top' },
    results: [], holes: new Map(), prec: 0, selected: -1,
    checked: new Set(), only: null, sort: { col: -1, dir: 1 }, tf: {}, shown: [],
  };

  /* ---------------------------------------------------------------- loading */
  function sheetToTable(wb, sheetName) {
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: true, defval: '' });
    const hi = aoa.findIndex((r) => r.some((c) => String(c).trim() !== ''));
    if (hi < 0) return { headers: [], rows: [] };
    const headers = aoa[hi].map((h, i) => (String(h).trim() ? String(h).trim() : 'Column ' + (i + 1)));
    const rows = aoa.slice(hi + 1).filter((r) => r.some((c) => String(c).trim() !== ''));
    return { headers, rows };
  }

  function autoMap(key, headers) {
    const map = {}, used = new Set();
    const low = headers.map((h) => h.toLowerCase().trim());
    // Assign more specific fields first so generic ones cannot steal their columns.
    const order = FIELDS[key].slice().sort((a, b) => (/^(pct|min)/.test(b.k) ? 1 : 0) - (/^(pct|min)/.test(a.k) ? 1 : 0));
    order.forEach((f) => {
      const i = low.findIndex((h, idx) => !used.has(idx) && f.test(h));
      if (i >= 0) { map[f.k] = i; used.add(i); } else map[f.k] = '';
    });
    return map;
  }

  async function loadFile(key, file) {
    try {
      const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
      const sheet = wb.SheetNames[0];
      const t = { fileName: file.name, wb, sheet, ...sheetToTable(wb, sheet) };
      t.map = autoMap(key, t.headers);
      state.tables[key] = t;
      const box = $(`.drop[data-key=${key}]`);
      box.classList.add('loaded');
      $('.fname', box).textContent = `✓ ${file.name} — ${t.rows.length} rows`;
      renderMaps();
      $('#calcMsg').textContent = '';
    } catch (e) {
      $('#calcMsg').textContent = 'Could not read ' + file.name + ': ' + e.message;
      $('#calcMsg').className = 'msg err';
    }
  }

  document.querySelectorAll('.drop').forEach((box) => {
    const key = box.dataset.key, input = $('input', box);
    $('button', box).onclick = () => input.click();
    input.onchange = () => input.files[0] && loadFile(key, input.files[0]);
    box.ondragover = (e) => { e.preventDefault(); box.classList.add('over'); };
    box.ondragleave = () => box.classList.remove('over');
    box.ondrop = (e) => { e.preventDefault(); box.classList.remove('over'); e.dataTransfer.files[0] && loadFile(key, e.dataTransfer.files[0]); };
  });

  /* ---------------------------------------------------------------- mapping UI */
  const missing = (key) => { const t = state.tables[key]; return FIELDS[key].filter((f) => f.req && (!t || t.map[f.k] === '')); };

  const NUMERIC = {
    collar: ['x', 'y', 'z', 'depth'], survey: ['x', 'y', 'z', 'az', 'dip', 'depth'],
    structures: ['depth', 'alpha', 'beta', 'gamma', 'coreLength', 'coreThick'],
  };
  // Returns a warning if the mapped column is mostly non-numeric text.
  function colWarning(key, k) {
    const t = state.tables[key], c = t && t.map[k];
    if (!t || c === '' || c === undefined || !NUMERIC[key].includes(k)) return '';
    const vals = t.rows.map((r) => r[Number(c)]).filter((v) => String(v ?? '').trim() !== '');
    if (!vals.length) return '';
    const bad = vals.filter((v) => Number.isNaN(num(v)));
    if (bad.length / vals.length < 0.3) return '';
    return `“${t.headers[Number(c)]}” looks like text (e.g. “${String(bad[0]).slice(0, 14)}”), not numbers`;
  }
  const blocking = () => Object.keys(FIELDS).flatMap((key) => (state.tables[key] ? FIELDS[key].filter((f) => f.req && colWarning(key, f.k)).map((f) => `${TITLES[key]} – ${f.label.replace(/ \(.*/, '')}: ${colWarning(key, f.k)}`) : []));

  function updateStatus() {
    Object.keys(FIELDS).forEach((key) => {
      const chip = $(`#chip-${key}`);
      if (chip) {
        const m = missing(key), w = FIELDS[key].filter((f) => f.req && colWarning(key, f.k)).length;
        chip.textContent = m.length ? m.length + ' required missing' : w ? 'Check mapping' : 'Mapped ✓';
        chip.className = 'chip' + (m.length || w ? '' : ' ok');
      }
      document.querySelectorAll(`#maps [data-w^="${key}:"]`).forEach((el) => { el.textContent = colWarning(key, el.dataset.w.split(':')[1]); });
    });
    const probs = blocking();
    $('#btnCalc').disabled = Object.keys(FIELDS).some((k) => !state.tables[k] || missing(k).length) || probs.length > 0;
    const msg = $('#calcMsg');
    if (probs.length) { msg.textContent = '⚠ ' + probs[0]; msg.className = 'msg err'; }
    else if (msg.classList.contains('err') && msg.textContent.startsWith('⚠')) { msg.textContent = ''; msg.className = 'msg'; }
  }

  function renderMaps() {
    const host = $('#maps');
    const keys = Object.keys(FIELDS).filter((k) => state.tables[k]);
    if (!keys.length) { host.innerHTML = '<div class="msg">Load a file in Step 1 to map its columns.</div>'; updateStatus(); return; }
    host.innerHTML = '';
    keys.forEach((key) => {
      const t = state.tables[key];
      const d = document.createElement('details');
      d.className = 'mapcard'; d.open = true;
      const opts = (sel) => '<option value="">— not mapped —</option>' + t.headers.map((h, i) => `<option value="${i}" ${String(sel) === String(i) ? 'selected' : ''}>${esc(h)}</option>`).join('');
      let extra = '';
      if (t.wb.SheetNames.length > 1) {
        extra += `<label class="f">Sheet<select data-sheet>${t.wb.SheetNames.map((s) => `<option ${s === t.sheet ? 'selected' : ''}>${esc(s)}</option>`).join('')}</select></label>`;
      }
      if (key === 'survey') {
        extra += `<label class="f">Dip sign convention<select data-opt="dipConv"><option value="neg" ${state.opts.dipConv === 'neg' ? 'selected' : ''}>Negative = pointing down (e.g. −60)</option><option value="pos" ${state.opts.dipConv === 'pos' ? 'selected' : ''}>Positive = pointing down (e.g. 60)</option></select></label>`;
      }
      if (key === 'structures') {
        extra += `<label class="f">Structure mDepth refers to the<select data-opt="depthMode"><option value="top" ${state.opts.depthMode === 'top' ? 'selected' : ''}>Top-depth (core length extends down-hole)</option><option value="bottom" ${state.opts.depthMode === 'bottom' ? 'selected' : ''}>Bottom-depth (core length extends up-hole)</option></select></label>`;
      }
      d.innerHTML = `<summary>${TITLES[key]} <small style="font-weight:400;color:var(--muted)">${esc(t.fileName)}</small><span class="chip" id="chip-${key}"></span></summary>
        <div class="opts">${extra}</div>
        <div class="mapgrid">${FIELDS[key].map((f) => `<label class="f">${f.label}${f.req ? ' <span class="req">*</span>' : ''}<select data-f="${f.k}">${opts(t.map[f.k])}</select><span class="warn" data-w="${key}:${f.k}"></span></label>`).join('')}</div>
        <div class="preview"><table><thead><tr>${t.headers.map((h) => `<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${t.rows.slice(0, 4).map((r) => `<tr>${t.headers.map((_, i) => `<td>${esc(r[i] ?? '')}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
      d.querySelectorAll('select[data-f]').forEach((s) => (s.onchange = () => { t.map[s.dataset.f] = s.value; updateStatus(); }));
      d.querySelectorAll('select[data-opt]').forEach((s) => (s.onchange = () => { state.opts[s.dataset.opt] = s.value; }));
      const sh = $('select[data-sheet]', d);
      if (sh) sh.onchange = () => { t.sheet = sh.value; Object.assign(t, sheetToTable(t.wb, t.sheet)); t.map = autoMap(key, t.headers); renderMaps(); };
      host.appendChild(d);
    });
    updateStatus();
  }

  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  /* ---------------------------------------------------------------- calculation */
  const num = (v) => {
    if (typeof v === 'number') return v;
    let s = String(v ?? '').trim();
    if (!s) return NaN;
    if (/^-?\d+,\d+$/.test(s)) s = s.replace(',', '.');
    return Number(s);
  };
  const str = (v) => String(v ?? '').trim();
  const hk = (v) => str(v).toLowerCase();

  function calculate() {
    const { collar, survey, structures } = state.tables;
    const get = (t, r, k) => (t.map[k] === '' ? '' : r[Number(t.map[k])]);

    const collars = new Map();
    collar.rows.forEach((r) => {
      const id = hk(get(collar, r, 'holeId'));
      if (id) collars.set(id, { name: str(get(collar, r, 'holeId')), xyz: [num(get(collar, r, 'x')), num(get(collar, r, 'y')), num(get(collar, r, 'z'))], depth: num(get(collar, r, 'depth')) });
    });

    const surveys = new Map();
    survey.rows.forEach((r) => {
      const id = hk(get(survey, r, 'holeId'));
      const depth = num(get(survey, r, 'depth')), az = num(get(survey, r, 'az')), dip = num(get(survey, r, 'dip'));
      if (!id || [depth, az, dip].some(Number.isNaN)) return;
      if (!surveys.has(id)) surveys.set(id, { rows: [], xyz: null });
      const s = surveys.get(id);
      s.rows.push({ depth, az: norm360(az), inc: state.opts.dipConv === 'neg' ? -dip : dip });
      const xyz = [num(get(survey, r, 'x')), num(get(survey, r, 'y')), num(get(survey, r, 'z'))];
      if (!s.xyz && xyz.every((v) => !Number.isNaN(v))) s.xyz = xyz;
    });

    const holes = new Map();
    const holeFor = (id) => {
      if (holes.has(id)) return holes.get(id);
      let h = null;
      const c = collars.get(id), s = surveys.get(id);
      const xyz = c && c.xyz.every((v) => !Number.isNaN(v)) ? c.xyz : s && s.xyz;
      if (xyz && s) {
        h = { name: c ? c.name : id, collar: xyz, trace: buildTrace(xyz, s.rows), maxDepth: c && c.depth > 0 ? c.depth : 0 };
      } else h = { error: !c && !s ? 'Hole not found in Collar or Survey' : !xyz ? 'No valid collar coordinates' : 'No survey stations' };
      holes.set(id, h);
      return h;
    };

    const results = [];
    structures.rows.forEach((r, i) => {
      const g = (k) => get(structures, r, k);
      const id = hk(g('holeId'));
      const res = {
        idx: results.length, row: i + 2, holeId: str(g('holeId')), type: str(g('type')), oriLine: str(g('oriLine')).toLowerCase(),
        style: str(g('style')), comment: str(g('comment')), gamma: num(g('gamma')),
        coreLength: num(g('coreLength')), coreThick: num(g('coreThick')),
        min: [1, 2, 3, 4].map((n) => ({ name: str(g('min' + n)), pct: num(g('pct' + n)) })),
        alpha: num(g('alpha')), beta: num(g('beta')),
      };
      const depth = num(g('depth')), len = Number.isNaN(res.coreLength) ? 0 : res.coreLength / 100;
      const mk = str(g('mark')).toLowerCase();
      const mode = mk.startsWith('b') ? 'bottom' : mk.startsWith('t') ? 'top' : state.opts.depthMode;
      const h = holeFor(id);
      if (h.error) res.error = h.error;
      else if (Number.isNaN(depth)) res.error = `Invalid depth “${str(g('depth'))}”`;
      else if (Number.isNaN(res.alpha) || Number.isNaN(res.beta)) res.error = 'Missing alpha/beta';
      else if (res.alpha < 0 || res.alpha > 90) res.error = 'Alpha outside 0–90°';
      else if (res.beta < 0 || res.beta > 360) res.error = 'Beta outside 0–360°';
      if (!res.error) {
        res.top = mode === 'top' ? depth : depth - len;
        res.bottom = mode === 'top' ? depth + len : depth;
        res.depth = (res.top + res.bottom) / 2;
        const at = h.trace.at(res.depth);
        const p = planeFromAlphaBeta(at.dir, at.az, res.alpha, res.beta % 360);
        Object.assign(res, { holeAz: at.az, holeInc: at.inc, pos: at.pos, normal: p.normal, dip: p.dip, dipDir: p.dipDir, strike: p.strike });
        h.maxDepth = Math.max(h.maxDepth, res.bottom);
      }
      results.push(res);
    });

    state.holes = holes;
    state.results = results;
    state.selected = -1;
    state.checked = new Set(); state.only = null; state.fscope = null; state.sort = { col: -1, dir: 1 }; state.tf = {};
    return results;
  }

  /* ---------------------------------------------------------------- formatting & table */
  const fmt = (v) => (Number.isFinite(v) ? v.toFixed(state.prec) + '°' : '—');
  const fmt3 = (v) => (Number.isFinite(v) ? v.toFixed(state.prec).padStart(state.prec ? 4 + state.prec : 3, '0') : '—');
  const pad = (v) => (Number.isFinite(v) ? Math.round(v * Math.pow(10, state.prec)) / Math.pow(10, state.prec) : v);

  function renderTable() {
    const cols = TCOLS();
    const arrow = (i) => (state.sort.col === i ? (state.sort.dir > 0 ? ' ▲' : ' ▼') : '');
    $('#tbl').innerHTML = `<thead><tr><th class="ck"><input type="checkbox" id="chkAll" title="Check / uncheck all rows currently listed"></th>${cols.map((c, i) => `<th class="sortable" data-c="${i}" title="Click to sort">${c.h}<span>${arrow(i)}</span></th>`).join('')}</tr>
      <tr class="frow"><th class="ck"></th>${cols.map((c, i) => `<th><input type="text" data-f="${i}" placeholder="filter" value="${esc(state.tf[i] || '')}" title="Text contains · numbers: &gt;10, &lt;=5, =3, 10..20"></th>`).join('')}</tr></thead><tbody></tbody>`;
    $('#tbl thead').onclick = (e) => {
      const th = e.target.closest('th.sortable');
      if (!th || e.target.tagName === 'INPUT') return;
      const i = Number(th.dataset.c);
      state.sort = state.sort.col !== i ? { col: i, dir: 1 } : state.sort.dir > 0 ? { col: i, dir: -1 } : { col: -1, dir: 1 };
      $$('#tbl th.sortable').forEach((t) => { t.lastChild.textContent = arrow(Number(t.dataset.c)); });
      renderBody();
    };
    $('#tbl thead').oninput = (e) => {
      const f = e.target.dataset.f;
      if (f !== undefined) { state.tf[f] = e.target.value; renderBody(); }
    };
    $('#chkAll').onchange = (e) => {
      state.shown.forEach((r) => { if (!r.error) e.target.checked ? state.checked.add(r.idx) : state.checked.delete(r.idx); });
      renderBody();
    };
    $('#tbl tbody').onclick = (e) => {
      const tr = e.target.closest('tr');
      if (!tr) return;
      const i = Number(tr.dataset.i);
      if (e.target.matches('input[type=checkbox]')) {
        e.target.checked ? state.checked.add(i) : state.checked.delete(i);
        tr.classList.toggle('chk', e.target.checked); updateTableUi();
      } else select(i, true);
    };
    renderBody();
    $('.tablewrap').style.setProperty('--thh', $('#tbl thead tr').offsetHeight - 1 + 'px');
  }

  function TCOLS() {
    const n = (v) => (Number.isFinite(v) ? v : '');
    return [
      { h: '#', v: (r) => r.row, n: 1 },
      { h: 'Hole', v: (r) => r.holeId },
      { h: 'Depth (m)', v: (r) => (r.error ? '' : r.depth), n: 1, f: (r) => (r.error ? '' : r.depth.toFixed(2)) },
      { h: 'Type', v: (r) => r.type },
      { h: 'Ori-line', v: (r) => r.oriLine },
      { h: 'Vein Min. 1', v: (r) => r.min[0].name },
      { h: 'Alpha', v: (r) => n(r.alpha), n: 1 },
      { h: 'Beta', v: (r) => n(r.beta), n: 1 },
      { h: 'Dip', v: (r) => (r.error ? '' : r.dip), n: 1, f: (r) => (r.error ? '' : fmt(r.dip)) },
      { h: 'Dip dir.', v: (r) => (r.error ? '' : r.dipDir), n: 1, f: (r) => (r.error ? '' : fmt(r.dipDir)) },
      { h: 'Strike', v: (r) => (r.error ? '' : r.strike), n: 1, f: (r) => (r.error ? '' : fmt(r.strike)) },
      { h: 'Hole az / inc', v: (r) => (r.error ? '' : r.holeAz), n: 1, f: (r) => (r.error ? '' : r.holeAz.toFixed(1) + ' / ' + r.holeInc.toFixed(1)) },
      { h: 'Status', v: (r) => r.error || 'OK' },
    ];
  }

  function rowMatches(c, r, q) {
    q = q.trim();
    if (!q) return true;
    const v = c.v(r), txt = String(c.f ? c.f(r) : v).toLowerCase();
    const one = (t) => {
      t = t.trim();
      if (!t) return null;
      let m;
      if (c.n && (m = t.match(/^(>=|<=|!=|>|<|=)\s*(-?\d+(?:\.\d+)?)$/))) {
        if (v === '') return false;
        const x = Number(m[2]);
        return { '>': v > x, '<': v < x, '>=': v >= x, '<=': v <= x, '=': v === x, '!=': v !== x }[m[1]];
      }
      if (c.n && (m = t.match(/^(-?\d+(?:\.\d+)?)\s*\.\.\s*(-?\d+(?:\.\d+)?)$/))) return v !== '' && v >= Math.min(+m[1], +m[2]) && v <= Math.max(+m[1], +m[2]);
      return txt.includes(t.toLowerCase());
    };
    // "a | b, c" = any of; terms starting with "!" exclude
    const terms = q.split(/[|,;]/).map((t) => t.trim()).filter(Boolean);
    const inc = terms.filter((t) => t[0] !== '!'), exc = terms.filter((t) => t[0] === '!' && t.length > 1).map((t) => t.slice(1));
    if (exc.some((t) => one(t))) return false;
    return !inc.length || inc.some((t) => one(t));
  }

  function renderBody() {
    const cols = TCOLS();
    let rows = state.results.filter((r) => cols.every((c, i) => rowMatches(c, r, state.tf[i] || '')));
    const active = Object.values(state.tf).some((v) => String(v).trim());
    const next = active ? new Set(rows.filter((r) => !r.error).map((r) => r.idx)) : null;
    const same = (!next && !state.fscope) || (next && state.fscope && next.size === state.fscope.size && [...next].every((i) => state.fscope.has(i)));
    state.fscope = next;
    if (!same && V.inited && V.items.length) V.refresh();
    if (state.sort.col >= 0) {
      const c = cols[state.sort.col], d = state.sort.dir;
      rows = rows.slice().sort((a, b) => {
        const x = c.v(a), y = c.v(b);
        if (x === '' && y === '') return 0;
        if (x === '') return 1;
        if (y === '') return -1;
        return (c.n ? x - y : String(x).localeCompare(String(y), undefined, { numeric: true, sensitivity: 'base' })) * d;
      });
    }
    state.shown = rows;
    $('#tbl tbody').innerHTML = rows.map((r) => `<tr data-i="${r.idx}" class="${r.error ? 'bad' : ''} ${r.idx === state.selected ? 'sel' : ''} ${state.checked.has(r.idx) ? 'chk' : ''}">
      <td class="ck"><input type="checkbox" ${r.error ? 'disabled' : ''} ${state.checked.has(r.idx) ? 'checked' : ''}></td>${cols.map((c) => `<td>${esc(String(c.f ? c.f(r) : c.v(r)))}</td>`).join('')}</tr>`).join('');
    updateTableUi();
  }

  function updateTableUi() {
    const ok = state.results.filter((r) => !r.error).length;
    const nChk = state.checked.size, listed = state.shown.filter((r) => !r.error);
    $('#tblMsg').textContent = `${ok} calculated, ${state.results.length - ok} skipped · listing ${state.shown.length} of ${state.results.length} · ${nChk} checked`;
    const all = $('#chkAll');
    if (all) { all.checked = listed.length > 0 && listed.every((r) => state.checked.has(r.idx)); all.indeterminate = !all.checked && listed.some((r) => state.checked.has(r.idx)); }
    $('#btnShowSel').classList.toggle('hidden', !nChk);
    $('#btnShowSel').textContent = `Show in viewer (${nChk})`;
    $('#btnUncheck').classList.toggle('hidden', !nChk);
    $('#btnShowAll').classList.toggle('hidden', !state.only);
    $('#btnClearTf').classList.toggle('hidden', !(state.sort.col >= 0 || Object.values(state.tf).some((v) => v)));
    $('#onlyNote').textContent = [state.fscope ? `Viewers follow the table filters (${state.fscope.size} structure(s)).` : '', state.only ? `Limited to ${state.only.size} checked structure(s).` : ''].filter(Boolean).join(' ');
  }

  function select(i, fly) {
    state.selected = i;
    document.querySelectorAll('#tbl tr.sel').forEach((e) => e.classList.remove('sel'));
    const tr = $(`#tbl tr[data-i="${i}"]`);
    if (tr) {
      tr.classList.add('sel');
      if (fly) { const w = $('.tablewrap'), hh = $('#tbl thead').offsetHeight; w.scrollTo({ top: Math.max(0, tr.offsetTop - hh - w.clientHeight / 2 + tr.offsetHeight), behavior: 'smooth' }); }
    }
    const r = state.results[i];
    const ok = r && !r.error;
    $('#selTitle').textContent = r ? `${r.holeId} @ ${ok ? r.depth.toFixed(2) + ' m' : '—'}` : '—';
    $('#rDip').textContent = ok ? fmt(r.dip) : '—';
    $('#rDd').textContent = ok ? fmt(r.dipDir) : '—';
    $('#rSt').textContent = ok ? fmt(r.strike) : '—';
    $('#rSD').textContent = ok ? `${fmt3(r.strike)} / ${fmt3(r.dip)}` : '— / —';
    $('#rInfo').textContent = !r ? 'Click a plane or a table row.' : ok ? `${r.type || 'structure'} · α ${r.alpha}° β ${r.beta}° · hole ${r.holeAz.toFixed(1)}° / ${r.holeInc.toFixed(1)}°` : r.error;
    drawCompass(ok ? r : null);
    V.updateSelection();
    SN.draw();
  }

  /* ---------------------------------------------------------------- compass */
  function drawCompass(r) {
    const C = 120, R = 92;
    const pt = (az, k) => [C + R * k * Math.sin(az * D2R), C - R * k * Math.cos(az * D2R)];
    const line = (a, b, col, w, extra = '') => `<line x1="${a[0]}" y1="${a[1]}" x2="${b[0]}" y2="${b[1]}" stroke="${col}" stroke-width="${w}" stroke-linecap="round" ${extra}/>`;
    let s = `<svg viewBox="0 0 240 240" width="240" height="240" role="img"><defs>
      <marker id="ah1" markerWidth="8" markerHeight="8" refX="6" refY="4" orient="auto"><path d="M0,0 L8,4 L0,8 z" fill="#cfe3ff"/></marker>
      <marker id="ah2" markerWidth="8" markerHeight="8" refX="6" refY="4" orient="auto"><path d="M0,0 L8,4 L0,8 z" fill="#e8a33d"/></marker></defs>
      <circle cx="${C}" cy="${C}" r="${R}" fill="#0b182a" stroke="#3a5a80" stroke-width="2"/>`;
    for (let a = 0; a < 360; a += 30) s += line(pt(a, 0.93), pt(a, 1), '#5f7fa6', 1.2);
    [['N', 0], ['E', 90], ['S', 180], ['W', 270]].forEach(([t, a]) => { const p = pt(a, 1.14); s += `<text x="${p[0]}" y="${p[1] + 5}" text-anchor="middle" font-size="14" font-weight="700" fill="#cfe3ff">${t}</text>`; });
    if (r) {
      s += line(pt(r.holeAz, 0), pt(r.holeAz, 0.9), '#cfe3ff', 3, 'marker-end="url(#ah1)" stroke-dasharray="1 0"');
      s += line(pt(r.strike, -0.85), pt(r.strike, 0.85), '#4da3ff', 5);
      const mid = pt(r.dipDir, 0.5);
      s += line([C, C], mid, '#e8a33d', 3, 'marker-end="url(#ah2)"');
      const lab = pt(r.dipDir, 0.7);
      s += `<text x="${lab[0]}" y="${lab[1] + 4}" text-anchor="middle" font-size="13" font-weight="700" fill="#b87a14">${Math.round(r.dip)}°</text>`;
      s += `<circle cx="${C}" cy="${C}" r="4" fill="#d33"/>`;
    } else s += `<text x="${C}" y="${C + 4}" text-anchor="middle" font-size="12" fill="#8aa">Awaiting result</text>`;
    $('#compass').innerHTML = s + '</svg>';
  }

  /* ---------------------------------------------------------------- 3D viewer */
  const PALETTE = ['#d64545', '#2e6fd8', '#2bab6a', '#e8a33d', '#8e5bd1', '#17a2b8', '#e2639e', '#7f8c3a', '#a0522d', '#4b5d67', '#f0d000', '#00897b'];
  const V = {
    inited: false, items: [], hidden: new Set(), cats: [], catColor: new Map(), hole: 'all', field: 'min1',
    depMin: 0, depMax: 0, bounds: [0, 0], origin: [0, 0, 0], extent: 100,
  };
  const to3 = (p) => new THREE.Vector3(p[0] - V.origin[0], p[2] - V.origin[2], -(p[1] - V.origin[1]));

  const catName = (r) => {
    let v = '';
    switch (V.field) {
      case 'min1': v = r.min[0].name; break;
      case 'type': v = r.type; break;
      case 'oriLine': v = r.oriLine; break;
      case 'holeId': v = r.holeId; break;
      case 'style': v = r.style; break;
    }
    return v || '(blank)';
  };

  V.init = function () {
    if (V.inited) return;
    V.inited = true;
    const canvas = $('#gl');
    V.renderer = new THREE.WebGLRenderer({ canvas, antialias: true, preserveDrawingBuffer: true });
    V.renderer.setPixelRatio(Math.min(2, window.devicePixelRatio || 1));
    V.scene = new THREE.Scene();
    V.scene.background = new THREE.Color('#0a1a2e');
    V.camera = new THREE.PerspectiveCamera(45, 1, 0.1, 1e6);
    V.controls = new THREE.OrbitControls(V.camera, canvas);
    V.controls.enableDamping = true;
    // Pan = left + right mouse buttons together (right alone does nothing; Shift/Ctrl-drag pan is disabled).
    V.controls.mouseButtons = { LEFT: THREE.MOUSE.ROTATE, MIDDLE: THREE.MOUSE.DOLLY, RIGHT: -1 };
    canvas.addEventListener('pointerdown', (e) => { if (e.shiftKey || e.ctrlKey || e.metaKey) e.stopImmediatePropagation(); }, true);
    canvas.addEventListener('pointermove', (e) => {
      if (e.pointerType !== 'mouse' || e.buttons !== 3) return;
      const cam = V.camera, tg = V.controls.target;
      const k = 2 * cam.position.distanceTo(tg) * Math.tan(cam.fov * D2R / 2) / Math.max(1, canvas.clientHeight);
      cam.updateMatrixWorld();
      const m = cam.matrixWorld.elements;
      const d = new THREE.Vector3(m[0], m[1], m[2]).multiplyScalar(-e.movementX * k)
        .add(new THREE.Vector3(m[4], m[5], m[6]).multiplyScalar(e.movementY * k));
      cam.position.add(d); tg.add(d); V.controls.update();
    });
    V.world = new THREE.Group(); V.scene.add(V.world);
    V.planes = new THREE.Group(); V.holesG = new THREE.Group(); V.selG = new THREE.Group(); V.axesG = new THREE.Group();
    V.world.add(V.holesG, V.planes, V.selG, V.axesG);
    const loop = () => { V.controls.update(); V.updateBorders(); V.renderer.render(V.scene, V.camera); requestAnimationFrame(loop); };
    new ResizeObserver(() => {
      const w = canvas.clientWidth, h = canvas.clientHeight;
      V.renderer.setSize(w, h, false); V.camera.aspect = w / Math.max(1, h); V.camera.updateProjectionMatrix();
    }).observe(canvas);
    loop();

    let down = null;
    canvas.addEventListener('pointerdown', (e) => { down = [e.clientX, e.clientY]; });
    canvas.addEventListener('pointerup', (e) => {
      if (!down || Math.hypot(e.clientX - down[0], e.clientY - down[1]) > 4) return;
      const b = canvas.getBoundingClientRect();
      const ndc = new THREE.Vector2(((e.clientX - b.left) / b.width) * 2 - 1, -((e.clientY - b.top) / b.height) * 2 + 1);
      const rc = new THREE.Raycaster(); rc.setFromCamera(ndc, V.camera);
      const meshes = V.items.filter((it) => it.group.visible).flatMap((it) => [it.mesh, it.back]);
      const hit = rc.intersectObjects(meshes, false)[0];
      if (hit) select(hit.object.userData.idx, true);
    });
  };

  V.visibleHoles = () => [...state.holes.values()].filter((h) => !h.error && (V.hole === 'all' || h.name.toLowerCase() === V.hole));

  V.computeExtent = function () {
    const box = new THREE.Box3();
    const pts = [];
    V.visibleHoles().forEach((h) => { pts.push(h.collar, h.trace.at(Math.max(h.maxDepth, 1)).pos); });
    if (!pts.length) pts.push([0, 0, 0]);
    pts.forEach((p) => box.expandByPoint(new THREE.Vector3(p[0], p[2], -p[1])));
    const size = new THREE.Vector3(); box.getSize(size);
    V.extent = Math.max(10, size.x, size.y, size.z);
    V.box = box;
  };

  V.planeSize = () => V.extent * 0.001 * Number($('#rngSize').value);

  V.rebuild = function () {
    V.init();
    const clear = (g) => { while (g.children.length) { const c = g.children.pop(); c.traverse((o) => { o.geometry && o.geometry.dispose(); o.material && (o.material.map && o.material.map.dispose(), o.material.dispose()); }); } };
    clear(V.planes); clear(V.holesG);
    V.items = [];
    // fixed origin so geometry stays stable while filtering
    const all = [...state.holes.values()].filter((h) => !h.error);
    if (!all.length) return;
    const c0 = all[0].collar;
    V.origin = [c0[0], c0[1], c0[2]];
    V.computeExtent();
    const S = V.planeSize(), half = S / 2;
    $('#szVal').textContent = S.toFixed(1) + ' m';

    all.forEach((h) => {
      const pts = [];
      const step = Math.max(1, h.maxDepth / 200);
      for (let d = 0; d < h.maxDepth + step; d += step) pts.push(to3(h.trace.at(Math.min(d, h.maxDepth)).pos));
      const line = new THREE.Line(new THREE.BufferGeometry().setFromPoints(pts), new THREE.LineBasicMaterial({ color: 0xcfe3ff }));
      line.userData.hole = h.name.toLowerCase();
      const dot = new THREE.Mesh(new THREE.SphereGeometry(Math.max(0.4, V.extent * 0.008), 12, 12), new THREE.MeshBasicMaterial({ color: 0xdd2222 }));
      dot.position.copy(to3(h.collar)); dot.userData.hole = line.userData.hole;
      const cv = document.createElement('canvas'); cv.width = 256; cv.height = 64;
      const cx = cv.getContext('2d'); cx.font = 'bold 34px sans-serif'; cx.textAlign = 'center'; cx.fillStyle = '#e3eefc'; cx.strokeStyle = '#0a1a2e'; cx.lineWidth = 6;
      cx.strokeText(h.name, 128, 44); cx.fillText(h.name, 128, 44);
      const sp = new THREE.Sprite(new THREE.SpriteMaterial({ map: new THREE.CanvasTexture(cv), depthTest: false, transparent: true }));
      const k = V.extent * 0.12; sp.scale.set(k, k / 4, 1);
      sp.position.copy(to3(h.collar)).add(new THREE.Vector3(0, V.extent * 0.03, 0)); sp.userData.hole = line.userData.hole; sp.userData.label = true;
      V.holesG.add(line, dot, sp);
    });

    const op = Number($('#rngOpacity').value) / 100;
    state.results.filter((r) => !r.error).forEach((r) => {
      const b = planeBasis(r.dipDir, r.dip), c = r.pos;
      const corner = (a, d) => to3([c[0] + b.strike[0] * a * half + b.down[0] * d * half, c[1] + b.strike[1] * a * half + b.down[1] * d * half, c[2] + b.strike[2] * a * half + b.down[2] * d * half]);
      const v = [corner(1, 1), corner(-1, 1), corner(-1, -1), corner(1, -1)];
      const geo = new THREE.BufferGeometry().setFromPoints(v);
      // Wind the quad so its front face is the upper (hanging-wall) side.
      const nUp = new THREE.Vector3(r.normal[0], r.normal[2], -r.normal[1]);
      const fn = new THREE.Vector3().subVectors(v[1], v[0]).cross(new THREE.Vector3().subVectors(v[2], v[0]));
      geo.setIndex(fn.dot(nUp) >= 0 ? [0, 1, 2, 0, 2, 3] : [0, 2, 1, 0, 3, 2]); geo.computeVertexNormals();
      const mk = (side) => new THREE.Mesh(geo, new THREE.MeshBasicMaterial({ color: 0xcccccc, transparent: true, opacity: op, side, depthWrite: false }));
      const mesh = mk(THREE.FrontSide), back = mk(THREE.BackSide);
      mesh.userData.idx = back.userData.idx = r.idx;
      const dashed = r.oriLine.startsWith('dash');
      const lg = new THREE.BufferGeometry().setFromPoints(v);
      const lm = new THREE.LineBasicMaterial({ color: 0x333333 });
      const dm = new THREE.LineDashedMaterial({ color: 0x333333, dashSize: S / 60, gapSize: S / 60 });
      const outline = new THREE.LineLoop(lg, lm);
      const outlineDot = new THREE.LineLoop(lg, dm);
      outlineDot.computeLineDistances(); outlineDot.visible = false;
      const group = new THREE.Group(); group.add(mesh, back, outline, outlineDot);
      V.planes.add(group);
      const centre = v.reduce((a, p) => a.add(p), new THREE.Vector3()).multiplyScalar(0.25);
      V.items.push({ res: r, group, mesh, back, outline, outlineDot, verts: v, nUp, centre });
    });
    V.buildAxes();
    V.recolor();
  };

  V.buildAxes = function () {
    while (V.axesG.children.length) { const c = V.axesG.children.pop(); c.traverse((o) => { o.geometry && o.geometry.dispose(); o.material && (o.material.map && o.material.map.dispose(), o.material.dispose()); }); }
    const b = V.box, off = new THREE.Vector3(V.origin[0], V.origin[2], -V.origin[1]);
    V.axesBox = b.clone();
    const size = new THREE.Vector3(); b.getSize(size);
    const minL = V.extent * 0.25;
    // Axes start at the south-west, lowest corner of the hole extents.
    // The north line runs from just south of the first collar, offset west by a gap; the elevation line stands at its
    // south end, so no axis touches the trace. The elevation line reaches up to the collar level.
    const gap = V.extent * 0.03;
    const bmin = b.min.clone().sub(off), bmax = b.max.clone().sub(off);
    const o = new THREE.Vector3(-gap, 0, gap);
    const oz = new THREE.Vector3(-gap, Math.min(bmin.y, 0) - gap, gap);
    const top = Math.max(0, bmax.y);
    const sc = V.extent * 0.012;
    const NICE = [1, 2, 5, 10, 20, 25, 50, 100];
    const pickStep = (L) => {
      const s = NICE.find((n) => L / n <= 8);
      if (s) return s;
      let m = 100; for (;;) { for (const k of [2, 5, 10]) if (L / (m * k) <= 8) return m * k; m *= 10; }
    };
    const tickOf = { x: new THREE.Vector3(0, 0, 1), y: new THREE.Vector3(-1, 0, 0), z: new THREE.Vector3(-1, 0, 0) };
    const defs = [
      { o: oz, dir: new THREE.Vector3(1, 0, 0), L: Math.max(gap + Math.max(bmax.x, 0), minL), col: '#ff6b6b', name: 'X (East)' },
      { o: oz, dir: new THREE.Vector3(0, 0, -1), L: Math.max(gap + Math.max(-bmin.z, 0), minL), col: '#5be08f', name: 'Y (North)' },
      { o: oz, dir: new THREE.Vector3(0, 1, 0), L: Math.max(top - oz.y, minL * 0.5), col: '#6db2ff', name: 'Z (Elev.)' },
    ];
    defs.forEach((d) => {
      const o = d.o;
      const L = d.L, end = o.clone().addScaledVector(d.dir, L);
      const abs = end.clone().add(off);
      V.axesBox.expandByPoint(abs); V.axesBox.expandByPoint(o.clone().add(off));
      const val = d.dir.x ? abs.x : d.dir.y ? abs.y : -abs.z;
      V.axesG.add(new THREE.Line(new THREE.BufferGeometry().setFromPoints([o, end]), new THREE.LineBasicMaterial({ color: d.col })));
      const tip = new THREE.Mesh(new THREE.ConeGeometry(sc * 0.9, sc * 3, 12), new THREE.MeshBasicMaterial({ color: d.col }));
      tip.position.copy(end); tip.quaternion.setFromUnitVectors(new THREE.Vector3(0, 1, 0), d.dir);
      V.axesG.add(tip);
      V.axesG.add(V.label(d.name, d.col, end.clone().addScaledVector(d.dir, sc * 3), V.extent * 0.22));
      const key = d.dir.x ? 'x' : d.dir.y ? 'z' : 'y';
      const sgn = d.dir.z ? -1 : 1;
      const start = sgn * (d.dir.x ? o.x + off.x : d.dir.y ? o.y + off.y : o.z + off.z);
      const step = pickStep(L), tv = tickOf[key], tl = sc * 2.5;
      const pts = [];
      for (let v = Math.ceil(start / step) * step; v <= start + L + 1e-9; v += step) {
        const p = o.clone().addScaledVector(d.dir, v - start);
        pts.push(p, p.clone().addScaledVector(tv, tl));
        V.axesG.add(V.label(String(v), d.col, p.clone().addScaledVector(tv, tl * 3), V.extent * 0.17, true));
      }
      if (pts.length) V.axesG.add(new THREE.LineSegments(new THREE.BufferGeometry().setFromPoints(pts), new THREE.LineBasicMaterial({ color: d.col })));
    });
    V.axesG.visible = $('#chkAxes').checked;
  };

  V.label = function (text, col, pos, width, small) {
    const cv = document.createElement('canvas'); cv.width = 512; cv.height = 96;
    const g = cv.getContext('2d'); g.font = `bold ${small ? 52 : 44}px sans-serif`; g.textAlign = 'center'; g.textBaseline = 'middle';
    g.lineWidth = 8; g.strokeStyle = '#0a1a2e'; g.strokeText(text, 256, 48); g.fillStyle = col; g.fillText(text, 256, 48);
    const sp = new THREE.Sprite(new THREE.SpriteMaterial({ map: new THREE.CanvasTexture(cv), depthTest: false, transparent: true }));
    sp.scale.set(width, width * 96 / 512, 1); sp.position.copy(pos); sp.renderOrder = 10;
    return sp;
  };

  // Solid border when the hanging-wall (upper) side faces the camera, dotted for the foot wall.
  V.updateBorders = function () {
    if (!V.items) return;
    const cp = V.camera.position, tmp = new THREE.Vector3();
    V.items.forEach((it) => {
      if (!it.group.visible) return;
      const hw = tmp.subVectors(cp, it.centre).dot(it.nUp) >= 0;
      it.outline.visible = hw; it.outlineDot.visible = !hw;
    });
  };

  V.recolor = function () {
    const counts = new Map();
    state.results.filter((r) => !r.error).forEach((r) => { const n = catName(r); const k = n.toLowerCase(); const e = counts.get(k) || { name: n, n: 0 }; e.n++; counts.set(k, e); });
    V.cats = [...counts.entries()].sort((a, b) => b[1].n - a[1].n).map(([k, e]) => ({ key: k, name: e.name }));
    V.catColor = new Map(V.cats.map((c, i) => [c.key, PALETTE[i % PALETTE.length]]));
    V.hidden = new Set();
    V.items.forEach((it) => {
      it.cat = catName(it.res).toLowerCase();
      const c = V.catColor.get(it.cat);
      it.mesh.material.color.set(c); it.back.material.color.set(c).multiplyScalar(0.55);
      it.outline.material.color.set(c).multiplyScalar(0.55);
      it.outlineDot.material.color.set(c).multiplyScalar(0.55);
    });
    V.refresh();
  };

  V.refresh = function () {
    V.items.forEach((it) => {
      const r = it.res;
      it.inScope = (V.hole === 'all' || r.holeId.toLowerCase() === V.hole) && r.depth >= V.depMin - 1e-9 && r.depth <= V.depMax + 1e-9 && (!state.only || state.only.has(r.idx)) && (!state.fscope || state.fscope.has(r.idx));
      it.group.visible = it.inScope && !V.hidden.has(it.cat);
    });
    V.holesG.children.forEach((o) => { o.visible = (V.hole === 'all' || o.userData.hole === V.hole) && (!o.userData.label || $('#chkLabels').checked); });
    V.legend();
    V.updateSelection();
    SN.draw();
  };

  V.legend = function () {
    $('#legend').classList.toggle('hidden', !$('#chkLegend').checked);
    const counts = new Map();
    V.items.forEach((it) => { if (it.inScope) counts.set(it.cat, (counts.get(it.cat) || 0) + 1); });
    const title = $('#selColor').selectedOptions[0].textContent;
    ['#legend', '#slegend'].forEach((sel) => {
      const el = $(sel);
      el.innerHTML = `<h4><span>${esc(title)}</span><span><a data-all="1">all</a> · <a data-all="0">none</a></span></h4>` + V.cats.map((c) =>
        `<div class="leg ${V.hidden.has(c.key) ? 'off' : ''}" data-k="${esc(c.key)}"><i style="background:${V.catColor.get(c.key)}"></i><span>${esc(c.name)}</span><em>${counts.get(c.key) || 0}</em></div>`).join('') +
        (sel === '#legend' ? '<div class="leg note"><i></i><span>Light = hanging wall<br>Dark = foot wall</span></div>' : '');
      el.querySelectorAll('.leg[data-k]').forEach((d) => (d.onclick = () => { const k = d.dataset.k; V.hidden.has(k) ? V.hidden.delete(k) : V.hidden.add(k); V.refresh(); }));
      el.querySelectorAll('a[data-all]').forEach((a) => (a.onclick = () => { V.hidden = a.dataset.all === '1' ? new Set() : new Set(V.cats.map((c) => c.key)); V.refresh(); }));
    });
  };

  V.updateSelection = function () {
    if (!V.inited) return;
    while (V.selG.children.length) { const c = V.selG.children.pop(); c.geometry.dispose(); c.material.dispose(); }
    const it = V.items.find((x) => x.res.idx === state.selected);
    if (!it || !it.group.visible) return;
    const m = new THREE.LineLoop(new THREE.BufferGeometry().setFromPoints(it.verts), new THREE.LineBasicMaterial({ color: 0x000000 }));
    const mid = it.verts.reduce((a, p) => a.add(p), new THREE.Vector3()).multiplyScalar(0.25);
    m.geometry.setFromPoints(it.verts.map((p) => p.clone().sub(mid).multiplyScalar(1.12).add(mid)));
    V.selG.add(m);
  };

  V.setDepthBounds = function () {
    const rs = state.results.filter((r) => !r.error && (V.hole === 'all' || r.holeId.toLowerCase() === V.hole));
    const lo = rs.length ? Math.floor(Math.min(...rs.map((r) => r.depth))) : 0;
    const hi = rs.length ? Math.ceil(Math.max(...rs.map((r) => r.depth))) : 100;
    V.bounds = [lo, hi]; V.depMin = lo; V.depMax = hi;
    ['rngMin', 'rngMax'].forEach((id) => { const e = $('#' + id); e.min = lo; e.max = hi; e.step = Math.max(0.1, (hi - lo) / 500); });
    V.syncDepthUI();
  };
  V.syncDepthUI = function () {
    $('#depMin').value = +V.depMin.toFixed(2); $('#depMax').value = +V.depMax.toFixed(2);
    $('#rngMin').value = V.depMin; $('#rngMax').value = V.depMax;
  };

  V.fit = function (north) {
    V.computeExtent();
    const fb = V.box.clone();
    if ($('#chkAxes').checked && V.axesBox) fb.union(V.axesBox);
    const sz = new THREE.Vector3(); fb.getSize(sz);
    const ext = Math.max(10, sz.x, sz.y, sz.z);
    const t = new THREE.Vector3(); fb.getCenter(t); t.sub(new THREE.Vector3(V.origin[0], V.origin[2], -V.origin[1]));
    const dist = (ext * 0.75) / Math.tan((V.camera.fov * D2R) / 2) * 1.1 + 5;
    V.controls.target.copy(t);
    if (north) V.camera.position.set(t.x, t.y + dist, t.z + dist * 0.0005);
    else V.camera.position.copy(t).add(new THREE.Vector3(-0.6, 0.55, 0.75).normalize().multiplyScalar(dist));
    V.camera.near = Math.max(0.05, dist / 1000); V.camera.far = dist * 50; V.camera.updateProjectionMatrix();
    V.controls.update();
  };

  V.snapshot = function () {
    V.updateBorders();
    V.renderer.render(V.scene, V.camera);
    const src = V.renderer.domElement, c = document.createElement('canvas');
    c.width = src.width; c.height = src.height;
    const g = c.getContext('2d'); g.fillStyle = '#0a1a2e'; g.fillRect(0, 0, c.width, c.height); g.drawImage(src, 0, 0);
    const k = c.width / src.clientWidth;
    g.fillStyle = '#e3eefc'; g.font = `bold ${14 * k}px sans-serif`;
    g.fillText(`Orion – Drillhole & structural planes   |   Hole: ${V.hole === 'all' ? 'All' : $('#selHole').selectedOptions[0].textContent}   |   Depth ${V.depMin.toFixed(1)}–${V.depMax.toFixed(1)} m`, 14 * k, 24 * k);
    if ($('#chkLegend').checked && V.cats.length) {
      const lh = 20 * k, w = 190 * k, h = lh * (V.cats.length + 1) + 12 * k, x = c.width - w - 12 * k, y = c.height - h - 12 * k;
      g.fillStyle = 'rgba(15,31,53,.92)'; g.fillRect(x, y, w, h); g.strokeStyle = '#25405f'; g.strokeRect(x, y, w, h);
      g.fillStyle = '#8fa9c8'; g.font = `bold ${11 * k}px sans-serif`; g.fillText($('#selColor').selectedOptions[0].textContent.toUpperCase(), x + 10 * k, y + 8 * k + lh * 0.6);
      V.cats.forEach((cat, i) => {
        const yy = y + 8 * k + lh * (i + 1), off = V.hidden.has(cat.key);
        g.globalAlpha = off ? 0.4 : 1; g.fillStyle = off ? '#0f1f35' : V.catColor.get(cat.key); g.fillRect(x + 10 * k, yy + 3 * k, 14 * k, 14 * k);
        g.strokeStyle = '#8fa9c8'; g.strokeRect(x + 10 * k, yy + 3 * k, 14 * k, 14 * k);
        g.fillStyle = '#e3eefc'; g.font = `${12 * k}px sans-serif`; g.fillText(cat.name, x + 32 * k, yy + 15 * k); g.globalAlpha = 1;
      });
    }
    return c;
  };

  function download(href, name) { const a = document.createElement('a'); a.href = href; a.download = name; a.click(); }
  function exportView() {
    const fmtSel = $('#selExport').value, c = V.snapshot(), base = 'orion_3d_view';
    if (fmtSel === 'png') download(c.toDataURL('image/png'), base + '.png');
    else if (fmtSel === 'jpeg') download(c.toDataURL('image/jpeg', 0.92), base + '.jpg');
    else {
      const { jsPDF } = window.jspdf;
      const w = 297, h = (297 * c.height) / c.width;
      const pdf = new jsPDF({ orientation: w >= h ? 'landscape' : 'portrait', unit: 'mm', format: [w, Math.max(h, 60)] });
      pdf.addImage(c.toDataURL('image/jpeg', 0.92), 'JPEG', 0, 0, w, h);
      pdf.save(base + '.pdf');
    }
  }

  /* ---------------------------------------------------------------- stereonet (equal-area, lower hemisphere) */
  const SN = { hits: [] };
  const SNs = () => ({
    lines: $('#snLines').checked, poles: $('#snPoles').checked, mean: $('#snMean').checked, meanPole: $('#snMeanPole').checked,
    grid: $('#snGrid').checked, dens: $('#snDens').checked, lw: Number($('#snLw').value), ps: Number($('#snPs').value), white: $('#snWhite').checked,
  });
  const ST = window.OrionStereo;
  const poleVec = ST.poleVec;

  // Mean plane: pole vectors are aligned to one hemisphere, averaged and normalised (see stereo.js).
  SN.mean = function (items) {
    return ST.meanPlane(items.map((it) => ({ dip: it.res.dip, dipDir: it.res.dipDir })));
  };

  SN.render = function (g, w, h, st, o) {
    const dark = !st.white, ink = dark ? '#e3eefc' : '#111', bg = dark ? '#0a1a2e' : '#ffffff';
    g.fillStyle = bg; g.fillRect(0, 0, w, h);
    const R = Math.min(w - 120, h - 90) / 2, cx = w / 2, cy = h / 2 + 8;
    const proj = (v) => { const p = ST.project(v); return [cx + R * p[0], cy - R * p[1]]; };
    g.save();
    g.beginPath(); g.arc(cx, cy, R, 0, 2 * Math.PI); g.fillStyle = dark ? '#0d2038' : '#fff'; g.fill(); g.clip();
    const items = V.items.filter((it) => it.group.visible);
    if (st.dens && items.length) {
      const dn = ST.density(items.map((it) => poleVec(it.res.dip, it.res.dipDir)), 90);
      const oc = document.createElement('canvas'); oc.width = oc.height = dn.size;
      const oc2 = oc.getContext('2d'), img = oc2.createImageData(dn.size, dn.size);
      for (let k = 0; k < dn.grid.length; k++) {
        const v = dn.grid[k] / dn.max;
        if (!(v > 0.1)) continue;
        const c = ST.blueRGB(v);
        img.data[4 * k] = c[0]; img.data[4 * k + 1] = c[1]; img.data[4 * k + 2] = c[2];
        img.data[4 * k + 3] = Math.round(255 * Math.min(1, (dark ? 0.3 : 0.15) + (dark ? 0.55 : 0.7) * v));
      }
      oc2.putImageData(img, 0, 0);
      g.imageSmoothingEnabled = true; g.imageSmoothingQuality = 'high';
      g.drawImage(oc, cx - R - R / (dn.size - 1), cy - R - R / (dn.size - 1), 2 * R + (2 * R) / (dn.size - 1), 2 * R + (2 * R) / (dn.size - 1));
      g.lineWidth = 1;
      [0.25, 0.4, 0.55, 0.7, 0.85].forEach((f, k) => {
        g.strokeStyle = dark ? ST.blueRamp(0.25 + 0.12 * k) : ST.blueRamp(0.5 + 0.1 * k); g.beginPath();
        ST.contourSegments(dn.grid, dn.size, f * dn.max).forEach((s) => {
          g.moveTo(cx + R * (-1 + (2 * s[0]) / (dn.size - 1)), cy + R * ((2 * s[1]) / (dn.size - 1) - 1));
          g.lineTo(cx + R * (-1 + (2 * s[2]) / (dn.size - 1)), cy + R * ((2 * s[3]) / (dn.size - 1) - 1));
        });
        g.stroke();
      });
    }
    if (st.grid) {
      g.strokeStyle = dark ? '#35557c' : '#c8c8c8'; g.lineWidth = 0.8; g.setLineDash([2, 3]);
      for (let p = 10; p < 90; p += 10) { const r = R * Math.SQRT2 * Math.sin(((90 - p) * D2R) / 2); g.beginPath(); g.arc(cx, cy, r, 0, 2 * Math.PI); g.stroke(); }
      for (let a = 0; a < 360; a += 15) { const t = a * D2R; g.beginPath(); g.moveTo(cx, cy); g.lineTo(cx + R * Math.sin(t), cy - R * Math.cos(t)); g.stroke(); }
      g.setLineDash([]);
    }
    g.strokeStyle = ink; g.lineWidth = 0.8;
    g.beginPath(); g.moveTo(cx, cy - R); g.lineTo(cx, cy + R); g.moveTo(cx - R, cy); g.lineTo(cx + R, cy); g.stroke();

    const geo = items.map((it) => {
      const pts = ST.greatCircle(it.res.dip, it.res.dipDir, 90).map(proj);
      return { it, pts, pole: proj(poleVec(it.res.dip, it.res.dipDir)), col: V.catColor.get(it.cat) || '#888' };
    });
    if (o && o.hits) SN.hits = [];
    const path = (pts) => { g.beginPath(); pts.forEach((p, i) => (i ? g.lineTo(p[0], p[1]) : g.moveTo(p[0], p[1]))); };
    if (st.lines) geo.forEach((q) => { g.strokeStyle = q.col; g.lineWidth = st.lw; g.globalAlpha = 0.9; path(q.pts); g.stroke(); });
    g.globalAlpha = 1;
    if (st.poles) geo.forEach((q) => { g.beginPath(); g.arc(q.pole[0], q.pole[1], st.ps, 0, 2 * Math.PI); g.fillStyle = q.col; g.fill(); g.strokeStyle = dark ? '#0a1a2e' : '#222'; g.lineWidth = 1; g.stroke(); });
    const sel = geo.find((q) => q.it.res.idx === state.selected);
    if (sel) {
      if (st.lines) { g.strokeStyle = dark ? '#fff' : '#000'; g.lineWidth = st.lw + 3; path(sel.pts); g.stroke(); g.strokeStyle = sel.col; g.lineWidth = st.lw + 1; path(sel.pts); g.stroke(); }
      g.beginPath(); g.arc(sel.pole[0], sel.pole[1], st.ps + 4, 0, 2 * Math.PI); g.strokeStyle = dark ? '#fff' : '#000'; g.lineWidth = 2; g.stroke();
      if (!st.poles) { g.beginPath(); g.arc(sel.pole[0], sel.pole[1], st.ps, 0, 2 * Math.PI); g.fillStyle = sel.col; g.fill(); }
    }
    const m = SN.mean(items); SN.meanRes = m;
    if (m && st.mean) {
      const pts = ST.greatCircle(m.dip, m.dipDir, 90).map(proj);
      g.setLineDash([9, 6]); g.lineWidth = st.lw + 3.5; g.strokeStyle = '#000'; path(pts); g.stroke();
      g.lineWidth = st.lw + 1.2; g.strokeStyle = '#ff2020'; path(pts); g.stroke(); g.setLineDash([]);
    }
    g.restore();
    if (m && st.meanPole) {
      const p = proj(m.pole), k = st.ps + 3;
      g.lineCap = 'round';
      [['#000', 7], ['#ff2020', 3.5]].forEach(([c, lw]) => { g.strokeStyle = c; g.lineWidth = lw; g.beginPath(); g.moveTo(p[0] - k, p[1] - k); g.lineTo(p[0] + k, p[1] + k); g.moveTo(p[0] + k, p[1] - k); g.lineTo(p[0] - k, p[1] + k); g.stroke(); });
      g.lineCap = 'butt';
    }
    g.strokeStyle = ink; g.lineWidth = 1.5; g.beginPath(); g.arc(cx, cy, R, 0, 2 * Math.PI); g.stroke();
    g.fillStyle = ink; g.font = '12px sans-serif'; g.textAlign = 'center'; g.textBaseline = 'middle';
    for (let a = 0; a < 360; a += 45) { const t = a * D2R; g.fillText(a + '°', cx + (R + 18) * Math.sin(t), cy - (R + 18) * Math.cos(t)); }
    if (o && o.title) {
      g.textAlign = 'left'; g.font = 'bold 14px sans-serif'; g.fillText(o.title, 14, 22);
      g.font = '12px sans-serif'; g.fillStyle = dark ? '#8fa9c8' : '#444';
      g.fillText(`Equal-area · lower hemisphere · ${items.length} structure(s)` + (m ? ` · mean plane ${fmt3(m.dipDir)}° dip dir / ${fmt3(m.dip)}° dip` : ''), 14, 42);
    }
    if (o && o.legend) {
      const lh = 20, lw = 190, hh = lh * (V.cats.length + 1) + 12, x = w - lw - 12, y = 12;
      const cnt = new Map(); items.forEach((it) => cnt.set(it.cat, (cnt.get(it.cat) || 0) + 1));
      g.fillStyle = dark ? 'rgba(15,31,53,.92)' : '#f4f4f4'; g.fillRect(x, y, lw, hh); g.strokeStyle = dark ? '#25405f' : '#999'; g.strokeRect(x, y, lw, hh);
      g.textAlign = 'left'; g.fillStyle = dark ? '#8fa9c8' : '#444'; g.font = 'bold 11px sans-serif'; g.fillText($('#selColor').selectedOptions[0].textContent.toUpperCase(), x + 10, y + 8 + lh * 0.55);
      V.cats.forEach((cat, i) => {
        const yy = y + 8 + lh * (i + 1), off = V.hidden.has(cat.key);
        g.globalAlpha = off ? 0.4 : 1; g.fillStyle = V.catColor.get(cat.key); g.fillRect(x + 10, yy + 3, 14, 14);
        g.fillStyle = ink; g.font = '12px sans-serif'; g.fillText(`${cat.name}: ${cnt.get(cat.key) || 0}`, x + 32, yy + 11); g.globalAlpha = 1;
      });
    }
    if (o && o.hits) {
      SN.hits = geo.map((q) => ({ idx: q.it.res.idx, pts: st.lines ? q.pts : null, pole: st.poles ? q.pole : null }));
    }
  };

  SN.draw = function () {
    const cv = $('#sn');
    if (!cv || !state.results.length) return;
    const w = cv.clientWidth, h = cv.clientHeight;
    if (!w || !h) return;
    const dpr = Math.min(2, window.devicePixelRatio || 1);
    if (cv.width !== Math.round(w * dpr) || cv.height !== Math.round(h * dpr)) { cv.width = Math.round(w * dpr); cv.height = Math.round(h * dpr); }
    const g = cv.getContext('2d'); g.setTransform(dpr, 0, 0, dpr, 0, 0);
    SN.render(g, w, h, SNs(), { hits: true });
    const m = SN.meanRes;
    $('#snInfo').innerHTML = m
      ? `<b>Mean plane</b> of ${m.n} structure${m.n > 1 ? 's' : ''} shown<br>Dip dir. <b>${fmt3(m.dipDir)}°</b> · Dip <b>${fmt3(m.dip)}°</b><br>Strike (RHR) <b>${fmt3(norm360(m.dipDir - 90))}°</b><br>Mean pole: trend <b>${fmt3(m.trend)}°</b> · plunge <b>${fmt3(m.plunge)}°</b>`
      : 'No structures shown.';
  };

  SN.pick = function (x, y) {
    let best = null, bd = 1e9;
    const st = SNs();
    SN.hits.forEach((q) => {
      let d = 1e9;
      if (q.pole) d = Math.min(d, Math.hypot(x - q.pole[0], y - q.pole[1]) - st.ps);
      if (q.pts) for (let i = 1; i < q.pts.length; i++) {
        const [ax, ay] = q.pts[i - 1], [bx, by] = q.pts[i], dx = bx - ax, dy = by - ay, l2 = dx * dx + dy * dy || 1;
        const t = Math.max(0, Math.min(1, ((x - ax) * dx + (y - ay) * dy) / l2));
        d = Math.min(d, Math.hypot(x - ax - t * dx, y - ay - t * dy));
      }
      if (d < bd) { bd = d; best = q; }
    });
    return best && bd <= 6 ? best.idx : -1;
  };

  SN.export = function () {
    const fmtSel = $('#selSnExport').value, c = document.createElement('canvas'), W = 1400, H = 1100, k = 2;
    c.width = W * k; c.height = H * k;
    const g = c.getContext('2d'); g.scale(k, k);
    SN.render(g, W, H, SNs(), { title: 'Orion – Stereonet', legend: $('#chkLegend').checked });
    const base = 'orion_stereonet';
    if (fmtSel === 'png') download(c.toDataURL('image/png'), base + '.png');
    else if (fmtSel === 'jpeg') download(c.toDataURL('image/jpeg', 0.92), base + '.jpg');
    else {
      const { jsPDF } = window.jspdf, w = 297, h = (297 * H) / W;
      const pdf = new jsPDF({ orientation: 'landscape', unit: 'mm', format: [w, h] });
      pdf.addImage(c.toDataURL('image/jpeg', 0.92), 'JPEG', 0, 0, w, h); pdf.save(base + '.pdf');
    }
  };

  /* ---------------------------------------------------------------- wiring */
  function fillHoleSelect() {
    const names = [...state.holes.values()].filter((h) => !h.error && state.results.some((r) => !r.error && r.holeId.toLowerCase() === h.name.toLowerCase())).map((h) => h.name);
    $('#selHole').innerHTML = '<option value="all">All holes</option>' + names.map((n) => `<option value="${esc(n.toLowerCase())}">${esc(n)}</option>`).join('');
    V.hole = 'all';
  }

  function run() {
    const msg = $('#calcMsg');
    try {
      calculate();
    } catch (e) { msg.textContent = 'Calculation failed: ' + e.message; msg.className = 'msg err'; console.error(e); return; }
    const ok = state.results.filter((r) => !r.error).length;
    msg.className = 'msg' + (ok ? '' : ' err');
    msg.textContent = `${ok} of ${state.results.length} structures calculated.`;
    $('#step3').classList.remove('hidden');
    V.field = $('#selColor').value;
    fillHoleSelect();
    V.init();
    V.rebuild();
    V.setDepthBounds();
    V.refresh();
    V.fit(false);
    renderTable();
    select(-1);
    $('#step3').scrollIntoView({ behavior: 'smooth' });
  }

  $('#btnCalc').onclick = run;
  $('#btnReset').onclick = () => { if (confirm('Clear all loaded files and results?')) location.reload(); };
  $('#selPrec').onchange = (e) => { state.prec = Number(e.target.value); renderTable(); select(state.selected); };
  $('#selHole').onchange = (e) => { V.hole = e.target.value; V.computeExtent(); V.rebuild(); V.setDepthBounds(); V.refresh(); V.fit(false); };
  $('#selColor').onchange = (e) => { V.field = e.target.value; V.recolor(); };
  $('#rngOpacity').oninput = (e) => { const op = Number(e.target.value) / 100; $('#opVal').textContent = Math.round((1 - op) * 100) + '% transparent'; V.items.forEach((it) => { it.mesh.material.opacity = op; it.back.material.opacity = op; }); };
  $('#rngSize').oninput = () => { const keep = new Set(V.hidden); V.rebuild(); V.hidden = keep; V.refresh(); };
  const depthIn = (which, val) => {
    val = Number(val); if (Number.isNaN(val)) return;
    if (which === 'min') V.depMin = Math.min(val, V.depMax); else V.depMax = Math.max(val, V.depMin);
    V.syncDepthUI(); V.refresh();
  };
  ['depMin', 'rngMin'].forEach((id) => ($('#' + id).oninput = (e) => depthIn('min', e.target.value)));
  ['depMax', 'rngMax'].forEach((id) => ($('#' + id).oninput = (e) => depthIn('max', e.target.value)));
  $('#chkLegend').onchange = () => V.legend();
  $('#chkLabels').onchange = () => V.refresh();
  $('#chkAxes').onchange = () => { if (V.axesG) V.axesG.visible = $('#chkAxes').checked; };
  $('#vpHead').onclick = () => { $('#vpanel').classList.toggle('collapsed'); $('#vpToggle').textContent = $('#vpanel').classList.contains('collapsed') ? '+' : '–'; };
  $('#btnReset3d').onclick = () => V.fit(false);
  $('#btnNorth').onclick = () => V.fit(true);
  $('#btnFull').onclick = () => { const w = $('#vwrap'); document.fullscreenElement ? document.exitFullscreen() : w.requestFullscreen(); };
  document.addEventListener('fullscreenchange', () => { $('#btnFull').textContent = document.fullscreenElement === $('#vwrap') ? '✕ Exit' : '⛶ Maximise'; });
  $('#btnExport').onclick = exportView;
  $('#opVal').textContent = '50% transparent';

  // stereonet controls
  ['snLines', 'snPoles', 'snDens', 'snMean', 'snMeanPole', 'snGrid', 'snWhite'].forEach((id) => ($('#' + id).onchange = () => SN.draw()));
  ['snLw', 'snPs'].forEach((id) => ($('#' + id).oninput = (e) => { $('#' + id + 'V').textContent = e.target.value; SN.draw(); }));
  $('#sn').addEventListener('click', (e) => {
    const b = e.currentTarget.getBoundingClientRect(), i = SN.pick(e.clientX - b.left, e.clientY - b.top);
    if (i >= 0) select(i, true);
  });
  $('#sn').addEventListener('mousemove', (e) => { const b = e.currentTarget.getBoundingClientRect(); e.currentTarget.style.cursor = SN.pick(e.clientX - b.left, e.clientY - b.top) >= 0 ? 'pointer' : 'default'; });
  new ResizeObserver(() => SN.draw()).observe($('#swrap'));
  $('#btnSnFull').onclick = () => { const w = $('#swrap'); document.fullscreenElement ? document.exitFullscreen() : w.requestFullscreen(); };
  document.addEventListener('fullscreenchange', () => { $('#btnSnFull').textContent = document.fullscreenElement === $('#swrap') ? '✕ Exit' : '⛶ Maximise'; });
  $('#btnSnExport').onclick = SN.export;

  // table → viewers
  const resetViewerScope = () => {
    const changed = V.hole !== 'all';
    $('#selHole').value = 'all'; V.hole = 'all';
    if (changed) { V.computeExtent(); V.rebuild(); }
    V.setDepthBounds();
  };
  $('#btnShowSel').onclick = () => {
    state.only = new Set(state.checked);
    resetViewerScope(); V.hidden = new Set(); V.refresh(); V.fit(false); updateTableUi();
  };
  $('#btnShowAll').onclick = () => { state.only = null; V.hidden = new Set(); V.refresh(); V.fit(false); updateTableUi(); };
  $('#btnUncheck').onclick = () => { state.checked.clear(); renderBody(); };
  $('#btnClearTf').onclick = () => { state.tf = {}; state.sort = { col: -1, dir: 1 }; renderTable(); };

  function tableAoa() {
    const head = ['Row', 'HoleID', 'Depth_mid', 'Depth_top', 'Depth_bottom', 'Type', 'Ori_line', 'Alpha', 'Beta', 'Gamma', 'Core_length_cm', 'Core_thickness_cm', 'Min1', 'Min1_pct', 'Min2', 'Min2_pct', 'Min3', 'Min3_pct', 'Min4', 'Min4_pct', 'Vein_style', 'Comment', 'Hole_azimuth', 'Hole_inclination_down', 'True_dip', 'Dip_direction', 'Strike_RHR', 'X', 'Y', 'Z', 'Status'];
    const n = (v) => (Number.isFinite(v) ? v : '');
    const rows = state.results.map((r) => [r.row, r.holeId, r.error ? '' : +r.depth.toFixed(3), r.error ? '' : +r.top.toFixed(3), r.error ? '' : +r.bottom.toFixed(3), r.type, r.oriLine, r.alpha, r.beta, n(r.gamma), n(r.coreLength), n(r.coreThick),
      ...r.min.flatMap((m) => [m.name, n(m.pct)]), r.style, r.comment,
      r.error ? '' : +r.holeAz.toFixed(2), r.error ? '' : +r.holeInc.toFixed(2), r.error ? '' : pad(r.dip), r.error ? '' : pad(r.dipDir), r.error ? '' : pad(r.strike),
      ...(r.error ? ['', '', ''] : r.pos.map((p) => +p.toFixed(3))), r.error || 'OK']);
    return [head, ...rows];
  }
  const saveSheet = (type) => { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(tableAoa()), 'Structures'); XLSX.writeFile(wb, 'orion_structures_calculated.' + type, { bookType: type }); };
  $('#btnCsv').onclick = () => saveSheet('csv');
  $('#btnXlsx').onclick = () => saveSheet('xlsx');

  drawCompass(null);
  window.__orion = { state, calculate, FIELDS, V };
})();
