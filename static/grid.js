/* Spreadsheet-style table editor: multi-cell selection, move (drag the selection edge, Alt+arrows,
   cut/paste), merge/unmerge, insert/delete rows and columns, clipboard (TSV, works with Excel),
   undo/redo and a right-click menu.

   Data model: rows = string[][] (always rectangular), merges = [{r, c, rs, cs}] where (r, c) is the
   top-left cell and rs/cs the row/column span. Cells covered by a merge keep "" and aren't rendered. */
function createGrid({ box, onChange, status, hasHeader, colNames = null, cellClass = null }) {
  // colNames: fixed column headings shown instead of A, B, C… — the column set is then locked
  // (no inserting, deleting or reordering columns; pastes are trimmed to fit).
  const fixedCols = !!colNames;
  let rows = [], merges = [];
  let anchor = [0, 0], focus = [0, 0];       // selection = rectangle spanned by anchor..focus (+ merges)
  let editing = null;                         // {r, c, td, before}
  let undo = [], redo = [];

  const table = document.createElement('table');
  table.className = 'sheet';
  const selRect = document.createElement('div');
  selRect.className = 'selrect';
  selRect.innerHTML = '<div class="edge t"></div><div class="edge r"></div><div class="edge b"></div><div class="edge l"></div>';
  const ghost = document.createElement('div');
  ghost.className = 'ghost';
  const menu = document.createElement('div');
  menu.className = 'ctxmenu'; menu.hidden = true;
  box.append(table, selRect, ghost);
  document.body.append(menu);
  box.tabIndex = 0;

  /* ---------- helpers ---------- */
  const nR = () => rows.length, nC = () => (rows[0] || []).length;
  const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));
  const colName = c => { let s = ''; c++; while (c) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = (c - m - 1) / 26; } return s; };
  const mergeAt = (r, c) => merges.find(m => r >= m.r && r < m.r + m.rs && c >= m.c && c < m.c + m.cs);
  const covered = (r, c) => { const m = mergeAt(r, c); return m && (m.r !== r || m.c !== c); };
  const blank = (h, w) => Array.from({ length: h }, () => Array(w).fill(''));
  const intersects = (m, x) => m.r <= x.r1 && m.r + m.rs - 1 >= x.r0 && m.c <= x.c1 && m.c + m.cs - 1 >= x.c0;
  const inside = (m, x) => m.r >= x.r0 && m.r + m.rs - 1 <= x.r1 && m.c >= x.c0 && m.c + m.cs - 1 <= x.c1;

  function expand(x) {  // grow a rectangle until it contains every merge it touches
    let grew = true;
    while (grew) {
      grew = false;
      for (const m of merges) {
        if (!intersects(m, x) || inside(m, x)) continue;
        x = { r0: Math.min(x.r0, m.r), c0: Math.min(x.c0, m.c), r1: Math.max(x.r1, m.r + m.rs - 1), c1: Math.max(x.c1, m.c + m.cs - 1) };
        grew = true;
      }
    }
    return x;
  }
  function sel() {
    return expand({ r0: Math.min(anchor[0], focus[0]), c0: Math.min(anchor[1], focus[1]),
                    r1: Math.max(anchor[0], focus[0]), c1: Math.max(anchor[1], focus[1]) });
  }
  function setSel(a, f = a) {
    if (!nR()) return;
    anchor = [clamp(a[0], 0, nR() - 1), clamp(a[1], 0, nC() - 1)];
    focus = [clamp(f[0], 0, nR() - 1), clamp(f[1], 0, nC() - 1)];
    paint();
  }
  function snapshot() { return JSON.stringify({ rows, merges }); }
  function record() { undo.push(snapshot()); if (undo.length > 200) undo.shift(); redo = []; }
  function restore(s) { ({ rows, merges } = JSON.parse(s)); setSel(anchor, focus); render(); changed(); }
  function changed() { onChange && onChange(); }
  function ensure(h, w) {  // grow the table to at least h x w
    if (fixedCols) w = colNames.length;
    while (nR() < h) rows.push(Array(Math.max(nC(), w)).fill(''));
    if (nC() < w) rows.forEach(row => { while (row.length < w) row.push(''); });
  }
  function dropOneByOne() { merges = merges.filter(m => m.rs > 1 || m.cs > 1); }

  /* ---------- rendering ---------- */
  function render() {
    table.innerHTML = '';
    if (!nR()) { selRect.style.display = 'none'; return; }
    const head = table.createTHead().insertRow();
    const corner = document.createElement('th'); corner.className = 'corner'; corner.title = 'Select all'; head.append(corner);
    for (let c = 0; c < nC(); c++) {
      const th = document.createElement('th'); th.className = 'ch' + (fixedCols ? ' named' : ''); th.dataset.c = c;
      th.textContent = fixedCols ? colNames[c] : colName(c); head.append(th);
    }
    const body = table.createTBody();
    rows.forEach((row, r) => {
      const tr = body.insertRow();
      if (hasHeader() && r === 0) tr.className = 'head';
      const th = document.createElement('th'); th.className = 'rh'; th.dataset.r = r; th.textContent = r + 1; tr.append(th);
      row.forEach((val, c) => {
        if (covered(r, c)) return;
        const td = tr.insertCell();
        const m = mergeAt(r, c);
        if (m) { td.rowSpan = m.rs; td.colSpan = m.cs; td.classList.add('merged'); }
        td.dataset.r = r; td.dataset.c = c; td.textContent = val;
        if (cellClass) { const cls = cellClass(r, c, rows[r]); if (cls) td.classList.add(...cls.split(' ')); }
        if (/^[\s؀-ۿݐ-ݿ]*[؀-ۿ]/.test(val)) td.dir = 'rtl';
      });
    });
    paint();
  }
  const cellEl = (r, c) => { const m = mergeAt(r, c); if (m) [r, c] = [m.r, m.c]; return table.querySelector(`td[data-r="${r}"][data-c="${c}"]`); };
  function place(el, x) {
    const a = cellEl(x.r0, x.c0), b = cellEl(x.r1, x.c1);
    if (!a || !b) { el.style.display = 'none'; return; }
    const o = box.getBoundingClientRect(), ra = a.getBoundingClientRect(), rb = b.getBoundingClientRect();
    const left = Math.min(ra.left, rb.left), top = Math.min(ra.top, rb.top);
    Object.assign(el.style, { display: 'block', left: left - o.left + box.scrollLeft + 'px', top: top - o.top + box.scrollTop + 'px',
                              width: Math.max(ra.right, rb.right) - left + 'px', height: Math.max(ra.bottom, rb.bottom) - top + 'px' });
  }
  function paint() {
    if (!nR()) return;
    const x = sel();
    table.querySelectorAll('.sel, .active, .hl').forEach(el => el.classList.remove('sel', 'active', 'hl'));
    table.querySelectorAll('td[data-r]').forEach(td => {
      const r = +td.dataset.r, c = +td.dataset.c;
      if (r >= x.r0 && r <= x.r1 && c >= x.c0 && c <= x.c1) td.classList.add('sel');
    });
    for (let c = x.c0; c <= x.c1; c++) table.querySelector(`th.ch[data-c="${c}"]`)?.classList.add('hl');
    for (let r = x.r0; r <= x.r1; r++) table.querySelector(`th.rh[data-r="${r}"]`)?.classList.add('hl');
    const act = cellEl(anchor[0], anchor[1]);
    if (act) { act.classList.add('active'); act.scrollIntoView({ block: 'nearest', inline: 'nearest' }); }
    place(selRect, x);
    selRect.classList.toggle('editing', !!editing);
  }

  /* ---------- editing ---------- */
  function startEdit(text) {
    if (!nR()) return;
    let [r, c] = anchor; const m = mergeAt(r, c); if (m) [r, c] = [m.r, m.c];
    const td = cellEl(r, c); if (!td) return;
    editing = { r, c, td, before: rows[r][c] };
    td.contentEditable = 'plaintext-only';
    if (td.contentEditable !== 'plaintext-only') td.contentEditable = 'true';
    td.classList.add('editing');
    if (text !== undefined) td.textContent = text;
    td.focus();
    td.addEventListener('blur', () => { if (editing && editing.td === td) stopEdit(true, false); }, { once: true });
    const s = getSelection(); s.selectAllChildren(td);
    if (text !== undefined) s.collapseToEnd();
    paint();
  }
  function stopEdit(commit, refocus = true) {
    if (!editing) return;
    const { r, c, td, before } = editing;
    const val = td.innerText.replace(/\n$/, '');
    editing = null;
    if (commit && val !== before) { record(); rows[r][c] = val; changed(); }
    render();
    if (refocus) box.focus({ preventScroll: true });  // not when focus left for another field
  }

  /* ---------- operations ---------- */
  function clearSel() {
    const x = sel(); record();
    for (let r = x.r0; r <= x.r1; r++) for (let c = x.c0; c <= x.c1; c++) rows[r][c] = '';
    render(); changed();
  }
  function merge() {
    const x = sel();
    if (x.r0 === x.r1 && x.c0 === x.c1) return status('Select two or more cells to merge.');
    record();
    const parts = [];
    for (let r = x.r0; r <= x.r1; r++) for (let c = x.c0; c <= x.c1; c++) {
      if (!covered(r, c) && rows[r][c].trim()) parts.push(rows[r][c].trim());
      rows[r][c] = '';
    }
    rows[x.r0][x.c0] = parts.join(' ');
    merges = merges.filter(m => !inside(m, x));
    merges.push({ r: x.r0, c: x.c0, rs: x.r1 - x.r0 + 1, cs: x.c1 - x.c0 + 1 });
    setSel([x.r0, x.c0]); render(); changed();
  }
  function unmerge() {
    const x = sel(), before = merges.length;
    if (!merges.some(m => intersects(m, x))) return status('No merged cells in the selection.');
    record(); merges = merges.filter(m => !intersects(m, x));
    render(); changed(); status(`Unmerged ${before - merges.length} cell group(s).`);
  }
  // Insert/delete along one axis; `axis` 'r' or 'c'.
  function insertAt(axis, at, n) {
    if (axis === 'c' && fixedCols) return status('This table has fixed columns.');
    record();
    if (axis === 'r') rows.splice(at, 0, ...blank(n, nC()));
    else rows.forEach(row => row.splice(at, 0, ...Array(n).fill('')));
    const [s, l] = axis === 'r' ? ['r', 'rs'] : ['c', 'cs'];
    merges.forEach(m => { if (m[s] >= at) m[s] += n; else if (m[s] + m[l] > at) m[l] += n; });
    render(); changed();
  }
  function deleteRange(axis, a, b) {
    if (axis === 'c' && fixedCols) return status('This table has fixed columns.');
    const n = b - a + 1, total = axis === 'r' ? nR() : nC();
    if (n >= total) return status(`Can't delete every ${axis === 'r' ? 'row' : 'column'} — use Clear instead.`);
    record();
    if (axis === 'r') rows.splice(a, n); else rows.forEach(row => row.splice(a, n));
    const [s, l] = axis === 'r' ? ['r', 'rs'] : ['c', 'cs'];
    merges = merges.filter(m => {
      const e = m[s] + m[l] - 1;
      if (e < a) return true;
      if (m[s] > b) { m[s] -= n; return true; }
      const overlap = Math.min(e, b) - Math.max(m[s], a) + 1;
      m[l] -= overlap; m[s] = Math.min(m[s], a);
      return m[l] > 0;
    });
    dropOneByOne();
    if (axis === 'r') setSel([Math.min(a, nR() - 1), anchor[1]]); else setSel([anchor[0], Math.min(a, nC() - 1)]);
    render(); changed();
  }
  // Move the selected block by (dr, dc), overwriting the destination (like Excel's drag-move).
  function moveBlock(dr, dc) {
    const x = sel();
    dr = Math.max(dr, -x.r0); dc = Math.max(dc, -x.c0);
    if (fixedCols) dc = Math.min(dc, nC() - 1 - x.c1);
    if (!dr && !dc) return;
    const h = x.r1 - x.r0 + 1, w = x.c1 - x.c0 + 1;
    const dst = { r0: x.r0 + dr, c0: x.c0 + dc, r1: x.r1 + dr, c1: x.c1 + dc };
    record();
    ensure(dst.r1 + 1, dst.c1 + 1);
    const block = rows.slice(x.r0, x.r1 + 1).map(row => row.slice(x.c0, x.c1 + 1));
    const moving = merges.filter(m => inside(m, x));
    merges = merges.filter(m => !inside(m, x) && !intersects(m, dst));
    for (let r = x.r0; r <= x.r1; r++) for (let c = x.c0; c <= x.c1; c++) rows[r][c] = '';
    for (let i = 0; i < h; i++) for (let j = 0; j < w; j++) rows[dst.r0 + i][dst.c0 + j] = block[i][j];
    moving.forEach(m => { m.r += dr; m.c += dc; merges.push(m); });
    anchor = [anchor[0] + dr, anchor[1] + dc]; focus = [focus[0] + dr, focus[1] + dc];
    render(); changed();
  }
  // Reorder whole rows/columns: shift the selected ones one step (dir = -1 / +1).
  function shiftLines(axis, dir) {
    if (axis === 'c' && fixedCols) return status('This table has fixed columns.');
    const x = sel();
    const [a, b] = axis === 'r' ? [x.r0, x.r1] : [x.c0, x.c1];
    const total = axis === 'r' ? nR() : nC();
    if ((dir < 0 && a === 0) || (dir > 0 && b === total - 1)) return;
    const order = [...Array(total).keys()];  // order[newIndex] = oldIndex
    if (dir < 0) order.splice(b, 0, ...order.splice(a - 1, 1)); else order.splice(a, 0, ...order.splice(b + 1, 1));
    const where = []; order.forEach((old, i) => { where[old] = i; });
    const [s, l] = axis === 'r' ? ['r', 'rs'] : ['c', 'cs'];
    for (const m of merges) for (let k = 1; k < m[l]; k++)
      if (where[m[s] + k] !== where[m[s]] + k) return status('A merged cell spans that boundary — unmerge it first.');
    record();
    if (axis === 'r') rows = order.map(i => rows[i]); else rows = rows.map(row => order.map(i => row[i]));
    merges.forEach(m => { m[s] = where[m[s]]; });
    if (axis === 'r') { anchor[0] += dir; focus[0] += dir; } else { anchor[1] += dir; focus[1] += dir; }
    render(); changed();
  }

  /* ---------- clipboard (TSV, compatible with Excel / Google Sheets) ---------- */
  function toTSV(x) {
    const q = v => /[\t\n"]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
    const out = [];
    for (let r = x.r0; r <= x.r1; r++) out.push(rows[r].slice(x.c0, x.c1 + 1).map(q).join('\t'));
    return out.join('\n');
  }
  function parseTSV(text) {
    const out = [[]]; let cur = '', quoted = false;
    text = text.replace(/\r\n?/g, '\n').replace(/\n$/, '');
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (quoted) {
        if (ch === '"' && text[i + 1] === '"') { cur += '"'; i++; }
        else if (ch === '"') quoted = false;
        else cur += ch;
      } else if (ch === '"' && cur === '') quoted = true;
      else if (ch === '\t') { out[out.length - 1].push(cur); cur = ''; }
      else if (ch === '\n') { out[out.length - 1].push(cur); out.push([]); cur = ''; }
      else cur += ch;
    }
    out[out.length - 1].push(cur);
    return out;
  }
  function paste(text) {
    const data = parseTSV(text);
    const h = data.length, x = sel(), r0 = x.r0, c0 = x.c0;
    let w = Math.max(...data.map(r => r.length));
    if (fixedCols) w = Math.min(w, nC() - c0);  // trim what doesn't fit
    record();
    ensure(r0 + h, c0 + w);
    const dst = { r0, c0, r1: r0 + h - 1, c1: c0 + w - 1 };
    merges = merges.filter(m => !intersects(m, dst));
    data.forEach((row, i) => { for (let j = 0; j < w; j++) rows[r0 + i][c0 + j] = row[j] ?? ''; });
    anchor = [r0, c0]; focus = [dst.r1, dst.c1];
    render(); changed();
  }
  let clipCut = null;  // remembered so a cut+paste inside the grid becomes a move (keeps merges)
  function copy(cut) {
    const x = sel(), text = toTSV(x);
    clipCut = cut ? { x, text } : null;
    return text;
  }
  function pasteFromGrid(text) {
    if (clipCut && clipCut.text === text) {  // internal cut -> move the block to the selection
      const x = clipCut.x, to = sel(); clipCut = null;
      anchor = [x.r0, x.c0]; focus = [x.r1, x.c1];
      return moveBlock(to.r0 - x.r0, to.c0 - x.c0);
    }
    paste(text);
  }

  /* ---------- mouse ---------- */
  let dragSel = null, dragMove = null;
  const hitCell = (e) => {
    for (const el of document.elementsFromPoint(e.clientX, e.clientY)) {
      if (el.matches?.('td[data-r]')) return [+el.dataset.r, +el.dataset.c];
      if (el.matches?.('th.ch')) return [-1, +el.dataset.c, 'col'];
      if (el.matches?.('th.rh')) return [+el.dataset.r, -1, 'row'];
    }
    return null;
  };
  box.addEventListener('pointerdown', e => {
    hideMenu();
    if (e.button !== 0 && e.button !== 2) return;
    if (e.target.closest('.edge')) {  // grab the selection border -> move
      const at = hitCell(e);
      if (at && at[2] === undefined) { dragMove = { start: at, x: sel() }; e.preventDefault(); }
      return;
    }
    const t = e.target;
    if (editing && t.closest('td') === editing.td) return;  // clicking inside the cell being edited
    if (editing) stopEdit(true);
    if (e.button === 2) {  // right-click inside the selection keeps it
      const at = hitCell(e), x = sel();
      if (at && at[0] >= x.r0 && at[0] <= x.r1 && at[1] >= x.c0 && at[1] <= x.c1) return;
    }
    if (t.closest('th.corner')) { setSel([0, 0], [nR() - 1, nC() - 1]); box.focus(); return; }
    const th = t.closest('th.ch, th.rh'), td = t.closest('td[data-r]');
    if (!th && !td) return;
    e.preventDefault(); box.focus({ preventScroll: true });
    if (th && th.classList.contains('ch')) {
      const c = +th.dataset.c;
      if (e.shiftKey) setSel([0, anchor[1]], [nR() - 1, c]); else setSel([0, c], [nR() - 1, c]);
      dragSel = 'col';
    } else if (th) {
      const r = +th.dataset.r;
      if (e.shiftKey) setSel([anchor[0], 0], [r, nC() - 1]); else setSel([r, 0], [r, nC() - 1]);
      dragSel = 'row';
    } else {
      const at = [+td.dataset.r, +td.dataset.c];
      if (e.shiftKey) setSel(anchor, at); else setSel(at);
      dragSel = 'cell';
    }
  });
  document.addEventListener('pointermove', e => {
    if (dragSel) {
      const at = hitCell(e); if (!at) return;
      if (dragSel === 'col' && at[1] >= 0) setSel([0, anchor[1]], [nR() - 1, at[1]]);
      else if (dragSel === 'row' && at[0] >= 0) setSel([anchor[0], 0], [at[0], nC() - 1]);
      else if (dragSel === 'cell' && at[0] >= 0 && at[1] >= 0) setSel(anchor, at);
    } else if (dragMove) {
      const at = hitCell(e); if (!at || at[0] < 0 || at[1] < 0) return;
      const { x, start } = dragMove;
      const dr = Math.max(at[0] - start[0], -x.r0), dc = Math.max(at[1] - start[1], -x.c0);
      dragMove.d = [dr, dc];
      const dst = { r0: x.r0 + dr, c0: x.c0 + dc, r1: Math.min(x.r1 + dr, nR() - 1), c1: Math.min(x.c1 + dc, nC() - 1) };
      place(ghost, dst);
    }
  });
  document.addEventListener('pointerup', () => {
    dragSel = null;
    if (dragMove) {
      ghost.style.display = 'none';
      const d = dragMove.d; dragMove = null;
      if (d) moveBlock(d[0], d[1]);
    }
  });
  box.addEventListener('dblclick', e => { if (e.target.closest('td[data-r]') && !editing) startEdit(); });

  /* ---------- context menu ---------- */
  const ITEMS = [
    ['Cut', 'Ctrl+X', () => clipAction('cut')], ['Copy', 'Ctrl+C', () => clipAction('copy')], ['Paste', 'Ctrl+V', () => clipAction('paste')], '-',
    ['Merge cells', 'Ctrl+M', merge], ['Unmerge cells', '', unmerge], '-',
    ['Insert row above', '', () => { const x = sel(); insertAt('r', x.r0, x.r1 - x.r0 + 1); }],
    ['Insert row below', '', () => { const x = sel(); insertAt('r', x.r1 + 1, x.r1 - x.r0 + 1); }],
    ['Insert column left', '', () => { const x = sel(); insertAt('c', x.c0, x.c1 - x.c0 + 1); }],
    ['Insert column right', '', () => { const x = sel(); insertAt('c', x.c1 + 1, x.c1 - x.c0 + 1); }],
    ['Delete rows', '', () => { const x = sel(); deleteRange('r', x.r0, x.r1); }],
    ['Delete columns', '', () => { const x = sel(); deleteRange('c', x.c0, x.c1); }], '-',
    ['Move rows up', 'Alt+↑', () => shiftLines('r', -1)], ['Move rows down', 'Alt+↓', () => shiftLines('r', 1)],
    ['Move columns left', 'Alt+←', () => shiftLines('c', -1)], ['Move columns right', 'Alt+→', () => shiftLines('c', 1)], '-',
    ['Clear contents', 'Del', clearSel], ['Undo', 'Ctrl+Z', () => api.undo()], ['Redo', 'Ctrl+Y', () => api.redo()],
  ];
  box.addEventListener('contextmenu', e => {
    if (!e.target.closest('td, th') || editing) return;
    e.preventDefault();
    menu.innerHTML = '';
    const items = fixedCols ? ITEMS.filter(it => it === '-' || !/column/i.test(it[0])) : ITEMS;
    for (const it of items) {
      if (it === '-') { menu.append(Object.assign(document.createElement('hr'))); continue; }
      const b = document.createElement('button');
      b.innerHTML = `<span>${it[0]}</span><kbd>${it[1]}</kbd>`;
      b.onclick = () => { hideMenu(); it[2](); box.focus({ preventScroll: true }); };
      menu.append(b);
    }
    menu.hidden = false;
    const w = menu.offsetWidth, h = menu.offsetHeight;
    menu.style.left = Math.min(e.clientX, innerWidth - w - 8) + 'px';
    menu.style.top = Math.min(e.clientY, innerHeight - h - 8) + 'px';
  });
  function hideMenu() { menu.hidden = true; }
  document.addEventListener('pointerdown', e => { if (!menu.contains(e.target) && !box.contains(e.target)) hideMenu(); });
  window.addEventListener('blur', hideMenu);
  async function clipAction(kind) {
    try {
      if (kind === 'paste') pasteFromGrid(await navigator.clipboard.readText());
      else { await navigator.clipboard.writeText(copy(kind === 'cut')); if (kind === 'cut') status('Cut — select the destination and paste to move.'); }
    } catch { status(`Use ${kind === 'paste' ? 'Ctrl+V' : kind === 'cut' ? 'Ctrl+X' : 'Ctrl+C'} — the browser blocked clipboard access from the menu.`, true); }
  }
  document.addEventListener('copy', e => { if (document.activeElement !== box || editing || !nR()) return; e.preventDefault(); e.clipboardData.setData('text/plain', copy(false)); });
  document.addEventListener('cut', e => { if (document.activeElement !== box || editing || !nR()) return; e.preventDefault(); e.clipboardData.setData('text/plain', copy(true)); status('Cut — select the destination and paste to move.'); });
  document.addEventListener('paste', e => {
    if (editing) { e.preventDefault(); document.execCommand('insertText', false, e.clipboardData.getData('text/plain')); return; }
    if (document.activeElement !== box || !nR()) return;
    e.preventDefault(); pasteFromGrid(e.clipboardData.getData('text/plain'));
  });

  /* ---------- keyboard ---------- */
  function step(dr, dc, extend) {
    // From a merged cell, step off its far edge rather than into its own covered cells.
    const base = extend ? focus : anchor;
    const m = mergeAt(base[0], base[1]);
    let r = base[0], c = base[1];
    if (m && !extend) { if (dr > 0) r = m.r + m.rs - 1; if (dc > 0) c = m.c + m.cs - 1; if (dr < 0) r = m.r; if (dc < 0) c = m.c; }
    const to = [clamp(r + dr, 0, nR() - 1), clamp(c + dc, 0, nC() - 1)];
    if (extend) setSel(anchor, to); else setSel(to);
  }
  box.addEventListener('keydown', e => {
    if (!nR()) return;
    const k = e.key, ctrl = e.ctrlKey || e.metaKey;
    if (editing) {
      if (k === 'Escape') { e.preventDefault(); stopEdit(false); }
      else if (k === 'Enter' && !e.altKey && !e.shiftKey) { e.preventDefault(); stopEdit(true); step(1, 0); }
      else if (k === 'Tab') { e.preventDefault(); stopEdit(true); step(0, e.shiftKey ? -1 : 1); }
      else if (k === 'Enter' && e.altKey) { e.preventDefault(); document.execCommand('insertLineBreak'); }
      return;
    }
    const arrows = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (arrows[k]) {
      e.preventDefault();
      const [dr, dc] = arrows[k];
      if (e.altKey) return shiftLines(dr ? 'r' : 'c', dr || dc);
      if (ctrl) { const to = [dr ? (dr < 0 ? 0 : nR() - 1) : (e.shiftKey ? focus : anchor)[0], dc ? (dc < 0 ? 0 : nC() - 1) : (e.shiftKey ? focus : anchor)[1]]; return e.shiftKey ? setSel(anchor, to) : setSel(to); }
      return step(dr, dc, e.shiftKey);
    }
    if (k === 'Tab') { e.preventDefault(); return step(0, e.shiftKey ? -1 : 1); }
    if (k === 'Enter' || k === 'F2') { e.preventDefault(); return startEdit(); }
    if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); return clearSel(); }
    if (k === 'Escape') { clipCut = null; return hideMenu(); }
    if (ctrl && k.toLowerCase() === 'z') { e.preventDefault(); return e.shiftKey ? api.redo() : api.undo(); }
    if (ctrl && k.toLowerCase() === 'y') { e.preventDefault(); return api.redo(); }
    if (ctrl && k.toLowerCase() === 'a') { e.preventDefault(); return setSel([0, 0], [nR() - 1, nC() - 1]); }
    if (ctrl && k.toLowerCase() === 'm') {  // toggle: a selection that is exactly one merged cell gets unmerged
      e.preventDefault();
      const x = sel(), m = mergeAt(x.r0, x.c0);
      return m && m.r === x.r0 && m.c === x.c0 && m.r + m.rs - 1 === x.r1 && m.c + m.cs - 1 === x.c1 ? unmerge() : merge();
    }
    if (k.length === 1 && !ctrl && !e.altKey) { e.preventDefault(); startEdit(k); }  // typing replaces the cell
  });
  window.addEventListener('resize', () => nR() && paint());

  /* ---------- public API ---------- */
  const api = {
    get rows() { return rows; },
    get merges() { return merges; },
    hasData: () => nR() > 0,
    canUndo: () => undo.length > 0,
    canRedo: () => redo.length > 0,
    setData(newRows, newMerges = []) {
      if (nR()) record();
      const w = fixedCols ? colNames.length : Math.max(1, ...newRows.map(r => r.length));
      rows = newRows.map(r => r.concat(Array(w - r.length).fill('')));
      merges = newMerges; anchor = [0, 0]; focus = [0, 0];
      render(); changed();
    },
    appendRows(more) {
      record();
      const w = fixedCols ? colNames.length : Math.max(nC(), ...more.map(r => r.length));
      ensure(nR(), w);
      more.forEach(r => rows.push(r.concat(Array(w - r.length).fill(''))));
      render(); changed();
    },
    ensureTable() { if (!nR()) { rows = blank(fixedCols ? 1 : 3, fixedCols ? colNames.length : 3); render(); changed(); } },
    // Keep a running number in column `c` (e.g. S.No.). Not an undoable edit; re-renders only if needed.
    renumber(c, skip = null) {  // skip(row): rows left unnumbered (and not counted), e.g. a totals row
      let dirty = false, n = 0;
      rows.forEach(row => {
        const v = skip && skip(row) ? '' : String(++n);
        if (row[c] !== v) { row[c] = v; dirty = true; }
      });
      if (dirty) render();
    },
    clear() { if (nR()) record(); rows = []; merges = []; render(); changed(); },
    render,
    undo() { if (!undo.length) return; redo.push(snapshot()); restore(undo.pop()); },
    redo() { if (!redo.length) return; undo.push(snapshot()); restore(redo.pop()); },
    merge, unmerge, clearSel, moveBlock, shiftLines, paste, copyText: () => toTSV({ r0: 0, c0: 0, r1: nR() - 1, c1: nC() - 1 }),
    insertRowsBelow() { api.ensureTable(); const x = sel(); insertAt('r', x.r1 + 1, x.r1 - x.r0 + 1); },
    insertColsRight() { api.ensureTable(); const x = sel(); insertAt('c', x.c1 + 1, x.c1 - x.c0 + 1); },
    deleteRows() { const x = sel(); deleteRange('r', x.r0, x.r1); },
    deleteCols() { const x = sel(); deleteRange('c', x.c0, x.c1); },
    select: setSel,
    get selection() { return sel(); },
  };
  return api;
}
