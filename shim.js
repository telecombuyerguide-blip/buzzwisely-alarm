// BuzzWisely alarm for Excel: a small stand-in for the Google services the alarm code uses (SpreadsheetApp,
// Utilities, google.script.run), so core.js (generated from alarm-script.gs) runs unchanged inside Excel.
// How it works: before each call, the lead tabs are read from the workbook into memory (values, display text);
// the alarm code reads and writes that copy; the writes are then sent to Excel in one batch.
(function () {
  'use strict';
  const ET = 'America/New_York';
  const LEAD_TABS = /DAY [123]|DEMO|RESCHED/i;

  // ---------- dates and Utilities.formatDate ----------
  const fmtCache = {};
  function parts(d, tz) {
    let f = fmtCache[tz];
    if (!f) f = fmtCache[tz] = new Intl.DateTimeFormat('en-US', {timeZone: tz, year: 'numeric', month: 'numeric', day: 'numeric',
      hour: 'numeric', minute: 'numeric', second: 'numeric', weekday: 'short', hourCycle: 'h23'});
    const o = {};
    f.formatToParts(d).forEach(p => { o[p.type] = p.value; });
    if (o.hour === '24') o.hour = '0';
    return o;
  }
  function offMin(d, tz) {
    const p = parts(d, tz);
    const asUtc = Date.UTC(+p.year, +p.month - 1, +p.day, +p.hour, +p.minute, +p.second);
    return Math.round((asUtc - (d.getTime() - d.getMilliseconds())) / 60000);
  }
  const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const WD = {Mon: 1, Tue: 2, Wed: 3, Thu: 4, Fri: 5, Sat: 6, Sun: 7};
  const pad = (n, k) => String(n).padStart(k, '0');
  function formatDate(d, tz, pat) {
    d = d instanceof Date ? d : new Date(d);
    const p = parts(d, tz || ET), H = +p.hour, h12 = H % 12 || 12;
    return String(pat).replace(/yyyy|MMM|MM|M|dd|d|EEE|HH|H|hh|h|mm|m|ss|s|a|u|Z|'[^']*'/g, t => {
      switch (t) {
        case 'yyyy': return p.year;
        case 'MMM': return MON[+p.month - 1];
        case 'MM': return pad(p.month, 2);
        case 'M': return String(+p.month);
        case 'dd': return pad(p.day, 2);
        case 'd': return String(+p.day);
        case 'EEE': return p.weekday;
        case 'HH': return pad(H, 2);
        case 'H': return String(H);
        case 'hh': return pad(h12, 2);
        case 'h': return String(h12);
        case 'mm': return pad(p.minute, 2);
        case 'm': return String(+p.minute);
        case 'ss': return pad(p.second, 2);
        case 's': return String(+p.second);
        case 'a': return H < 12 ? 'AM' : 'PM';
        case 'u': return String(WD[p.weekday]);
        case 'Z': { const o = offMin(d, tz || ET), a = Math.abs(o); return (o < 0 ? '-' : '+') + pad(Math.floor(a / 60), 2) + pad(a % 60, 2); }
        default: return t.slice(1, -1);
      }
    });
  }
  // Excel keeps dates as serial numbers (days since 12/30/1899, wall time). The board's times are Eastern.
  function serialToDate(n) {
    const wall = Math.round((n - 25569) * 86400000);
    const off = offMin(new Date(wall), ET);
    return new Date(wall - off * 60000);
  }
  function isDateFormat(f) {
    f = String(f || '').toLowerCase().replace(/"[^"]*"|\[[^\]]*\]|\\./g, '');
    if (!f || f === 'general' || f === '@') return false;
    return /[dyh]/.test(f) || (/m/.test(f) && /[:\/]/.test(f));
  }

  // ---------- the in-memory copy of the lead tabs ----------
  let BOOK = null;     // {sheets: [Sheet]}
  let OPS = [];        // queued writes

  function chainProxy(result) {
    const p = new Proxy(function () {}, {get: (t, k) => k === 'build' ? () => result : () => p});
    return p;
  }

  class Sheet {
    constructor(name, values, text) { this.name = name; this.v = values; this.t = text; }
    getName() { return this.name; }
    getLastRow() { return this.v.length; }
    getLastColumn() { return this.v.reduce((m, r) => Math.max(m, r.length), 0); }
    getRange(a, b, c, d) {
      if (typeof a === 'string') { const m = a.toUpperCase().match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/); if (!m) throw new Error('Bad range ' + a);
        const col = s => s.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
        const r1 = +m[2], c1 = col(m[1]), r2 = m[4] ? +m[4] : r1, c2 = m[3] ? col(m[3]) : c1;
        return new Range(this, r1, c1, r2 - r1 + 1, c2 - c1 + 1); }
      return new Range(this, a, b, c || 1, d || 1);
    }
    getDataRange() { return new Range(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
    cell(r, c) { const row = this.v[r - 1]; return row && c - 1 < row.length ? row[c - 1] : ''; }
    txt(r, c) { const row = this.t[r - 1]; return row && c - 1 < row.length ? row[c - 1] : ''; }
    put(r, c, val, shown) {
      while (this.v.length < r) { this.v.push([]); this.t.push([]); }
      while (this.v[r - 1].length < c) { this.v[r - 1].push(''); this.t[r - 1].push(''); }
      this.v[r - 1][c - 1] = val; this.t[r - 1][c - 1] = shown;
    }
  }

  class Range {
    constructor(sh, r, c, nr, nc) { this.sh = sh; this.r = r; this.c = c; this.nr = nr; this.nc = nc; }
    getSheet() { return this.sh; }
    getRow() { return this.r; }
    getColumn() { return this.c; }
    getNumRows() { return this.nr; }
    getNumColumns() { return this.nc; }
    grid(f) { const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(f(this.r + i, this.c + j)); out.push(row); } return out; }
    getValues() { return this.grid((r, c) => this.sh.cell(r, c)); }
    getDisplayValues() { return this.grid((r, c) => this.sh.txt(r, c)); }
    getValue() { return this.sh.cell(this.r, this.c); }
    getDisplayValue() { return this.sh.txt(this.r, this.c); }
    getRichTextValue() { return {getText: () => this.getDisplayValue(), getLinkUrl: () => ''}; }
    each(f) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) f(this.r + i, this.c + j, i, j); }
    write(r, c, val, extra) {
      const shown = val instanceof Date ? formatDate(val, ET, 'M/d/yyyy H:mm') : String(val == null ? '' : val);
      this.sh.put(r, c, val, shown);
      OPS.push(Object.assign({kind: 'value', sheet: this.sh.name, r, c, value: val}, extra || {}));
    }
    setValue(v) { this.each((r, c) => this.write(r, c, v)); return this; }
    setValues(vs) { this.each((r, c, i, j) => this.write(r, c, vs[i][j])); return this; }
    clearContent() { return this.setValue(''); }
    setRichTextValue(rt) { const t = rt && rt.text != null ? rt.text : ''; this.each((r, c) => this.write(r, c, t, {stamp: true})); return this; }
    fmt(prop, val) { this.each((r, c) => OPS.push({kind: 'format', sheet: this.sh.name, r, c, prop, val})); return this; }
    setHorizontalAlignment(a) { return this.fmt('h', a); }
    setVerticalAlignment(a) { return this.fmt('v', a); }
    setWrap(w) { return this.fmt('wrap', !!w); }
    setNote(n) { this.each((r, c) => OPS.push({kind: 'note', sheet: this.sh.name, r, c, text: String(n || '')})); return this; }
    clearNote() { return this.setNote(''); }
    setDataValidation() { return this; }   // the tick dropdowns already exist in the workbook
  }

  window.SpreadsheetApp = {
    getActive: () => ({getSheets: () => BOOK.sheets, getSheetByName: n => BOOK.sheets.find(s => s.name === n) || null,
      toast: (msg) => { if (typeof window.flash === 'function') window.flash(msg); }}),
    getActiveSpreadsheet: () => window.SpreadsheetApp.getActive(),
    flush: () => {},
    newRichTextValue: () => { const o = {text: ''}; const b = {setText(t) { o.text = t; return b; }, setTextStyle() { return b; }, build() { return o; }}; return b; },
    newTextStyle: () => chainProxy({}),
    newDataValidation: () => chainProxy({})
  };
  window.Utilities = {formatDate};
  window.Session = {getScriptTimeZone: () => ET};
  // Parts of the Google script that are not built for Excel yet (stage 3: moves between tabs).
  window.stageEdit_ = () => false;
  ['moveLead_', 'sortDemo_', 'fillRemindAt_', 'reschedToDemo_', 'sortResched_'].forEach(n => {
    window[n] = () => { throw new Error(n + ' is not built in the Excel version yet'); };
  });

  // ---------- reading and writing the workbook ----------
  async function loadBook() {
    await Excel.run(async ctx => {
      const wss = ctx.workbook.worksheets; wss.load('items/name');
      await ctx.sync();
      const want = wss.items.filter(w => LEAD_TABS.test(w.name));
      const used = want.map(w => { const u = w.getUsedRangeOrNullObject(true); u.load('rowIndex,rowCount,columnIndex,columnCount'); return u; });
      await ctx.sync();
      const rngs = want.map((w, i) => {
        const u = used[i]; if (u.isNullObject) return null;
        const g = w.getRangeByIndexes(0, 0, u.rowIndex + u.rowCount, u.columnIndex + u.columnCount);
        g.load('values,text,numberFormat'); return g;
      });
      await ctx.sync();
      BOOK = {sheets: want.map((w, i) => {
        const g = rngs[i]; if (!g) return new Sheet(w.name, [], []);
        const vals = g.values.map((row, ri) => row.map((x, ci) => (typeof x === 'number' && isDateFormat(g.numberFormat[ri][ci])) ? serialToDate(x) : x));
        return new Sheet(w.name, vals, g.text);
      })};
    });
  }
  const looksLikeNumber = s => typeof s === 'string' && s !== '' && /^[\d\s\/:.\-]+(\s*[AP]M)?$/i.test(s.trim());
  async function flush() {
    if (!OPS.length) return;
    const ops = OPS; OPS = [];
    await Excel.run(async ctx => {
      const ws = n => ctx.workbook.worksheets.getItem(n);
      for (const o of ops) {
        const cell = ws(o.sheet).getRangeByIndexes(o.r - 1, o.c - 1, 1, 1);
        if (o.kind === 'value') {
          let v = o.value;
          if (v instanceof Date) v = (v.getTime() + offMin(v, ET) * 60000) / 86400000 + 25569;
          if (looksLikeNumber(v)) cell.numberFormat = [['@']];   // keep "11/17 5:29 PM" as text, never a date
          cell.values = [[v]];
          if (o.stamp) { cell.format.font.color = '#1e8e3e'; cell.format.font.bold = true; cell.format.font.size = 9; }
        } else if (o.kind === 'format') {
          if (o.prop === 'h') cell.format.horizontalAlignment = {left: 'Left', right: 'Right', center: 'Center'}[o.val] || 'General';
          if (o.prop === 'v') cell.format.verticalAlignment = {top: 'Top', bottom: 'Bottom', middle: 'Center'}[o.val] || 'Bottom';
          if (o.prop === 'wrap') cell.format.wrapText = o.val;
        }
      }
      await ctx.sync();
      // Hover notes (Excel "Notes", ExcelApi 1.18). If this Excel can't do notes, the stamp still lands.
      const notes = ops.filter(o => o.kind === 'note');
      if (notes.length && Office.context.requirements.isSetSupported('ExcelApi', '1.18')) {
        try {
          const found = notes.map(o => { const a = colLetter_(o.c) + o.r, n = ws(o.sheet).notes.getItemOrNullObject(a); n.load('content'); return n; });
          await ctx.sync();
          notes.forEach((o, i) => {
            if (!found[i].isNullObject) found[i].delete();
            if (o.text) ctx.workbook.notes.add(ws(o.sheet).getRangeByIndexes(o.r - 1, o.c - 1, 1, 1), o.text);
          });
          await ctx.sync();
        } catch (e) { console.warn('Note not saved:', e); }
      }
    });
  }

  // ---------- google.script.run ----------
  const ready = new Promise(res => Office.onReady(() => res()));
  let chain = Promise.resolve();
  function enqueue(job) { const p = chain.then(job); chain = p.catch(() => {}); return p; }
  async function callServer(name, args) {
    await ready;
    await loadBook();
    const fn = window[name];
    if (typeof fn !== 'function') throw new Error(name + ' is not in the Excel version');
    let out;
    try { out = fn.apply(null, args); } catch (e) { OPS = []; throw e; }
    await flush();
    return out === undefined ? undefined : JSON.parse(JSON.stringify(out));
  }
  function runner() {
    let ok = null, fail = null;
    const base = {withSuccessHandler(f) { ok = f; return p; }, withFailureHandler(f) { fail = f; return p; }};
    const p = new Proxy(base, {get: (t, k) => k in t ? t[k] : (...args) => {
      enqueue(() => callServer(k, args)).then(v => { if (ok) ok(v); }, e => { console.error(e); if (fail) fail(e); });
    }});
    return p;
  }
  window.google = {script: {}};
  Object.defineProperty(window.google.script, 'run', {get: runner});

  // ---------- Dawn's own ticks from the dropdown (stamp + hover note, like onEdit in Google) ----------
  const TICK_HEAD = /^(call [1-4]|text|email|last call|reminder (call|text|email)|roland reminded)$/i;
  function handEdit(sheetName, address) {
    return enqueue(async () => {
      await loadBook();
      const sh = BOOK.sheets.find(s => s.name === sheetName); if (!sh) return;
      const a = address.split('!').pop().replace(/\$/g, '');
      const rg = sh.getRange(a);
      if (rg.getNumRows() * rg.getNumColumns() > 60) return;
      const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(x => String(x).trim());
      let tick = true;
      for (let c = rg.getColumn(); c < rg.getColumn() + rg.getNumColumns(); c++) if (!TICK_HEAD.test(head[c - 1] || '')) tick = false;
      if (!tick || rg.getRow() < 2) return;
      onEdit({range: rg});
      await flush();
    });
  }
  ready.then(() => Excel.run(async ctx => {
    ctx.workbook.worksheets.onChanged.add(ev => {
      if (ev.triggerSource === 'ThisLocalAddin') return;
      return Excel.run(async c2 => {
        const w = c2.workbook.worksheets.getItem(ev.worksheetId); w.load('name'); await c2.sync();
        if (LEAD_TABS.test(w.name)) await handEdit(w.name, ev.address);
      }).catch(e => console.warn('Hand tick not stamped:', e));
    });
    await ctx.sync();
  })).catch(e => console.warn('Could not watch for hand ticks:', e));

  window.__bw = {formatDate, serialToDate, loadBook, flush, setBook: b => { BOOK = b; }, Sheet, ops: () => OPS};
})();
