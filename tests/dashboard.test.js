#!/usr/bin/env node
/*
 * Test runner for the dashboard (no external dependencies).
 *
 *     node tests/dashboard.test.js
 *
 * Loads the real app.js + bo_master.js with a stubbed DOM/fetch/XLSX/Chart,
 * checks the repo files (index.html hardening, manifest.json, history.csv),
 * then evals tests/app.testbody.js in the same scope so the tests can call
 * the app's functions directly. Exit code 0 = all assertions passed.
 */
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const appCode = fs.readFileSync(path.join(ROOT, 'app.js'), 'utf8');
const masterCode = fs.readFileSync(path.join(ROOT, 'bo_master.js'), 'utf8');
global.BO_MASTER = new Function(masterCode + '\nreturn BO_MASTER;')();

// ---- stubs ---------------------------------------------------------------
const els = {};
global.els = els;
const makeEl = (id) => ({
  id, value: '', textContent: '', innerHTML: '', style: {}, dataset: {},
  classList: { add() {}, remove() {} },
  onclick: null, onchange: null, oninput: null,
  ondragover: null, ondragleave: null, ondrop: null,
  _h: {},
  addEventListener(t, fn) { this._h[t] = fn; },
});
global.document = { getElementById(id) { return els[id] || (els[id] = makeEl(id)); } };
global.window = { print() {} };
global.fetchHandler = async () => ({ ok: false, status: 404, statusText: 'not found' });
global.fetch = (url) => global.fetchHandler(url);
global.XLSX = {
  read: () => ({ Sheets: {}, SheetNames: [] }),
  utils: { sheet_to_json: () => [], json_to_sheet: () => ({}), book_new: () => ({}), book_append_sheet: () => {} },
  writeFile: () => {},
};
global.chartCreated = 0;
global.chartUpdated = 0;
global.Chart = class {
  constructor(el, cfg) { this.data = cfg.data; this.options = cfg.options; global.chartCreated++; }
  update() { global.chartUpdated++; }
  destroy() {}
};
global.alert = () => {};
global.FileReader = class {};
global.__failures = 0;
global.T = (name, cond) => { console.log((cond ? 'PASS' : 'FAIL') + ' ' + name); if (!cond) global.__failures++; };

// ---- repo file checks ------------------------------------------------------
const html = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');
T('index.html loads external app.js', /<script src="app\.js"><\/script>/.test(html));
T('every <script> tag has a src (no inline script)', (html.match(/<script(?![^>]*src=)[^>]*>/g) || []).length === 0);
T('CSP meta restricts script-src to self + jsdelivr', /http-equiv="Content-Security-Policy"/.test(html) && /script-src 'self' https:\/\/cdn\.jsdelivr\.net/.test(html));
T('both CDN scripts carry SRI sha384 + crossorigin', (html.match(/integrity="sha384-[A-Za-z0-9+/]{64}" crossorigin="anonymous"/g) || []).length === 2);
T('index.html has no inline on* handlers (CSP script-src has no unsafe-inline)', !/\son[a-z]+\s*=\s*"/i.test(html));
T('app.js wires controls via addEventListener', (appCode.match(/addEventListener/g) || []).length >= 10);
T('app.js has no dead totalOf/ingestRows', !/function (totalOf|ingestRows)\b/.test(appCode));

const manifest = JSON.parse(fs.readFileSync(path.join(ROOT, 'data/exports/manifest.json'), 'utf8'));
T('manifest.json is a non-empty array with file+label', Array.isArray(manifest) && manifest.length > 0 && manifest.every(e => e.file && e.label));
T('manifest-referenced files exist', manifest.every(e => fs.existsSync(path.join(ROOT, e.file))));

const histLines = fs.readFileSync(path.join(ROOT, 'data/history.csv'), 'utf8').trim().split('\n');
T('history.csv header is date,bo,so,sol,kind,kpi,value', histLines[0] === 'date,bo,so,sol,kind,kpi,value');
T('history.csv: ISO dates, >=1 row, kind=bo rows are never SO-named, kind=so rows exist', histLines.length > 1 && histLines.slice(1).every(l => {
  const c = l.split(',');
  return /^\d{4}-\d{2}-\d{2}$/.test(c[0]) && (c[4] !== 'bo' || !/\bs\.?\s?o\.?\s*$/i.test(c[1]));
}) && histLines.slice(1).some(l => l.split(',')[4] === 'so'));

// ---- run app.js + behavioral test body in one scope -------------------------
const testBody = fs.readFileSync(path.join(__dirname, 'app.testbody.js'), 'utf8');
(0, eval)(appCode + '\n;\n' + testBody);
