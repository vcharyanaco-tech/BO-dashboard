/* Behavioral tests for app.js — eval'd in the same scope as app.js by
 * tests/dashboard.test.js, so app functions (esc, render, autoLoad, ...) and
 * harness globals (T, els, fetchHandler, chartCreated, XLSX) are in scope.
 * Plain file, so normal JS escaping applies. Ends by setting process.exitCode. */
(async () => {
  // ---------- pure helpers ----------
  T('esc escapes html', esc('<b>"x"&\'') === '&lt;b&gt;&quot;x&quot;&amp;&#39;');
  const csv = 'date,bo,so,sol,kpi,value\n2026-10-30,"Adhoyi B.O",SO1,1,Accounts Opened,5\n2026-10-30,"BO ""Q"" Chandigarh",SO2,2,Accounts Opened,7\n';
  const pr = parseCsvText(csv);
  T('csv parse rows+quotes', pr.length === 3 && pr[1][1] === 'Adhoyi B.O' && pr[2][1] === 'BO "Q" Chandigarh' && pr[2][5] === '7');
  T('matchBO exact', matchBO('Adhoyi B.O') === 'Adhoyi B.O');
  T('matchBO clean-index', matchBO('adhoyi  b.o') === 'Adhoyi B.O');
  T('matchBO unknown passthrough', matchBO('Random Place') === 'Random Place');
  T('isSORow catches S.O and SO rows', isSORow('Adhoya S.O') === true && isSORow('Achina SO') === true);
  T('isSORow keeps real BOs (master-guarded)', isSORow('Adhoyi B.O') === false && isSORow('Adhoyi B.O') === false);

  // ---------- loadHistory: no global side effects, SO rows dropped ----------
  fetchHandler = async (url) => {
    if (url.includes('history.csv')) return { ok: true, text: async () => 'date,bo,so,sol,kpi,value\n2026-10-30,Adhoyi B.O,SO1,1,Accounts Opened,5\n2026-10-30,Adhoya S.O,,999,Accounts Opened,50\n2026-10-31,Adhoyi B.O,SO1,1,Accounts Opened,8\n2026-10-31,Achhanpur B.O,SO2,2,Accounts Opened,3\n' };
    return { ok: false, status: 404, statusText: 'nf' };
  };
  const hist = await loadHistory();
  T('loadHistory returns data w/o touching globals', hist && hist.rows.length === 3 && ROWS.length === 0);
  T('loadHistory drops SO rows + canonicalizes BOs', !hist.rows.some(r => /s\.?\s?o\.?\s*$/i.test(r.bo)) && hist.rows.every(r => ['Adhoyi B.O', 'Achhanpur B.O'].includes(r.bo)));
  ROWS = hist.rows; METRICS = hist.metrics;

  // ---------- overlay: same date replaced, other dates kept ----------
  overlayKpiParts([{ label: 'Accounts Opened', date: new Date('2026-10-31T00:00:00Z'), totals: { 'Adhoyi B.O': 99, 'Achhanpur B.O': 42 }, so: { 'Adhoyi B.O': 'SO1', 'Achhanpur B.O': 'SO2' }, sol: {} }]);
  const a31 = ROWS.filter(r => r.date === '2026-10-31');
  T('overlay replaces same-date rows', a31.length === 2 && a31.find(r => r.bo === 'Adhoyi B.O')['Accounts Opened'] === 99);
  T('overlay keeps other dates', ROWS.some(r => r.date === '2026-10-30' && r['Accounts Opened'] === 5));

  // ---------- finishIngest + render ----------
  finishIngest();
  T('DATES sorted', DATES.join(',') === '2026-10-30,2026-10-31');
  T('dataAsOf set', els['dataAsOf'].textContent === '2026-10-31');
  els['mode'].value = 'upto'; els['from'].value = '2026-10-30'; els['to'].value = '2026-10-31';
  els['bo'].value = ''; els['so'].value = ''; els['groupBy'].value = 'bo'; els['quick'].value = 'all';
  render();
  const tbl = els['tbl'].innerHTML;
  T('table has BO rows', tbl.includes('Adhoyi B.O') && tbl.includes('Achhanpur B.O'));
  T('table has Total col+value', tbl.includes('Total') && tbl.includes('>104<')); // 5+99
  T('no SO pseudo-BO in table', !/Adhoya/.test(tbl));
  T('KPI delta vs prev day shown', els['kpis'].innerHTML.includes('vs 2026-10-30') && /[▲▼] \d/.test(els['kpis'].innerHTML));
  const deltaTxt = els['kpis'].innerHTML.match(/class="d (good|bad)">([^<]*)</);
  T('KPI delta value correct (+136: (99+42)-5)', deltaTxt && deltaTxt[2].includes('136') && deltaTxt[1] === 'good');
  T('chart created once', chartCreated === 1);
  render();
  T('chart updated in place', chartCreated === 1 && chartUpdated >= 1);
  T('tblInfo + printHead filled', els['tblInfo'].textContent.includes('sorted by') && els['printHead'].innerHTML.includes('2026-10-31'));

  // ---------- sorting / quick view / grouping ----------
  const savedROWS = ROWS;
  ROWS = [...ROWS, ...Array.from({ length: 15 }, (_, i) => ({ date: '2026-10-31', bo: 'Fake BO ' + i, so: 'SOX', sol: '', 'Accounts Opened': i + 1 }))];
  els['quick'].value = 'top'; render();
  const topRows = (els['tbl'].innerHTML.match(/<tbody>([\s\S]*?)<\/tbody>/)[1].match(/<tr>/g) || []).length;
  T('top10 quick view shows 10 rows', topRows === 10 && els['tblInfo'].textContent.includes('top 10'));
  ROWS = savedROWS; els['quick'].value = 'all'; render();
  sortBy('total'); const i1 = els['tblInfo'].textContent; const t1 = els['tbl'].innerHTML;
  sortBy('total'); const i2 = els['tblInfo'].textContent; const t2 = els['tbl'].innerHTML;
  T('sortBy toggles direction', i1.includes('▲') && i2.includes('▼') && t1 !== t2);
  els['groupBy'].value = 'so'; sortState = { key: 'name', dir: 1 }; render();
  T('SO-wise grouping', els['tblTitle'].textContent.includes('SO-wise') && els['tbl'].innerHTML.includes('SO1') && els['tblInfo'].textContent.includes('SOs'));
  els['groupBy'].value = 'bo';

  // ---------- filters ----------
  els['mode'].value = 'single'; els['to'].value = '2026-10-30'; render();
  T('single-date filter', els['tbl'].innerHTML.includes('>5<') && !els['tbl'].innerHTML.includes('>99<'));
  els['mode'].value = 'upto'; els['to'].value = '2026-10-31'; render();
  els['bo'].value = 'achhanpur'; render();
  T('BO substring filter', els['tbl'].innerHTML.includes('Achhanpur B.O') && !els['tbl'].innerHTML.includes('Adhoyi B.O'));
  els['bo'].value = '';
  els['mode'].value = 'single'; updateMode(); T('from hidden in single', els['fromWrap'].style.display === 'none');
  els['mode'].value = 'range'; updateMode(); T('from shown in range', els['fromWrap'].style.display === '');
  els['mode'].value = 'upto'; updateMode(); T('from hidden in upto', els['fromWrap'].style.display === 'none');

  // ---------- race guard: manual upload must survive in-flight auto-load ----------
  ROWS = [{ date: '2026-10-30', bo: 'Adhoyi B.O', so: 'SO1', sol: '1', 'Accounts Opened': 5 }]; METRICS = ['Accounts Opened']; DATES = ['2026-10-30'];
  userLoaded = true;
  fetchHandler = async () => ({ ok: true, text: async () => 'date,bo,so,sol,kpi,value\n2026-01-01,Adhoyi B.O,,1,X,1\n' });
  await autoLoad();
  T('race guard: auto skipped, ROWS intact', els['autoStatus'].textContent.includes('Manual file loaded') && ROWS.length === 1 && ROWS[0].date === '2026-10-30');
  userLoaded = false;

  // ---------- full autoLoad: manifest + xlsx (SO row + Total row skipped) ----------
  fetchHandler = async (url) => {
    if (url.includes('history.csv')) return { ok: true, text: async () => 'date,bo,so,sol,kpi,value\n2026-10-30,Adhoyi B.O,SO1,1,Accounts Opened,5\n' };
    if (url.includes('manifest.json')) return { ok: true, json: async () => [{ file: 'data/exports/accounts_opened_merged.xlsx', label: 'Accounts Opened', date: '2026-10-31' }] };
    if (url.includes('.xlsx')) return { ok: true, arrayBuffer: async () => new ArrayBuffer(8) };
    return { ok: false, status: 404, statusText: 'nf' };
  };
  XLSX = {
    read: () => ({ Sheets: { S: 1 }, SheetNames: ['S'] }),
    utils: { sheet_to_json: () => [['Name', 'SOL ID / BOCODE', 'SO Name', 'MIS', 'RD'],
      ['Adhoya S.O', '13320501', '', 10, 10], ['Adhoyi B.O', '1', 'SO1', 3, 4], ['Total:', '', '', 7, 0]] },
  };
  ROWS = []; METRICS = []; DATES = [];
  await autoLoad();
  T('autoLoad history+manifest', ROWS.length === 2 && DATES.join(',') === '2026-10-30,2026-10-31');
  const r31 = ROWS.find(r => r.date === '2026-10-31' && r.bo === 'Adhoyi B.O');
  T('manifest date used; SO+Total rows skipped; product cols summed', !!r31 && r31['Accounts Opened'] === 7 && !ROWS.some(r => /s\.?\s?o\.?\s*$/i.test(r.bo) || String(r.bo).startsWith('Total')));
  T('autoStatus summary', els['autoStatus'].textContent.includes('history 2 day(s)') && els['autoStatus'].textContent.includes('1/1 KPI export'));

  console.log(__failures ? `FAILED: ${__failures} assertion(s)` : 'ALL PASS');
  process.exitCode = __failures ? 1 : 0;
})().catch(e => { console.error('TEST BODY ERROR', e); process.exit(1); });
