/* Behavioral tests for app.js — eval'd in the same scope as app.js by
 * tests/dashboard.test.js, so app functions (esc, render, autoLoad, ...) and
 * harness globals (T, els, fetchHandler, chartCreated, XLSX) are in scope.
 * Plain file, so normal JS escaping applies. Ends by setting process.exitCode. */
(async () => {
  // ---------- pure helpers ----------
  T('esc escapes html', esc('<b>"x"&\'') === '&lt;b&gt;&quot;x&quot;&amp;&#39;');
  T('fmtDate renders DD-MM-YYYY', fmtDate('2026-11-02') === '02-11-2026' && fmtDate('') === '');
  const csv = 'date,bo,so,sol,kpi,value\n2026-11-01,"Adhoyi B.O",SO1,1,Accounts Opened,5\n2026-11-01,"BO ""Q"" Chandigarh",SO2,2,Accounts Opened,7\n';
  const pr = parseCsvText(csv);
  T('csv parse rows+quotes', pr.length === 3 && pr[1][1] === 'Adhoyi B.O' && pr[2][1] === 'BO "Q" Chandigarh' && pr[2][5] === '7');

  // ---------- identity is the BO code (no fuzzy name matching) ----------
  T('findOffice resolves by code', findOffice('Adhoyi B.O', '', 'I1895').code === 'I1895');
  T('code wins over a conflicting name', findOffice('Adhoyi B.O', 'Adhoya S.O', 'J6642').code === 'J6642');
  T('findOffice resolves by unique name', findOffice('adhoyi  b.o', '', '').code === 'I1895');
  T('findOffice unknown returns null', findOffice('Random Place', '', 'ZZZZZ') === null);
  // A name shared by two offices resolves only by code — that is why codes are the key.
  const twin = MASTER_LIST.filter(o => MASTER_LIST.filter(x => normName(x.name) === normName(o.name)).length > 1)[0];
  T('master sheet has duplicate names', !!twin);
  T('ambiguous name alone does not resolve', findOffice(twin.name, '', '') === null);
  T('ambiguous name resolves by code', findOffice('', '', twin.code) === twin || (findOffice('', '', twin.code).code === twin.code && findOffice('', '', twin.code).so === twin.so));
  T('officeInfo carries division + sub division', (() => {
    const i = officeInfo('Adhoyi B.O', 'Adhoya S.O', 'I1895');
    return i.kind === 'bo' && i.id === 'I1895' && i.dv === 'Ambala Division' && i.sd === 'Ambala City Sub- Division';
  })());
  T('officeInfo leaves an unknown office unpinned', (() => {
    const i = officeInfo('Nowhere B.O', 'X S.O', 'Z9999');
    return i.kind === 'bo' && i.id === 'Z9999' && i.dv === '' && i.sd === '';
  })());
  T('isSORow catches S.O / SO rows', isSORow('Adhoya S.O', '', '13320501') === true && isSORow('Achina SO', '', '') === true);
  T('isSORow never demotes a real master BO', isSORow('Adhoyi B.O', 'Adhoya S.O', 'I1895') === false);
  T('isOfficeRow: SO summary row (blank SO Name) vs BO row', isOfficeRow('Adhoya S.O', '', '13320501') === true && isOfficeRow('Adhoyi B.O', 'Adhoya S.O', 'I1895') === false);

  // ---------- loadHistory: no global side effects; kind column respected ----------
  fetchHandler = async (url) => {
    if (url.includes('history.csv')) return { ok: true, text: async () => 'date,bo,so,sol,kind,kpi,value\n2026-11-01,Adhoyi B.O,SO1,1,bo,Accounts Opened,5\n2026-11-01,Adhoya S.O,,999,so,Accounts Opened,50\n2026-11-02,Adhoyi B.O,SO1,1,bo,Accounts Opened,8\n2026-11-02,Achhanpur B.O,SO2,2,bo,Accounts Opened,3\n' };
    return { ok: false, status: 404, statusText: 'nf' };
  };
  const hist = await loadHistory();
  T('loadHistory returns data w/o touching globals', hist && hist.rows.length === 4 && ROWS.length === 0);
  T('loadHistory keeps SO row as kind=so, keys BO rows by code', hist.rows.some(r => r.kind === 'so' && r.bo === 'Adhoya S.O' && r['Accounts Opened'] === 50) && hist.rows.filter(r => r.kind !== 'so').every(r => ['Adhoyi B.O', 'Achhanpur B.O'].includes(r.bo)) && hist.rows.some(r => r.id === 'I1895' && r.dv === 'Ambala Division'));
  ROWS = hist.rows; METRICS = hist.metrics;

  // ---------- overlay: same date replaced, other dates kept ----------
  overlayKpiParts([{ label: 'Accounts Opened', date: new Date('2026-11-02T00:00:00Z'),
    bos: [{ id: 'I1895', name: 'Adhoyi B.O', so: 'SO1', sol: 'I1895', dv: 'Ambala Division', sd: 'Ambala City Sub- Division', value: 99 },
          { id: 'J6642', name: 'Achhanpur B.O', so: 'SO2', sol: 'J6642', dv: 'Karnal Division', sd: 'Karnal East Sub Division', value: 42 }],
    offices: [{ id: 'so:adhoya s.o', name: 'Adhoya S.O', so: 'Adhoya S.O', sol: '999', value: 50 }] }]);
  const a31 = ROWS.filter(r => r.date === '2026-11-02' && r.kind !== 'so');
  T('overlay replaces same-date BO rows', a31.length === 2 && a31.find(r => r.bo === 'Adhoyi B.O')['Accounts Opened'] === 99);
  T('overlay adds SO-level row as kind=so (separate)', ROWS.some(r => r.kind === 'so' && r.bo === 'Adhoya S.O' && r.date === '2026-11-02' && r['Accounts Opened'] === 50));
  T('overlay keeps other dates', ROWS.some(r => r.date === '2026-11-01' && r['Accounts Opened'] === 5));

  // ---------- finishIngest + render ----------
  finishIngest();
  T('DATES sorted', DATES.join(',') === '2026-11-01,2026-11-02');
  T('dataAsOf set in DD-MM-YYYY', els['dataAsOf'].textContent === '02-11-2026');
  els['mode'].value = 'upto'; els['from'].value = '2026-11-01'; els['to'].value = '2026-11-02';
  els['bo'].value = ''; els['so'].value = ''; els['groupBy'].value = 'bo'; els['quick'].value = 'all';
  els['division'].value = ''; els['subDivision'].value = '';
  render();
  const tbl = els['tbl'].innerHTML;
  T('table has BO rows', tbl.includes('Adhoyi B.O') && tbl.includes('Achhanpur B.O'));
  T('BO rows carry division columns', tbl.includes('<th class="sortable" data-k="dv">Division') && tbl.includes('Ambala Division') && tbl.includes('Karnal Division'));
  const zeroRow = tbl.split('</tr>').find(r => r.includes('<td>Abubshahar BO</td>'));
  T('every master BO is shown (missing report row -> 0)', !!zeroRow && (zeroRow.match(/>0<\/b>/g) || []).length >= 2);
  T('BO count includes the full master list', els['tblInfo'].textContent.includes(`${MASTER_LIST.length} of ${MASTER_LIST.length} BOs`));
  T('table has Total col+value', tbl.includes('Total') && tbl.includes('>104<')); // 5+99
  T('no SO pseudo-BO row in the BO table', !/<tr><td>Adhoya S\.O<\/td>/.test(tbl));
  T('KPI delta vs prev day shown', els['kpis'].innerHTML.includes('vs 01-11-2026') && /[▲▼] \d/.test(els['kpis'].innerHTML));
  const deltaTxt = els['kpis'].innerHTML.match(/class="d (good|bad)">([^<]*)</);
  T('KPI delta value correct (+136: (99+42)-5)', deltaTxt && deltaTxt[2].includes('136') && deltaTxt[1] === 'good');
  T('chart created once', chartCreated === 1);
  render();
  T('chart updated in place', chartCreated === 1 && chartUpdated >= 1);
  T('tblInfo + printHead filled', els['tblInfo'].textContent.includes('sorted by') && els['printHead'].innerHTML.includes('02-11-2026'));
  T('chart labels use DD-MM-YYYY', chartInstance.data.labels.join(',') === '01-11-2026,02-11-2026');

  // ---------- sorting / quick view / grouping ----------
  const savedROWS = ROWS;
  ROWS = [...ROWS, ...Array.from({ length: 15 }, (_, i) => ({ date: '2026-11-02', id: 'FAKE' + i, bo: 'Fake BO ' + i, so: 'SOX', sol: '', kind: 'bo', 'Accounts Opened': i + 1 }))];
  els['quick'].value = 'top'; render();
  const topRows = (els['tbl'].innerHTML.match(/<tbody>([\s\S]*?)<\/tbody>/)[1].match(/<tr>/g) || []).length;
  T('top10 quick view shows 10 rows', topRows === 10 && els['tblInfo'].textContent.includes('top 10'));
  ROWS = savedROWS; els['quick'].value = 'all'; render();
  sortBy('total'); const i1 = els['tblInfo'].textContent; const t1 = els['tbl'].innerHTML;
  sortBy('total'); const i2 = els['tblInfo'].textContent; const t2 = els['tbl'].innerHTML;
  T('sortBy toggles direction', i1.includes('▲') && i2.includes('▼') && t1 !== t2);

  // CSP-safe wiring: controls must be driven by listeners attached in app.js
  // (inline on* HTML attributes are refused by the CSP). Fire the stored handlers.
  T('all controls have addEventListener handlers', ['mode','from','to','bo','so','division','subDivision','groupBy','quick','btnExcel','btnPdf','tbl'].every(id => els[id] && els[id]._h && Object.keys(els[id]._h).length > 0));
  els['quick'].value = 'top'; els['quick']._h.change();
  T('quick-view listener applies top 10', els['tblInfo'].textContent.includes('top 10'));
  els['quick'].value = 'all'; els['quick']._h.change();
  const infoBefore = els['tblInfo'].textContent;
  els['tbl']._h.click({ target: { closest: () => ({ dataset: { k: 'name' } }) } });
  T('delegated th click sorts', els['tblInfo'].textContent.includes('sorted by BO') && els['tblInfo'].textContent !== infoBefore);
  els['groupBy'].value = 'so'; els['groupBy']._h.change();
  T('groupBy listener switches to SO-wise', els['tblTitle'].textContent.includes('SO-wise'));
  els['groupBy'].value = 'bo'; els['groupBy']._h.change();
  els['bo'].value = 'achhanpur'; els['bo']._h.input();
  await new Promise(r => setTimeout(r, 250)); // onFilterInput debounces 150ms
  T('BO input listener filters (debounced)', els['tbl'].innerHTML.includes('Achhanpur B.O') && !els['tbl'].innerHTML.includes('Adhoyi B.O'));
  els['bo'].value = ''; els['bo']._h.input();
  await new Promise(r => setTimeout(r, 250));
  els['groupBy'].value = 'so'; sortState = { key: 'name', dir: 1 }; render();
  T('SO-wise grouping', els['tblTitle'].textContent.includes('SO-wise') && els['tbl'].innerHTML.includes('SO1') && els['tblInfo'].textContent.includes('SOs'));
  els['groupBy'].value = 'bo';

  // ---------- filters ----------
  els['mode'].value = 'single'; els['to'].value = '2026-11-01'; render();
  T('single-date filter', els['tbl'].innerHTML.includes('>5<') && !els['tbl'].innerHTML.includes('>99<'));
  els['mode'].value = 'upto'; els['to'].value = '2026-11-02'; render();
  els['bo'].value = 'achhanpur'; render();
  T('BO substring filter', els['tbl'].innerHTML.includes('Achhanpur B.O') && !els['tbl'].innerHTML.includes('Adhoyi B.O'));
  els['bo'].value = '';
  els['mode'].value = 'single'; updateMode(); T('from hidden in single', els['fromWrap'].style.display === 'none');
  els['mode'].value = 'range'; updateMode(); T('from shown in range', els['fromWrap'].style.display === '');
  els['mode'].value = 'upto'; updateMode(); T('from hidden in upto', els['fromWrap'].style.display === 'none');

  // ---------- race guard: manual upload must survive in-flight auto-load ----------
  ROWS = [{ date: '2026-11-01', id: 'I1895', kind: 'bo', bo: 'Adhoyi B.O', so: 'SO1', sol: '1', dv: 'Ambala Division', sd: 'Ambala City Sub- Division', 'Accounts Opened': 5 }];
  METRICS = ['Accounts Opened']; DATES = ['2026-11-01'];
  userLoaded = true;
  fetchHandler = async () => ({ ok: true, text: async () => 'date,bo,so,sol,kpi,value\n2026-01-01,Adhoyi B.O,,1,X,1\n' });
  await autoLoad();
  T('race guard: auto skipped, ROWS intact', els['autoStatus'].textContent.includes('Manual file loaded') && ROWS.length === 1 && ROWS[0].date === '2026-11-01');
  userLoaded = false;

  // ---------- full autoLoad: manifest + xlsx (SO row + Total row skipped) ----------
  fetchHandler = async (url) => {
    if (url.includes('history.csv')) return { ok: true, text: async () => 'date,bo,so,sol,kpi,value\n2026-11-01,Adhoyi B.O,SO1,1,Accounts Opened,5\n' };
    if (url.includes('manifest.json')) return { ok: true, json: async () => [{ file: 'data/exports/accounts_opened_merged.xlsx', label: 'Accounts Opened', date: '2026-11-02' }] };
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
  T('autoLoad history+manifest', ROWS.length === 3 && DATES.join(',') === '2026-11-01,2026-11-02');
  const r31 = ROWS.find(r => r.date === '2026-11-02' && r.bo === 'Adhoyi B.O');
  T('manifest date used; product cols summed; Total row skipped', !!r31 && r31['Accounts Opened'] === 7 && !ROWS.some(r => String(r.bo).startsWith('Total')));
  T('office row kept separately as kind=so (not in BO totals)', ROWS.some(r => r.kind === 'so' && r.bo === 'Adhoya S.O' && r['Accounts Opened'] === 20) && !ROWS.some(r => r.kind !== 'so' && /s\.?\s?o\.?\s*$/i.test(r.bo)));
  T('autoStatus summary', els['autoStatus'].textContent.includes('history 2 day(s)') && els['autoStatus'].textContent.includes('1/1 KPI export'));

  // ---------- SO-wise view: ΣBO and SO-reported kept separate; '(no SO)' bucket ----------
  ROWS = [
    { date: '2026-11-02', id: 'I1895', bo: 'Adhoyi B.O', so: 'Adhoya S.O', sol: 'I1895', kind: 'bo', dv: 'Ambala Division', sd: 'Ambala City Sub- Division', 'Accounts Opened': 7 },
    { date: '2026-11-02', id: 'AV785', bo: 'Brahman Majra B.O', so: '', sol: '', kind: 'bo', dv: 'Ambala Division', sd: 'Ambala Sub-Division', 'Accounts Opened': 2 },
    { date: '2026-11-02', id: 'so:adhoya s.o', bo: 'Adhoya S.O', so: 'Adhoya S.O', sol: '13320501', kind: 'so', 'Accounts Opened': 20 },
  ];
  METRICS = ['Accounts Opened']; DATES = ['2026-11-02'];
  finishIngest();
  els['mode'].value = 'single'; els['to'].value = '2026-11-02';
  els['bo'].value = ''; els['so'].value = ''; els['quick'].value = 'all';
  els['groupBy'].value = 'so'; render();
  const soTbl = els['tbl'].innerHTML;
  T('SO view has separate ΣBO and SO columns', soTbl.includes('(ΣBO)') && soTbl.includes('(SO)'));
  T('SO row value shown in SO column', soTbl.includes('>20<'));
  T('ΣBO excludes SO row (7, not 27)', soTbl.includes('>7<') && !soTbl.includes('>27<'));
  T('blank-SO BO lands in (no SO) bucket', soTbl.includes('(no SO)'));
  T('KPI cards count BO rows only (9, not 29)', els['kpis'].innerHTML.includes('>9<') && !els['kpis'].innerHTML.includes('>29<'));
  els['groupBy'].value = 'bo'; render();
  T('BO view hides SO summary rows', !/<tr><td>Adhoya S\.O<\/td>/.test(els['tbl'].innerHTML) && els['tbl'].innerHTML.includes('Brahman Majra B.O'));

  // ---------- division / sub-division filter + grouping ----------
  ROWS = [
    { date: '2026-11-02', id: 'I1895', bo: 'Adhoyi B.O', so: 'Adhoya S.O', sol: 'I1895', kind: 'bo', dv: 'Ambala Division', sd: 'Ambala City Sub- Division', 'Accounts Opened': 7 },
    { date: '2026-11-02', id: 'J6642', bo: 'Achhanpur B.O', so: 'Nissing S.O', sol: 'J6642', kind: 'bo', dv: 'Karnal Division', sd: 'Karnal East Sub Division', 'Accounts Opened': 3 },
    { date: '2026-11-02', id: 'so:adhoya s.o', bo: 'Adhoya S.O', so: 'Adhoya S.O', sol: '13320501', kind: 'so', 'Accounts Opened': 20 },
  ];
  METRICS = ['Accounts Opened']; DATES = ['2026-11-02'];
  els['groupBy'].value = 'bo'; els['division'].value = ''; els['subDivision'].value = '';
  finishIngest(); // rebuilds the division / sub-division option lists
  els['mode'].value = 'single'; els['to'].value = '2026-11-02';
  els['bo'].value = ''; els['so'].value = ''; els['quick'].value = 'all';
  T('division options come from the master sheet', els['division'].innerHTML.includes('Ambala Division') && els['division'].innerHTML.includes('Sonipat Division'));
  const ambSubs = [...new Set(MASTER_LIST.filter(o => o.division === 'Ambala Division').map(o => o.subDivision).filter(Boolean))];
  const otherSubs = [...new Set(MASTER_LIST.filter(o => o.division !== 'Ambala Division').map(o => o.subDivision).filter(Boolean))];
  els['division'].value = 'Ambala Division'; els['division']._h.change();
  T('sub division list narrows to the selected division', ambSubs.length > 1 && ambSubs.every(s => els['subDivision'].innerHTML.includes(s)) && !otherSubs.some(s => els['subDivision'].innerHTML.includes('>' + s + '<')));
  T('division filter drops other divisions BOs', els['tbl'].innerHTML.includes('Adhoyi B.O') && !els['tbl'].innerHTML.includes('Achhanpur B.O'));
  const ambCount = MASTER_LIST.filter(o => o.division === 'Ambala Division').length;
  T('division filter keeps unreported master BOs of that division', els['tblInfo'].textContent.includes(`${ambCount} of ${ambCount} BOs`));
  // Tier cards (0 / under 10 / under 100 accounts) must follow the same scope as the
  // BOs card, so expected counts are derived from the master sheet + report rows.
  const kpiVal = (label) => {
    const m = els['kpis'].innerHTML.match(new RegExp('<div class="l">' + label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '</div><div class="v">([^<]*)</div>'));
    return m ? m[1].replace(/,/g, '') : null;
  };
  const tierCheck = (ids) => {
    const tot = new Map();
    ROWS.filter(r => r.kind !== 'so' && ids.has(r.id)).forEach(r => tot.set(r.id, (tot.get(r.id) || 0) + (Number(r['Accounts Opened']) || 0)));
    const list = [...ids];
    return kpiVal('BOs') === String(list.length)
      && kpiVal('0 Accounts') === String(list.filter(id => !tot.get(id)).length)
      && kpiVal('Under 10 Accounts') === String(list.filter(id => (tot.get(id) || 0) < 10).length)
      && kpiVal('Under 100 Accounts') === String(list.filter(id => (tot.get(id) || 0) < 100).length);
  };
  const ambIds = new Set(MASTER_LIST.filter(o => o.division === 'Ambala Division').map(o => normCode(o.code) || normName(o.name)));
  T('tier cards count the selected division only', tierCheck(ambIds) && ambIds.size > 1 && kpiVal('0 Accounts') !== '0');
  T('BOs card names the selected division', els['kpis'].innerHTML.includes('<div class="n">Division: Ambala Division</div>'));
  els['subDivision'].value = 'Karnal East Sub Division'; els['subDivision']._h.change();
  await new Promise(r => setTimeout(r, 250));
  T('division + sub division combine with AND', els['tblInfo'].textContent.includes('0 of 0 BOs'));
  els['subDivision'].value = 'Ambala City Sub- Division'; els['subDivision']._h.change();
  await new Promise(r => setTimeout(r, 250));
  const cityCount = MASTER_LIST.filter(o => o.division === 'Ambala Division' && o.subDivision === 'Ambala City Sub- Division').length;
  T('sub division filter narrows within the division', els['tbl'].innerHTML.includes('Adhoyi B.O') && !els['tbl'].innerHTML.includes('Achhanpur B.O') && els['tblInfo'].textContent.includes(`${cityCount} of ${cityCount} BOs`));
  const cityIds = new Set(MASTER_LIST.filter(o => o.division === 'Ambala Division' && o.subDivision === 'Ambala City Sub- Division').map(o => normCode(o.code) || normName(o.name)));
  T('tier cards count the selected sub division', tierCheck(cityIds) && cityIds.size < ambIds.size && els['kpis'].innerHTML.includes('<div class="n">Sub Division: Ambala City Sub- Division</div>'));
  els['division'].value = ''; els['subDivision'].value = '';
  els['division']._h.change();
  T('clearing the division restores the full office list', els['tblInfo'].textContent.includes(`${MASTER_LIST.length} of ${MASTER_LIST.length} BOs`));
  const allIds = new Set(MASTER_LIST.map(o => normCode(o.code) || normName(o.name)));
  T('tier cards count all BOs when no division is selected', tierCheck(allIds) && els['kpis'].innerHTML.includes('<div class="n">All divisions (overall)</div>'));
  els['groupBy'].value = 'div'; render();
  const divTbl = els['tbl'].innerHTML;
  const nDiv = new Set(MASTER_LIST.map(o => o.division).filter(Boolean)).size;
  T('division-wise grouping shows every division', els['tblTitle'].textContent.includes('Division-wise') && divTbl.includes('Ambala Division') && divTbl.includes('Sonipat Division') && els['tblInfo'].textContent.includes(`${nDiv} of ${nDiv} Divisions`));
  T('division-wise keeps the SO figure separate from ΣBO', divTbl.includes('>7<') && divTbl.includes('>20<') && !divTbl.includes('>27<'));
  els['groupBy'].value = 'sub'; render();
  const nSub = new Set(MASTER_LIST.map(o => o.subDivision).filter(Boolean)).size;
  T('sub-division-wise grouping', els['tblTitle'].textContent.includes('Sub-division-wise') && els['tbl'].innerHTML.includes('Ambala City Sub- Division') && els['tblInfo'].textContent.includes(`${nSub} of ${nSub} Sub Divisions`));
  els['groupBy'].value = 'bo'; els['quick'].value = 'all'; sortState = { key: 'name', dir: 1 };

  // ---------- reporting-period cutoff ----------
  ROWS = [
    { date: '2026-10-31', id: 'OLD01', bo: 'Old BO', so: 'S1', sol: '1', kind: 'bo', 'Accounts Opened': 100 },
    { date: '2026-11-01', id: 'NEW02', bo: 'New BO', so: 'S1', sol: '2', kind: 'bo', 'Accounts Opened': 5 },
  ];
  METRICS = ['Accounts Opened']; finishIngest();
  T('rows before 2026-11-01 are excluded', ROWS.length === 1 && ROWS[0].bo === 'New BO' && DATES.join(',') === '2026-11-01');

  console.log(__failures ? `FAILED: ${__failures} assertion(s)` : 'ALL PASS');
  process.exitCode = __failures ? 1 : 0;
})().catch(e => { console.error('TEST BODY ERROR', e); process.exit(1); });
