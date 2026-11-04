let ROWS=[], METRICS=[], DATES=[], chartInstance=null, userLoaded=false;
let sortState={key:'name',dir:1};
const $=id=>document.getElementById(id);
const fmt=n=>Number(n||0).toLocaleString('en-IN');
const esc=s=>String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
let debT=null;
function onFilterInput(){clearTimeout(debT);debT=setTimeout(render,150);}
function parseDate(v){if(v instanceof Date)return v;const d=new Date(v);return isNaN(d)?null:d;}
function parseDMY(s){const m=String(s||'').match(/(\d{2})-(\d{2})-(\d{4})/);if(!m)return null;return new Date(+m[3],+m[2]-1,+m[1]);}
const iso=d=>d.toISOString().slice(0,10);
// All dates shown in the UI/exports use DD-MM-YYYY. ISO dates stay internal for
// sorting, filtering, history.csv and manifest.json.
const fmtDate=v=>{const s=String(v||'').slice(0,10);const m=s.match(/^(\d{4})-(\d{2})-(\d{2})$/);return m?`${m[3]}-${m[2]}-${m[1]}`:s;};
// BO identity is the office CODE (SOL ID / BOCODE), not the name: bo_master.js is
// generated from the circle master data sheet and holds {name, code, so,
// subDivision, division}. 2196 offices have only 2093 distinct names, so the same
// name can belong to two different BOs and names alone cannot identify an office.
// There is no fuzzy name matching: a report row either resolves to a master office
// (by code, then by name, then by name+SO) or it stays an unknown office under its
// own reported name.
const MASTER_LIST=(typeof BO_MASTER!=='undefined'&&Array.isArray(BO_MASTER))?BO_MASTER:[];
// Dashboard reporting period: earlier history snapshots (e.g. 31-10-2026) stay in
// history.csv for audit, but must never appear in KPIs, trend, table or exports.
const DATA_START='2026-11-01';
// Group-by views: 'bo' is the office list, the others roll offices up by the
// master sheet's SO / Division / Sub Division columns.
const GROUP_META={bo:{name:'BO',unit:'BOs',title:'BO-wise'},so:{name:'SO',unit:'SOs',title:'SO-wise'},
  div:{name:'Division',unit:'Divisions',title:'Division-wise'},sub:{name:'Sub Division',unit:'Sub Divisions',title:'Sub-division-wise'}};
const normName=n=>String(n||'').trim().replace(/\s+/g,' ').replace(/\s*B\.?O\.?$/i,' BO').replace(/\s+/g,' ').toLowerCase();
const soKey=s=>String(s||'').trim().toLowerCase().replace(/\./g,'').replace(/\s+/g,' ');
const normCode=c=>String(c||'').trim().toUpperCase().replace(/\s+/g,'');
const MASTER_BY_CODE=new Map(), MASTER_BY_NAME=new Map(), MASTER_BY_NAME_SO=new Map();
MASTER_LIST.forEach(o=>{
  const code=normCode(o.code);
  if(code)MASTER_BY_CODE.set(code,o);
  const k=normName(o.name);
  if(!k)return;
  // A name used by more than one office cannot identify one on its own.
  if(MASTER_BY_NAME.has(k))MASTER_BY_NAME.set(k,null);else MASTER_BY_NAME.set(k,o);
  const ks=k+'|'+soKey(o.so);
  if(!MASTER_BY_NAME_SO.has(ks))MASTER_BY_NAME_SO.set(ks,o);
});
function findOffice(name,so,id){
  const code=normCode(id);
  if(code&&MASTER_BY_CODE.has(code))return MASTER_BY_CODE.get(code);
  const k=normName(name);
  if(!k)return null;
  const byName=MASTER_BY_NAME.get(k);
  if(byName)return byName;
  return MASTER_BY_NAME_SO.get(k+'|'+soKey(so))||null;
}
// Display name: master spelling when the office is known, else the reported name;
// a master office whose sheet row has no name falls back to its code.
function officeLabel(o,name){
  if(o)return String(o.name||'').trim()||normCode(o.code)||String(name||'').trim();
  return String(name||'').trim();
}
// Grouping key for a row: resolved office code, else the reported code, else the
// reported name (uploads without a code column).
function rowId(o,name,id){return normCode(o&&o.code)||normCode(id)||normName(name)||String(name||'').trim();}
// The CEPT grid interleaves office summary rows among the BO rows ("Adhoya S.O",
// "Jamalpur SO Bhiwani", "Neemariwali PO", "Ambala G.P.O."). They are not BOs: only a
// row that resolves to a master office is one. Two signals identify an office row -
// an S.O/H.O/P.O/G.P.O marker in its name, or an all-digit code (master BO codes are
// always 5 alphanumeric characters, SO/HO codes are 8 digits).
const SO_ROW_RE=/\b(s\.?\s?o\.?|h\.?\s?o\.?|p\.?\s?o\.?|g\.?\s?p\.?\s?o\.?|head\s*office|sub\s*office)\b/i;
const OFFICE_CODE_RE=/^\d{6,}$/;
function isSORow(name,so,id){
  if(findOffice(name,so,id))return false;
  const s=String(name||'').trim();
  return SO_ROW_RE.test(s)||OFFICE_CODE_RE.test(normCode(id));
}
// In the CEPT export, office summary rows also have a BLANK 'SO Name' cell - their
// figures are SO-level, not BO-level, so they are kept as kind='so' rows and shown
// separately. (Manual uploads may legitimately lack an SO column, so there only the
// name/code pattern is used.)
function isOfficeRow(name,soVal,id){
  if(findOffice(name,soVal,id))return false;
  return isSORow(name,soVal,id)||String(soVal||'').trim()==='';
}
// One place that turns a reported (name, SO, code) triple into an office: identity,
// display name and its division/sub-division. SO summary rows live in their own
// identity space so an office named like a BO cannot merge with it.
function officeInfo(name,so,sol){
  const off=findOffice(name,so,sol);
  const soRow=isSORow(name,so,sol);
  return {off,kind:soRow?'so':'bo',
    id:soRow?'so:'+soKey(name):rowId(off,name,sol),
    label:soRow?String(name||'').trim():officeLabel(off,name),
    dv:off?off.division||'':'',sd:off?off.subDivision||'':''};
}
function setBOOptions(list){
  const all=[...new Set([...MASTER_LIST.map(o=>o.name),...list].filter(Boolean))].sort((a,b)=>a.localeCompare(b));
  $('boList').innerHTML=all.map(b=>`<option value="${esc(b)}">`).join('');
}
function setSOOptions(list){
  const all=[...new Set([...list.filter(Boolean),...MASTER_LIST.map(o=>o.so).filter(Boolean)])].sort((a,b)=>a.localeCompare(b));
  $('soList').innerHTML=all.map(s=>`<option value="${esc(s)}">`).join('');
}
// Division / sub-division options come from the master sheet, so a division with no
// accounts opened today can still be selected. Sub-division list follows the division.
function setOfficeOptions(){
  const dv=[...new Set(MASTER_LIST.map(o=>o.division).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
  const d=$('division').value;
  $('division').innerHTML=`<option value="">All divisions</option>`+dv.map(v=>`<option value="${esc(v)}">${esc(v)}</option>`).join('');
  if(d&&dv.includes(d))$('division').value=d;
  const sd=[...new Set(MASTER_LIST.filter(o=>!$('division').value||o.division===$('division').value)
    .map(o=>o.subDivision).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
  const s=$('subDivision').value;
  $('subDivision').innerHTML=`<option value="">All sub divisions</option>`+sd.map(v=>`<option value="${esc(v)}">${esc(v)}</option>`).join('');
  if(s&&sd.includes(s))$('subDivision').value=s;else $('subDivision').value='';
}
// Division / sub-division of an S.O, used to place office summary rows in the
// hierarchy. First from the master sheet's own SO column (an SO is pinned to a
// sub-division only when all its BOs agree), then from the BO rows reporting that
// SO name, so an office above a BO network the sheet does not list still resolves.
function soOfficeMap(){
  const m=new Map(), agg=new Map();
  MASTER_LIST.forEach(o=>{
    const k=soKey(o.so);if(!k)return;
    const a=agg.get(k)??(agg.set(k,{dv:new Set(),sd:new Set()}),agg.get(k));
    if(o.division)a.dv.add(o.division);
    if(o.subDivision)a.sd.add(o.subDivision);
  });
  agg.forEach((a,k)=>{
    const dv=a.dv.size===1?[...a.dv][0]:'',sd=a.sd.size===1?[...a.sd][0]:'';
    if(dv||sd)m.set(k,{dv,sd});
  });
  ROWS.forEach(r=>{
    if(r.kind==='so'||!r.so)return;
    const k=soKey(r.so);
    if(k&&!m.has(k)&&(r.dv||r.sd))m.set(k,{dv:r.dv||'',sd:r.sd||''});
  });
  return m;
}
function activeMetrics(){return METRICS;}
function isNumCell(v){return typeof v==='number'?Number.isFinite(v):(typeof v==='string'&&v.trim()!==''&&!isNaN(Number(v)));}
function findHeaderRow(matrix){
  for(let i=0;i<matrix.length;i++){
    const row=matrix[i].map(c=>String(c??'').trim());
    if(row.some(c=>/solarid|solar id|bocode|bo code|^bo$|^bo name$|branch office|^name$/i.test(c))){
      return i;
    }
  }
  return -1;
}
function extractReportDate(matrix,hIdx){
  for(let i=hIdx-1;i>=0;i--){
    const line=matrix[i].map(c=>String(c??'')).join(' ');
    const m=line.match(/Dated:\s*(\d{2}-\d{2}-\d{4})/i);
    if(m){const d=parseDMY(m[1]);if(d)return d;}
  }
  return new Date();
}
const isHeaderLike=name=>/^name$|^bo$|^bo name$|branch office|sol\s*id|bocode|^so\s*name$/i.test(name);
function updateMode(){
  // 'from' is only meaningful in range mode: 'upto' always starts at DATES[0]
  // and 'single' uses only the 'to' date.
  $('fromWrap').style.display=($('mode').value==='range')?'':'none';
}
function finishIngest(){
  ROWS=ROWS.filter(r=>r.date>=DATA_START);
  DATES=[...new Set(ROWS.map(r=>r.date))].sort();
  if(DATES.length){$('from').value=DATES[0];$('to').value=DATES[DATES.length-1];}
  $('dataAsOf').textContent=DATES.length?fmtDate(DATES[DATES.length-1]):'—';
  setBOOptions([...new Set(ROWS.filter(r=>r.kind!=='so').map(r=>r.bo))]);
  setSOOptions(ROWS.map(r=>r.so));
  setOfficeOptions();
  $('msg').textContent='';updateMode();render();
}
function ingestMatrix(matrix){
  const hIdx=findHeaderRow(matrix);
  if(hIdx<0)throw new Error('Header row (SOL ID / BOCODE / Name / BO) nahi mili. Pehli few rows: '+matrix.slice(0,3).map(r=>r.join(' | ')).join(' ;; '));
  const headers=matrix[hIdx].map((c,i)=>String(c??'').trim()||`col${i}`);
  const rows=matrix.slice(hIdx+1).filter(r=>r.some(c=>String(c??'').trim()!==''));
  if(!rows.length)throw new Error('Header ke baad koi data row nahi mili.');
  const dateIdx=headers.findIndex(h=>/date|din|tarikh/i.test(h));
  const nameIdx=headers.findIndex(h=>/^name$|^bo$|^bo name$|branch office/i.test(h));
  const solIdIdx=headers.findIndex(h=>/solarid|solar id|bocode|bo code/i.test(h));
  const soIdx=headers.findIndex(h=>/^so\s*name$|^so$|sub\s*office/i.test(h));
  const boIdx=nameIdx>=0?nameIdx:solIdIdx;
  if(boIdx<0)throw new Error('BO column (Name / SOL ID / BOCODE / BO) nahi mila. Headers: '+headers.join(', '));
  let fallbackDate=null;
  if(dateIdx<0){fallbackDate=extractReportDate(matrix,hIdx);}
  const skip=new Set([dateIdx,boIdx,solIdIdx,soIdx].filter(i=>i>=0));
  const metricIdxs=headers.map((h,i)=>i).filter(i=>!skip.has(i)&&rows.some(r=>isNumCell(r[i])));
  if(!metricIdxs.length)throw new Error('Koi numeric performance column nahi mila.');
  ROWS=rows.map(r=>{
    const name=String(r[boIdx]??'').trim();
    if(!name||/^total\b/i.test(name)||isHeaderLike(name))return null;
    const d=dateIdx>=0?parseDate(r[dateIdx]):fallbackDate;
    if(!d)return null;
    const soVal=soIdx>=0?String(r[soIdx]??'').trim():'';
    const solRaw=solIdIdx>=0?String(r[solIdIdx]??'').trim():'';
    const inf=officeInfo(name,soVal,solRaw);
    const o={date:iso(d),id:inf.id,bo:inf.label,so:inf.kind==='so'?name:soVal,sol:solRaw,
      kind:inf.kind,dv:inf.dv,sd:inf.sd};
    metricIdxs.forEach(i=>o[headers[i]]=Number(r[i])||0);
    return o;
  }).filter(Boolean);
  METRICS=metricIdxs.map(i=>headers[i]);
  finishIngest();
}
function handleFile(f){
  userLoaded=true; // guard: a pending auto-load must not overwrite a manual upload
  const r=new FileReader();
  r.onload=e=>{try{
    const wb=XLSX.read(e.target.result,{type:'array',cellDates:true});
    const ws=wb.Sheets[wb.SheetNames[0]];
    const matrix=XLSX.utils.sheet_to_json(ws,{header:1,defval:null,raw:true});
    if(!matrix.length)throw new Error('File me koi data nahi mili.');
    ingestMatrix(matrix);
  }catch(err){$('msg').textContent=err.message}};
  r.readAsArrayBuffer(f);
}
// The automation writes one merged workbook per KPI plus a manifest.json listing
// them (with the report date, since merged files no longer keep the 'Dated:'
// title row). data/history.csv holds the daily snapshots that make the trend
// chart meaningful. Load history first, then overlay the latest exports.
const KPI_SOURCES=[
  {file:'data/exports/accounts_opened_merged.xlsx',label:'Accounts Opened'}
];
function parseCsvText(t){
  const rows=[];let cur=[],f='',q=false;
  for(let i=0;i<t.length;i++){
    const c=t[i];
    if(q){
      if(c==='"'){if(t[i+1]==='"'){f+='"';i++;}else q=false;}
      else f+=c;
    }else if(c==='"')q=true;
    else if(c===','){cur.push(f);f='';}
    else if(c==='\n'||c==='\r'){
      if(c==='\r'&&t[i+1]==='\n')i++;
      cur.push(f);f='';
      if(cur.length>1||cur[0]!=='')rows.push(cur);
      cur=[];
    }else f+=c;
  }
  if(f!==''||cur.length){cur.push(f);rows.push(cur);}
  return rows;
}
async function loadHistory(){
  // Returns {rows, metrics} WITHOUT touching globals - the caller applies them
  // only after the userLoaded race guard, so a pending auto-load can never
  // clobber a manual upload.
  const res=await fetch('data/history.csv',{cache:'no-store'});
  if(!res.ok)throw new Error(`${res.status}`);
  const rows=parseCsvText(await res.text());
  if(rows.length<2)throw new Error('history khali hai');
  const h=rows[0].map(s=>String(s).trim().toLowerCase());
  const iDate=h.indexOf('date'),iBo=h.indexOf('bo'),iSo=h.indexOf('so'),iSol=h.indexOf('sol'),iKind=h.indexOf('kind'),iKpi=h.indexOf('kpi'),iVal=h.indexOf('value');
  if(iDate<0||iBo<0||iKpi<0||iVal<0)throw new Error('history headers galat hain: '+rows[0].join(','));
  const byKey=new Map();const metrics=new Set();
  rows.slice(1).forEach(r=>{
    const d=String(r[iDate]||'').trim(),bo=String(r[iBo]||'').trim(),kpi=String(r[iKpi]||'').trim();
    if(!d||!bo||!kpi)return;
    const v=Number(String(r[iVal]||'').replace(/,/g,''))||0;
    const soVal=iSo>=0?String(r[iSo]||'').trim():'';
    const solRaw=iSol>=0?String(r[iSol]||'').trim():'';
    const inf=officeInfo(bo,soVal,solRaw);
    // The master sheet decides: a row whose code (or unique name) is a master BO is a
    // BO even when the snapshot filed it as an office summary row. Rows the sheet does
    // not know keep whatever kind the snapshot recorded.
    const isSo=inf.kind==='so'||((iKind>=0&&String(r[iKind]||'').trim().toLowerCase()==='so')&&!inf.off);
    const id=isSo?'so:'+soKey(bo):inf.id;
    const key=d+'|'+id;
    let row=byKey.get(key);
    if(!row){row={date:d,id,bo:isSo?bo:inf.label,so:isSo?bo:soVal,sol:solRaw,kind:isSo?'so':'bo',dv:inf.dv,sd:inf.sd};byKey.set(key,row);}
    row[kpi]=v;
    if(!row.so&&isSo)row.so=bo;
    if(!row.so&&soVal)row.so=soVal;
    if(!row.sol&&solRaw)row.sol=solRaw;
    metrics.add(kpi);
  });
  if(!byKey.size)throw new Error('history me koi row nahi mili');
  return {rows:[...byKey.values()], metrics:[...metrics]};
}
function parseKpiWorkbook(matrix,label){
  const hIdx=findHeaderRow(matrix);
  if(hIdx<0)throw new Error(label+': header row (SOL ID / BOCODE / Name) nahi mili.');
  const headers=matrix[hIdx].map((c,i)=>String(c??'').trim()||`col${i}`);
  const rows=matrix.slice(hIdx+1).filter(r=>r&&r.some(c=>String(c??'').trim()!==''));
  if(!rows.length)throw new Error(label+': header ke baad koi data row nahi mili.');
  const nameIdx=headers.findIndex(h=>/^name$|^bo$|^bo name$|branch office/i.test(h));
  const solIdIdx=headers.findIndex(h=>/solarid|solar id|bocode|bo code/i.test(h));
  const soIdx=headers.findIndex(h=>/^so\s*name$|^so$|sub\s*office/i.test(h));
  const dateIdx=headers.findIndex(h=>/date|din|tarikh/i.test(h));
  const boIdx=nameIdx>=0?nameIdx:solIdIdx;
  if(boIdx<0)throw new Error(label+': BO column (Name / SOL ID / BOCODE) nahi mila. Headers: '+headers.join(', '));
  const isIdCol=h=>/solarid|solar id|bocode|bo code|^so name$|^so$|sub\s*office|^sr\b|^s\.no|^sl\.?no/i.test(h);
  // Confirmed against the live CBSTrans3 grid: the export has no 'Total' column,
  // only the product columns (MIS/PPFGP/SSA/RD/... up to MSSC) plus a 'Total:' footer
  // row, so a BO's KPI value is the sum of its product cells.
  const valueIdxs=headers
    .map((h,i)=>({h,i}))
    .filter(({h,i})=>
      i!==boIdx && i!==solIdIdx && i!==dateIdx && i!==soIdx &&
      !/^__source$/i.test(h) && !isIdCol(h) &&
      rows.some(r=>isNumCell(r[i]))
    )
    .map(({i})=>i);
  if(!valueIdxs.length)throw new Error(label+': koi numeric column nahi mila. Headers: '+headers.join(', '));
  const bos=new Map(), offices=new Map();
  rows.forEach(r=>{
    const name=String(r[boIdx]??'').trim();
    // The exported sheet carries a 'Total:' footer row and office summary rows
    // (S.O/H.O/G.P.O - recognisable by a blank SO Name cell); repeated headers are not rows.
    if(!name||/^total\b/i.test(name)||isHeaderLike(name))return;
    const soVal=soIdx>=0?String(r[soIdx]??'').trim():'';
    const solRaw=solIdIdx>=0?String(r[solIdIdx]??'').trim():'';
    const v=valueIdxs.reduce((s,i)=>s+(Number(r[i])||0),0);
    if(isOfficeRow(name,soVal,solRaw)){
      // SO-level figure from the office's own summary row - kept SEPARATE from BO sums.
      const k='so:'+soKey(name);
      const e=offices.get(k)??(offices.set(k,{id:k,name,so:name,sol:'',value:0}),offices.get(k));
      e.value+=v;
      if(solIdIdx>=0&&!e.sol)e.sol=solRaw;
      return;
    }
    const inf=officeInfo(name,soVal,solRaw);
    const e=bos.get(inf.id)??(bos.set(inf.id,{id:inf.id,name:inf.label,so:soVal,sol:solRaw,dv:inf.dv,sd:inf.sd,value:0}),bos.get(inf.id));
    e.value+=v;
    if(!e.so&&soVal)e.so=soVal;
    if(!e.sol&&solRaw)e.sol=solRaw;
  });
  return {label,date:extractReportDate(matrix,hIdx),bos:[...bos.values()],offices:[...offices.values()]};
}
function overlayKpiParts(parts){
  // Latest exports replace any history rows of the same date, keyed by date+office,
  // so a re-run for one date never double-counts in the trend chart.
  const key=(d,b)=>d+'|'+b;
  const byKey=new Map(ROWS.map(r=>[key(r.date,r.id),r]));
  parts.forEach(p=>{
    const d=iso(p.date||new Date());
    (p.bos||[]).forEach(e=>{
      const k=key(d,e.id);
      let row=byKey.get(k);
      if(!row){row={date:d,id:e.id,bo:e.name,so:e.so||'',sol:e.sol||'',kind:'bo',dv:e.dv||'',sd:e.sd||''};byKey.set(k,row);}
      row.bo=row.bo||e.name;row[p.label]=Number(e.value)||0;
      if(!row.so&&e.so)row.so=e.so;
      if(!row.sol&&e.sol)row.sol=e.sol;
      if(!row.dv&&e.dv)row.dv=e.dv;
      if(!row.sd&&e.sd)row.sd=e.sd;
    });
    // SO-level summary rows from the export: kept as kind='so' so the SO-wise view
    // can show the SO's own reported figure separately from the sum of its BOs.
    (p.offices||[]).forEach(e=>{
      const k=key(d,e.id);
      let row=byKey.get(k);
      if(!row){row={date:d,id:e.id,bo:e.name,so:e.so||e.name,sol:e.sol||'',kind:'so',dv:'',sd:''};byKey.set(k,row);}
      row[p.label]=Number(e.value)||0;
    });
  });
  ROWS=[...byKey.values()];
  METRICS=[...new Set([...METRICS,...parts.map(p=>p.label)])];
}
async function autoLoad(){
  let hist=null;
  try{hist=await loadHistory();}catch(e){console.warn('history load failed:',e);}
  if(userLoaded){$('autoStatus').textContent='Manual file loaded — auto data skipped.';return;}
  let sources=KPI_SOURCES;
  try{
    const res=await fetch('data/exports/manifest.json',{cache:'no-store'});
    if(res.ok){
      const m=await res.json();
      if(Array.isArray(m)&&m.length)sources=m.map(s=>({file:s.file,label:s.label,date:s.date}));
    }
  }catch(e){/* manifest missing -> fall back to hardcoded KPI_SOURCES */}
  const parts=[],missing=[];
  for(const src of sources){
    if(userLoaded)break;
    try{
      const res=await fetch(src.file,{cache:'no-store'});
      if(!res.ok)throw new Error(`${res.status} ${res.statusText||''}`.trim());
      const buf=await res.arrayBuffer();
      const wb=XLSX.read(buf,{type:'array',cellDates:true});
      const matrix=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:null,raw:true});
      if(matrix.length<2)throw new Error('sheet khali hai');
      const part=parseKpiWorkbook(matrix,src.label);
      if(src.date){const d=parseDate(src.date);if(d)part.date=d;}
      parts.push(part);
    }catch(e){missing.push(src.label);}
  }
  if(userLoaded){$('autoStatus').textContent='Manual file loaded — auto data skipped.';return;}
  if(!hist&&!parts.length){
    $('autoStatus').textContent='Auto data abhi available nahi hai — manual upload use karein.';
    return;
  }
  if(hist){ROWS=hist.rows;METRICS=hist.metrics;}
  if(parts.length)overlayKpiParts(parts);
  finishIngest();
  const soCount=new Set(ROWS.map(r=>r.so).filter(Boolean)).size;
  const reportedIds=new Set(ROWS.filter(r=>r.kind!=='so').map(r=>r.id));
  const missingMaster=MASTER_LIST.filter(o=>!reportedIds.has(normCode(o.code)||normName(o.name))).length;
  const bits=[];
  if(hist)bits.push(`history ${DATES.length} day(s)`);
  if(parts.length)bits.push(`${parts.length}/${sources.length} KPI export(s) @ ${fmtDate(DATES[DATES.length-1])}`);
  if(missing.length)bits.push(`missing: ${missing.join(', ')}`);
  bits.push(`${ROWS.length} rows, ${reportedIds.size} BOs reported, ${missingMaster} master BOs at 0, ${soCount} SOs`);
  $('autoStatus').textContent='Auto-loaded '+bits.join(' • ')+'.';
}
function filtered(){
  const mode=$('mode').value,from=$('from').value,to=$('to').value;
  const bo=$('bo').value.trim().toLowerCase(),so=$('so').value.trim().toLowerCase();
  const dv=$('division').value,sd=$('subDivision').value;
  let a=from,b=to;
  if(mode==='single'){a=b=to}
  if(mode==='upto'){a=DATES[0]}
  const offMap=(dv||sd)?soOfficeMap():null;
  return ROWS.filter(r=>{
    if(a&&r.date<a)return false;
    if(b&&r.date>b)return false;
    if(bo&&!r.bo.toLowerCase().includes(bo))return false;
    if(so&&!String(r.so||'').toLowerCase().includes(so))return false;
    if(dv||sd){
      const o=r.kind==='so'?(offMap.get(soKey(r.so||r.bo))||{dv:'',sd:''}):r;
      if(dv&&o.dv!==dv)return false;
      if(sd&&o.sd!==sd)return false;
    }
    return true;
  });
}
function sortBy(key){
  if(sortState.key===key)sortState.dir*=-1;
  else sortState={key,dir:1};
  render();
}
// Office metadata (SO / code / division) from every loaded row, not just the
// selected date range, keyed by office id (= master BO code) so a BO that opened
// no accounts in the window still shows its own SO, division and sub-division.
function officeMetaMap(){
  const meta=new Map();
  ROWS.forEach(r=>{
    if(r.kind==='so')return;
    let m=meta.get(r.id);
    if(!m){m={name:r.bo,so:'',sol:'',dv:'',sd:''};meta.set(r.id,m);}
    if(!m.name&&r.bo)m.name=r.bo;
    if(!m.so&&r.so)m.so=r.so;
    if(!m.sol&&r.sol)m.sol=r.sol;
    if(!m.dv&&r.dv)m.dv=r.dv;
    if(!m.sd&&r.sd)m.sd=r.sd;
  });
  return meta;
}
// BO-wise view = offices present in the report UNION every office from
// bo_master.js (2,196 in the Haryana Circle sheet). An office with no accounts
// opened in the selected window stays in the table with zeros instead of
// disappearing. Ids are codes, so two offices that share a name stay separate rows.
function visibleOfficeMap(boRows){
  const meta=officeMetaMap(), view=new Map();
  boRows.forEach(r=>{
    let m=view.get(r.id);
    if(!m){m={id:r.id,name:r.bo,so:r.so||'',sol:r.sol||'',dv:r.dv||'',sd:r.sd||''};view.set(r.id,m);}
    else{
      if(!m.name&&r.bo)m.name=r.bo;
      if(!m.so&&r.so)m.so=r.so;
      if(!m.sol&&r.sol)m.sol=r.sol;
      if(!m.dv&&r.dv)m.dv=r.dv;
      if(!m.sd&&r.sd)m.sd=r.sd;
    }
  });
  const boF=$('bo').value.trim().toLowerCase(),soF=$('so').value.trim().toLowerCase();
  const dvF=$('division').value,sdF=$('subDivision').value;
  MASTER_LIST.forEach(o=>{
    const id=normCode(o.code)||normName(o.name);
    if(!id||view.has(id))return;
    const name=String(o.name||'').trim();
    if(boF&&!(name.toLowerCase().includes(boF)||id.toLowerCase().includes(boF)))return;
    // With an SO filter, only add master offices whose own SO matches.
    if(soF&&!String(o.so||'').toLowerCase().includes(soF))return;
    if(dvF&&String(o.division||'')!==dvF)return;
    if(sdF&&String(o.subDivision||'')!==sdF)return;
    const known=meta.get(id)||{name:'',so:'',sol:'',dv:'',sd:''};
    view.set(id,{id,name:officeLabel(o,name)||known.name,so:o.so||known.so,sol:o.code||known.sol,
      dv:o.division||known.dv,sd:o.subDivision||known.sd});
  });
  return view;
}
function render(){
  if(!ROWS.length)return;
  const rows=filtered(), act=activeMetrics();
  const gby=$('groupBy').value||'bo';
  const gm=GROUP_META[gby]||GROUP_META.bo;
  // SO summary rows (kind='so') are office-level figures: they never join the BO
  // totals, only the (SO) columns of the SO / division / sub-division views.
  const boRows=rows.filter(r=>r.kind!=='so');
  const soRows=rows.filter(r=>r.kind==='so');
  const boView=visibleOfficeMap(boRows);
  const wide=gby!=='bo';
  const offMap=soOfficeMap();
  const inh=r=>r.kind==='so'?(offMap.get(soKey(r.so||r.bo))||{dv:'',sd:''}):r;
  const days=[...new Set(boRows.map(r=>r.date))].sort();
  // Per-day sums drive the day-over-day delta (▲/▼ vs previous day) on metric KPI cards.
  const daySum={};
  boRows.forEach(r=>{daySum[r.date]??={};act.forEach(m=>daySum[r.date][m]=(daySum[r.date][m]||0)+(Number(r[m])||0));});
  const lastD=days[days.length-1], prevD=days.length>1?days[days.length-2]:null;
  const deltaOf=m=>prevD?((daySum[lastD]?.[m]||0)-(daySum[prevD]?.[m]||0)):null;
  const kpis=[['BOs',boView.size],
              ['Divisions',new Set([...boView.values()].map(m=>m.dv).filter(Boolean)).size],
              ['Sub Divisions',new Set([...boView.values()].map(m=>m.sd).filter(Boolean)).size],
              ['SOs',new Set(rows.map(r=>r.so).filter(Boolean)).size],
              ...act.map(m=>[m,boRows.reduce((s,r)=>s+(Number(r[m])||0),0)])];
  $('kpis').innerHTML=kpis.map(([l,v])=>{
    const d=act.includes(l)?deltaOf(l):null;
    const sub=d===null?'':`<div class="d ${d>=0?'good':'bad'}">${d>=0?'▲':'▼'} ${fmt(Math.abs(d))} vs ${esc(fmtDate(prevD))}</div>`;
    return `<div class="kpi"><div class="l">${esc(l)}</div><div class="v">${fmt(v)}</div>${sub}</div>`;
  }).join('');
  // Aggregate into groups: BO-wise = one row per office (id = BO code), the others
  // = one row per SO / Division / Sub Division with the office rows summed up.
  const keyOf=r=>gby==='bo'?r.id
    :gby==='so'?(soKey(r.so)||'(no so)')
    :gby==='div'?(inh(r).dv||'(no division)')
    :(inh(r).sd||'(no sub division)');
  const nameOf=r=>{
    if(gby==='bo')return r.bo||r.id;
    // BO rows are named by their SO cell; office summary rows by their own name.
    if(gby==='so')return String((r.kind==='so'?r.bo:r.so)||'').trim()||'(no SO)';
    if(gby==='div')return inh(r).dv||'(no division)';
    return inh(r).sd||'(no sub division)';
  };
  const groups=new Map();
  const getG=(k,name)=>{
    let g=groups.get(k);
    if(!g){g={key:k,name:name||k,bos:new Set(),so:'',dv:'',sd:'',vals:{},total:0,soVals:null};groups.set(k,g);}
    else if(name&&(!g.name||g.name===k))g.name=name;
    return g;
  };
  // Seed the BO-wise view with every visible office (report UNION master sheet) so
  // offices that opened no accounts in this window still appear, with zeros.
  if(gby==='bo')boView.forEach(m=>{
    const g=getG(m.id,m.name);
    g.so=m.so||'';g.dv=m.dv||'';g.sd=m.sd||'';g.bos.add(m.id);
  });
  // Division / sub-division views seed every visible office's group as well, so a
  // division that opened nothing in the window still appears (and can be compared).
  if(gby==='div'||gby==='sub')boView.forEach(m=>{
    const k=gby==='div'?(m.dv||'(no division)'):(m.sd||'(no sub division)');
    if(k!=='(no division)'&&k!=='(no sub division)')getG(k,k);
  });
  (wide?rows:boRows).forEach(r=>{
    const g=getG(keyOf(r),nameOf(r));
    if(r.kind==='so'){
      g.soVals=g.soVals||{};
      act.forEach(m=>g.soVals[m]=(g.soVals[m]||0)+(Number(r[m])||0));
      if(!g.name||g.name==='(no SO)')g.name=String(r.bo||'').trim()||g.name;
      return;
    }
    g.bos.add(r.id);
    if(!g.so&&r.so)g.so=r.so;
    if(!g.dv&&r.dv)g.dv=r.dv;
    if(!g.sd&&r.sd)g.sd=r.sd;
    act.forEach(m=>{const v=Number(r[m])||0;g.vals[m]=(g.vals[m]||0)+v;g.total+=v;});
  });
  let list=[...groups.values()].map(g=>({name:g.name,so:g.so,dv:g.dv,sd:g.sd,bos:g.bos.size,
    vals:g.vals,total:g.total,soVals:g.soVals}));
  const q=$('quick').value;
  let sk=sortState.key, sd=sortState.dir;
  if(q==='top'){sk='total';sd=-1;}
  else if(q==='bottom'){sk='total';sd=1;}
  if(sk==='bos'&&gby==='bo')sk='name';
  const soValOf=(g,base)=>!g.soVals?0:(base==='total'?Object.values(g.soVals).reduce((x,y)=>x+Number(y||0),0):(g.soVals[base]||0));
  list.sort((a,b)=>{
    let va,vb;
    if(sk==='bos'){va=a.bos;vb=b.bos;}
    else if(sk==='total'){va=a.total;vb=b.total;}
    else if(sk.indexOf('so:')===0){const base=sk.slice(3);va=soValOf(a,base);vb=soValOf(b,base);}
    else if(act.includes(sk)){va=a.vals[sk]||0;vb=b.vals[sk]||0;}
    else{va=String(a[sk]??a.name??'');vb=String(b[sk]??b.name??'');}
    return (typeof va==='string')?sd*va.localeCompare(vb):sd*(va-vb);
  });
  const totalGroups=list.length;
  if(q==='top'||q==='bottom')list=list.slice(0,10);
  const arrow=key=>(sk===key)?(sd>0?' ▲':' ▼'):'';
  const th=(key,label)=>`<th class="sortable" data-k="${esc(key)}">${esc(label)}${arrow(key)}</th>`;
  // Wide views keep the two figures separate: "<metric> (ΣBO)" = sum of the offices
  // under the group, "<metric> (SO)" = the office's own reported summary figure
  // ('—' when the report has no such row).
  const cols=gby==='bo'
    ?[['name','BO'],['so','SO'],['dv','Division'],['sd','Sub Division']]
    :[['name',gm.name],['bos','BOs']];
  const mHead=wide
    ?act.map(m=>th(m,m+' (ΣBO)')+th('so:'+m,m+' (SO)')).join('')+th('total','Total (ΣBO)')+th('so:total','Total (SO)')
    :act.map(m=>th(m,m)).join('')+th('total','Total');
  const head='<thead><tr>'+cols.map(([k,l])=>th(k,l)).join('')+mHead+'</tr></thead>';
  const body=list.map(g=>'<tr>'+cols.map(([k])=>k==='bos'?`<td>${g.bos}</td>`:`<td>${esc(g[k]??'')}</td>`).join('')+
    (wide
      ?act.map(m=>`<td><b>${fmt(g.vals[m]||0)}</b></td><td>${g.soVals?fmt(g.soVals[m]||0):'—'}</td>`).join('')+
       `<td><b>${fmt(g.total)}</b></td><td>${g.soVals?fmt(soValOf(g,'total')):'—'}</td>`
      :act.map(m=>`<td><b>${fmt(g.vals[m]||0)}</b></td>`).join('')+`<td><b>${fmt(g.total)}</b></td>`)+
    '</tr>').join('');
  const boMetricTotal=m=>boRows.reduce((s,r)=>s+(Number(r[m])||0),0);
  const soMetricTotal=m=>soRows.reduce((s,r)=>s+(Number(r[m])||0),0);
  // Sum of every metric of a row, then over rows (the footer's 'Total' column).
  const rowTotal=r=>act.reduce((ss,m)=>ss+(Number(r[m])||0),0);
  const boGrandTotal=boRows.reduce((s,r)=>s+rowTotal(r),0);
  const soGrandTotal=soRows.reduce((s,r)=>s+rowTotal(r),0);
  const foot='<tfoot><tr><td><b>Total</b></td>'+cols.slice(1).map(()=>'<td></td>').join('')+
    (wide
      ?act.map(m=>`<td><b>${fmt(boMetricTotal(m))}</b></td><td><b>${fmt(soMetricTotal(m))}</b></td>`).join('')+
       `<td><b>${fmt(boGrandTotal)}</b></td>`+
       `<td><b>${fmt(soGrandTotal)}</b></td>`
      :act.map(m=>`<td><b>${fmt(boMetricTotal(m))}</b></td>`).join('')+
       `<td><b>${fmt(boGrandTotal)}</b></td>`)+
    '</tr></tfoot>';
  $('tbl').innerHTML=head+'<tbody>'+body+'</tbody>'+foot;
  $('tblTitle').textContent=gm.title+' cumulative performance';
  const skLabel=cols.some(([k,l])=>k===sk)?cols.find(([k,l])=>k===sk)[1]
    :(sk.indexOf('so:')===0?sk.slice(3)+' (SO)':sk);
  $('tblInfo').textContent=`Showing ${list.length} of ${totalGroups} ${gm.unit}`+
    (q!=='all'?` (${q} 10)`:'')+` • ${boRows.length} BO rows${soRows.length?` + ${soRows.length} SO rows`:''} • sorted by ${skLabel} ${sd>0?'▲':'▼'}`;
  $('printHead').innerHTML=`<h1 style="font-size:16px">BO Performance Dashboard</h1>`+
    `<div style="font-size:12px">Mode: ${esc($('mode').value)} • From: ${esc($('from').value?fmtDate($('from').value):'—')} • To: ${esc($('to').value?fmtDate($('to').value):'—')}`+
    ` • BO filter: ${esc($('bo').value||'All')} • SO filter: ${esc($('so').value||'All')}`+
    ` • Division: ${esc($('division').value||'All')} • Sub Division: ${esc($('subDivision').value||'All')}`+
    ` • Group by: ${esc(gm.title)} • Data as of ${esc($('dataAsOf').textContent)}</div>`;
  let run={};const series={};act.forEach(m=>series[m]=[]);
  days.forEach(d=>{
    boRows.filter(r=>r.date===d).forEach(r=>act.forEach(m=>run[m]=(run[m]||0)+(Number(r[m])||0)));
    act.forEach(m=>series[m].push(run[m]||0));
  });
  const data={labels:days.map(fmtDate),datasets:act.map(m=>({label:m,data:series[m],tension:.3}))};
  const opts={responsive:true,
    plugins:{legend:{labels:{color:'#e2e8f0'}},tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${fmt(c.parsed.y)}`}}},
    scales:{x:{ticks:{color:'#94a3b8',autoSkip:true,maxRotation:45}},y:{beginAtZero:true,ticks:{color:'#94a3b8',callback:v=>fmt(v)}}}};
  if(chartInstance){chartInstance.data=data;chartInstance.options=opts;chartInstance.update();}
  else chartInstance=new Chart($('chart'),{type:'line',data,options:opts});
}
function exportExcel(){
  if(!ROWS.length)return alert('Pehle data load karein.');
  const act=activeMetrics();
  const all=filtered();
  const rows=all.filter(r=>r.kind!=='so'); // office-level pivot from BO rows only
  const offMap=soOfficeMap();
  // SO summary rows carry no division of their own; inherit it from their BO rows.
  const dvOf=r=>r.kind==='so'?((offMap.get(soKey(r.so||r.bo))||{}).dv||''):r.dv||'';
  const sdOf=r=>r.kind==='so'?((offMap.get(soKey(r.so||r.bo))||{}).sd||''):r.sd||'';
  const out=all.map(r=>{
    const o={Date:fmtDate(r.date),Kind:r.kind||'bo',BO:r.bo,SO:r.so||'','SOL ID / BOCODE':r.sol||'',
      Division:dvOf(r),'Sub Division':sdOf(r)};
    act.forEach(m=>o[m]=Number(r[m])||0);
    return o;
  });
  // BO-wise pivot sheet mirroring the on-screen table: every office from the master
  // sheet appears (keyed by BO code), with zeros when it opened no accounts.
  const byId=new Map();
  visibleOfficeMap(rows).forEach(m=>{
    byId.set(m.id,{BO:m.name,'BO ID':m.id,SO:m.so||'',Division:m.dv||'','Sub Division':m.sd||'',_t:0});
  });
  rows.forEach(r=>{
    const o=byId.get(r.id)??byId.set(r.id,{BO:r.bo,'BO ID':r.id,SO:r.so||'',Division:r.dv||'',
      'Sub Division':r.sd||'',_t:0}).get(r.id);
    if(!o.BO&&r.bo)o.BO=r.bo;
    if(!o.SO&&r.so)o.SO=r.so;
    if(!o.Division&&r.dv)o.Division=r.dv;
    if(!o['Sub Division']&&r.sd)o['Sub Division']=r.sd;
    act.forEach(m=>{const v=Number(r[m])||0;o[m]=(o[m]||0)+v;o._t+=v;});
  });
  const nameOf=o=>String(o.BO||o['BO ID']||'').trim();
  const piv=[...byId.values()].sort((a,b)=>nameOf(a).localeCompare(nameOf(b))||String(a['BO ID']).localeCompare(String(b['BO ID'])))
    .map(o=>{
      const p={BO:nameOf(o),'BO ID':o['BO ID'],SO:o.SO,Division:o.Division,'Sub Division':o['Sub Division']};
      act.forEach(m=>p[m]=o[m]||0);
      p.Total=o._t;
      return p;
    });
  const grand=m=>rows.reduce((s,r)=>s+(Number(r[m])||0),0);
  const grandTotal=rows.reduce((s,r)=>s+act.reduce((ss,m)=>ss+(Number(r[m])||0),0),0);
  const tot={BO:'Total','BO ID':'','SO':'',Division:'','Sub Division':''};
  act.forEach(m=>tot[m]=grand(m));
  tot.Total=grandTotal;
  piv.push(tot);
  // Roll-ups: one sheet per grouping level, offices summed within each group.
  const rollUp=keyOf=>{
    const g=new Map();
    rows.forEach(r=>{
      const k=keyOf(r)||'(unassigned)';
      let o=g.get(k);
      if(!o){o={name:k,bos:new Set(),_t:0};g.set(k,o);}
      o.bos.add(r.id);
      act.forEach(m=>{const v=Number(r[m])||0;o[m]=(o[m]||0)+v;o._t+=v;});
    });
    const list=[...g.values()].sort((a,b)=>b._t-a._t).map(o=>{
      const p={Group:o.name,BOs:o.bos.size};
      act.forEach(m=>p[m]=o[m]||0);
      p.Total=o._t;
      return p;
    });
    const t={Group:'Total',BOs:new Set(rows.map(r=>r.id)).size};
    act.forEach(m=>t[m]=grand(m));
    t.Total=grandTotal;
    list.push(t);
    return list;
  };
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(out),'Filtered Data');
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(piv),'BO-wise');
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rollUp(r=>dvOf(r))),'Division-wise');
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rollUp(r=>sdOf(r))),'SubDivision-wise');
  XLSX.writeFile(wb,`BO_dashboard_${fmtDate($('to').value)||'export'}.xlsx`);
}
const drop=$('drop');drop.onclick=()=>$('file').click();
$('file').onchange=e=>e.target.files[0]&&handleFile(e.target.files[0]);
drop.ondragover=e=>{e.preventDefault();drop.classList.add('drag')};
drop.ondragleave=()=>drop.classList.remove('drag');
drop.ondrop=e=>{e.preventDefault();drop.classList.remove('drag');e.dataTransfer.files[0]&&handleFile(e.dataTransfer.files[0])};
// All control wiring lives here via addEventListener: the CSP (script-src without
// 'unsafe-inline') refuses inline on* attributes in the HTML, so handlers must be
// attached from this external script.
$('mode').addEventListener('change',()=>{updateMode();render();});
$('from').addEventListener('input',onFilterInput);
$('from').addEventListener('change',onFilterInput);
$('to').addEventListener('input',onFilterInput);
$('to').addEventListener('change',onFilterInput);
$('bo').addEventListener('input',onFilterInput);
$('so').addEventListener('input',onFilterInput);
// Division drives the sub-division list, so changing it re-fills the second filter
// (setOfficeOptions keeps a still-valid selection) before re-rendering.
$('division').addEventListener('change',()=>{setOfficeOptions();render();});
$('subDivision').addEventListener('change',onFilterInput);
$('groupBy').addEventListener('change',render);
$('quick').addEventListener('change',render);
$('btnExcel').addEventListener('click',exportExcel);
$('btnPdf').addEventListener('click',()=>window.print());
// Sortable table headers: one delegated listener instead of per-th inline onclick.
$('tbl').addEventListener('click',e=>{
  const th=e.target.closest('th.sortable');
  if(th)sortBy(th.dataset.k);
});
updateMode();
autoLoad();
