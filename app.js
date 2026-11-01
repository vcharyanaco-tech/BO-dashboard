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
const MASTER_LIST=(typeof BO_MASTER!=='undefined')?BO_MASTER:[];
function normalizeBO(name){
  return String(name||'').trim().replace(/\s+/g,' ')
    .replace(/\s*B\.?O\.?$/i,' BO').replace(/\s+/g,' ').trim().toLowerCase();
}
const MASTER_MAP={};
MASTER_LIST.forEach(b=>MASTER_MAP[normalizeBO(b)]=b);
// Pre-built cleaned-key index so fuzzy matching is O(1) per name instead of a
// full scan over the master list for every row.
const CLEAN_MAP={};
for(const [mk,mv] of Object.entries(MASTER_MAP)){
  const k=mk.replace(/[^a-z0-9]/g,'');
  if(!(k in CLEAN_MAP))CLEAN_MAP[k]=mv;
}
function matchBO(name){
  const key=normalizeBO(name);
  if(MASTER_MAP[key])return MASTER_MAP[key];
  const k=key.replace(/[^a-z0-9]/g,'');
  if(CLEAN_MAP[k])return CLEAN_MAP[k];
  return String(name||'').trim();
}
// The CEPT grid interleaves S.O summary rows ("Adhoya S.O", "Achina SO") among the
// BO rows. They are not BOs: drop them unless the name is a known master BO.
const SO_ROW_RE=/\bs\.?\s?o\.?\s*$/i;
function isSORow(name){
  const n=String(name||'').trim();
  return SO_ROW_RE.test(n)&&!MASTER_MAP[normalizeBO(n)];
}
// In the CEPT export, office summary rows (S.O / H.O / G.P.O like 'Ambala G.P.O.')
// have a BLANK 'SO Name' cell - their figures are SO-level, not BO-level, so they are
// kept as kind='so' rows and shown separately. (Manual uploads may legitimately lack
// an SO column, so there only the name pattern is used.) A name present in the BO
// master list is always a real BO and is never reclassified.
function isOfficeRow(name,soVal){
  const n=String(name||'').trim();
  if(MASTER_MAP[normalizeBO(n)])return false;
  return SO_ROW_RE.test(n)||String(soVal||'').trim()==='';
}
function soKey(s){return String(s||'').trim().toLowerCase().replace(/\.+$/,'').replace(/\s+/g,' ');}
function setBOOptions(list){
  const all=[...new Set([...MASTER_LIST,...list].filter(Boolean))].sort((a,b)=>a.localeCompare(b));
  $('boList').innerHTML=all.map(b=>`<option value="${esc(b)}">`).join('');
}
function setSOOptions(list){
  const all=[...new Set(list.filter(Boolean))].sort((a,b)=>a.localeCompare(b));
  $('soList').innerHTML=all.map(s=>`<option value="${esc(s)}">`).join('');
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
  DATES=[...new Set(ROWS.map(r=>r.date))].sort();
  if(DATES.length){$('from').value=DATES[0];$('to').value=DATES[DATES.length-1];}
  $('dataAsOf').textContent=DATES.length?DATES[DATES.length-1]:'—';
  setBOOptions([...new Set(ROWS.filter(r=>r.kind!=='so').map(r=>r.bo))]);
  setSOOptions(ROWS.map(r=>r.so));
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
    const soVal=soIdx>=0?String(r[soIdx]??'').trim():'';const soRow=isSORow(name);
    const o={date:iso(d),bo:soRow?name:matchBO(name),so:soRow?name:soVal,sol:solIdIdx>=0?String(r[solIdIdx]??'').trim():'',kind:soRow?'so':'bo'};
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
    const kind=(iKind>=0&&String(r[iKind]||'').trim().toLowerCase()==='so')?'so':'bo';
    const key=d+'|'+bo+'|'+kind;
    let row=byKey.get(key);
    if(!row){row={date:d,bo:kind==='so'?bo:matchBO(bo),so:'',sol:'',kind};byKey.set(key,row);}
    row[kpi]=v;
    if(kind==='so'){row.bo=bo;row.so=bo;}
    if(iSo>=0&&!row.so)row.so=String(r[iSo]||'').trim();
    if(iSol>=0&&!row.sol)row.sol=String(r[iSol]||'').trim();
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
  const totals={},so={},sol={},soTotals={};
  rows.forEach(r=>{
    const name=String(r[boIdx]??'').trim();
    // The exported sheet carries a 'Total:' footer row and office summary rows
    // (S.O/H.O/G.P.O - recognisable by a blank SO Name cell); repeated headers are not rows.
    if(!name||/^total\b/i.test(name)||isHeaderLike(name))return;
    const soVal=soIdx>=0?String(r[soIdx]??'').trim():'';
    const v=valueIdxs.reduce((s,i)=>s+(Number(r[i])||0),0);
    if(isOfficeRow(name,soVal)){
      // SO-level figure from the office's own summary row - kept SEPARATE from BO sums.
      const k=soKey(name);
      const e=soTotals[k]??(soTotals[k]={name,total:0,sol:''});
      e.total+=v;
      if(solIdIdx>=0&&!e.sol)e.sol=String(r[solIdIdx]??'').trim();
      return;
    }
    const bo=matchBO(name);
    if(!bo)return;
    totals[bo]=(totals[bo]||0)+v;
    if(soIdx>=0&&!so[bo])so[bo]=soVal;
    if(solIdIdx>=0&&!sol[bo])sol[bo]=String(r[solIdIdx]??'').trim();
  });
  return {label,date:extractReportDate(matrix,hIdx),totals,soTotals,so,sol};
}
function overlayKpiParts(parts){
  // Latest exports replace any history rows of the same date, keyed by date|BO,
  // so a re-run for one date never double-counts in the trend chart.
  const key=(d,b)=>d+'|'+b;
  const byKey=new Map(ROWS.map(r=>[key(r.date,r.bo),r]));
  parts.forEach(p=>{
    const d=iso(p.date||new Date());
    Object.entries(p.totals).forEach(([bo,v])=>{
      const k=key(d,bo);
      let row=byKey.get(k);
      if(!row){row={date:d,bo,so:'',sol:'',kind:'bo'};byKey.set(k,row);}
      row[p.label]=Number(v)||0;
      if(!row.so&&p.so&&p.so[bo])row.so=p.so[bo];
      if(!row.sol&&p.sol&&p.sol[bo])row.sol=p.sol[bo];
    });
    // SO-level summary rows from the export: kept as kind='so' so the SO-wise view
    // can show the SO's own reported figure separately from the sum of its BOs.
    Object.entries(p.soTotals||{}).forEach(([k,e])=>{
      const name=e.name||k;
      const kk=key(d,name);
      let row=byKey.get(kk);
      if(!row){row={date:d,bo:name,so:name,sol:e.sol||'',kind:'so'};byKey.set(kk,row);}
      row[p.label]=Number(e.total)||0;
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
  const bits=[];
  if(hist)bits.push(`history ${DATES.length} day(s)`);
  if(parts.length)bits.push(`${parts.length}/${sources.length} KPI export(s) @ ${DATES[DATES.length-1]}`);
  if(missing.length)bits.push(`missing: ${missing.join(', ')}`);
  bits.push(`${ROWS.length} rows, ${new Set(ROWS.map(r=>r.bo)).size} BOs, ${soCount} SOs`);
  $('autoStatus').textContent='Auto-loaded '+bits.join(' • ')+'.';
}
function filtered(){
  const mode=$('mode').value,from=$('from').value,to=$('to').value;
  const bo=$('bo').value.trim().toLowerCase(),so=$('so').value.trim().toLowerCase();
  let a=from,b=to;
  if(mode==='single'){a=b=to}
  if(mode==='upto'){a=DATES[0]}
  return ROWS.filter(r=>(!a||r.date>=a)&&(!b||r.date<=b)&&
    (!bo||r.bo.toLowerCase().includes(bo))&&
    (!so||String(r.so||'').toLowerCase().includes(so)));
}
function sortBy(key){
  if(sortState.key===key)sortState.dir*=-1;
  else sortState={key,dir:1};
  render();
}
function render(){
  if(!ROWS.length)return;
  const rows=filtered(), act=activeMetrics();
  const gby=$('groupBy').value;
  // SO summary rows (kind='so') are SO-level figures: exclude them from KPI cards,
  // the trend chart and the BO-wise table so nothing is double-counted.
  const boRows=rows.filter(r=>r.kind!=='so');
  const soRows=rows.filter(r=>r.kind==='so');
  const days=[...new Set(boRows.map(r=>r.date))].sort();
  // Per-day sums drive the day-over-day delta (▲/▼ vs previous day) on metric KPI cards.
  const daySum={};
  boRows.forEach(r=>{daySum[r.date]??={};act.forEach(m=>daySum[r.date][m]=(daySum[r.date][m]||0)+(Number(r[m])||0));});
  const lastD=days[days.length-1], prevD=days.length>1?days[days.length-2]:null;
  const deltaOf=m=>prevD?((daySum[lastD]?.[m]||0)-(daySum[prevD]?.[m]||0)):null;
  const kpis=[['BOs',new Set(boRows.map(r=>r.bo)).size],
              ['SOs',new Set(rows.map(r=>r.so).filter(Boolean)).size],
              ...act.map(m=>[m,boRows.reduce((s,r)=>s+(Number(r[m])||0),0)])];
  $('kpis').innerHTML=kpis.map(([l,v])=>{
    const d=act.includes(l)?deltaOf(l):null;
    const sub=d===null?'':`<div class="d ${d>=0?'good':'bad'}">${d>=0?'▲':'▼'} ${fmt(Math.abs(d))} vs ${esc(prevD)}</div>`;
    return `<div class="kpi"><div class="l">${esc(l)}</div><div class="v">${fmt(v)}</div>${sub}</div>`;
  }).join('');
  // Aggregate into groups. BO-wise = BO rows only. SO-wise = BO rows summed per SO,
  // with the SO's own reported (summary-row) figures kept separately in soVals.
  let list;
  if(gby==='so'){
    const groups=new Map();
    boRows.forEach(r=>{
      const k=soKey(r.so)||'(no so)';
      let g=groups.get(k);
      if(!g){g={name:(r.so||'').trim()||'(no SO)',bos:new Set(),vals:{},total:0,soVals:null};groups.set(k,g);}
      g.bos.add(r.bo);
      act.forEach(m=>{const v=Number(r[m])||0;g.vals[m]=(g.vals[m]||0)+v;g.total+=v;});
    });
    soRows.forEach(r=>{
      const k=soKey(r.so||r.bo);
      let g=groups.get(k);
      if(!g){g={name:(r.bo||'').trim(),bos:new Set(),vals:{},total:0,soVals:{}};groups.set(k,g);}
      if(g.name==='(no SO)')g.name=(r.bo||'').trim();
      g.soVals=g.soVals||{};
      act.forEach(m=>g.soVals[m]=Number(r[m])||0);
    });
    list=[...groups.values()].map(g=>({name:g.name,so:'',bos:g.bos.size,vals:g.vals,total:g.total,soVals:g.soVals}));
  }else{
    const groups=new Map();
    boRows.forEach(r=>{
      let g=groups.get(r.bo);
      if(!g){g={name:r.bo,so:'',bos:new Set(),vals:{},total:0};groups.set(r.bo,g);}
      if(!g.so&&r.so)g.so=r.so;
      g.bos.add(r.bo);
      act.forEach(m=>{const v=Number(r[m])||0;g.vals[m]=(g.vals[m]||0)+v;g.total+=v;});
    });
    list=[...groups.values()].map(g=>({name:g.name,so:g.so,bos:g.bos.size,vals:g.vals,total:g.total,soVals:null}));
  }
  const q=$('quick').value;
  let sk=sortState.key, sd=sortState.dir;
  if(q==='top'){sk='total';sd=-1;}
  else if(q==='bottom'){sk='total';sd=1;}
  const soValOf=(g,base)=>!g.soVals?0:(base==='total'?Object.values(g.soVals).reduce((x,y)=>x+Number(y||0),0):(g.soVals[base]||0));
  list.sort((a,b)=>{
    let va,vb;
    if(sk==='name'){va=a.name;vb=b.name;}
    else if(sk==='so'){va=a.so;vb=b.so;}
    else if(sk==='bos'){va=a.bos;vb=b.bos;}
    else if(sk==='total'){va=a.total;vb=b.total;}
    else if(sk.indexOf('so:')===0){const base=sk.slice(3);va=soValOf(a,base);vb=soValOf(b,base);}
    else{va=a.vals[sk]||0;vb=b.vals[sk]||0;}
    return (typeof va==='string')?sd*va.localeCompare(vb):sd*(va-vb);
  });
  const totalGroups=list.length;
  if(q==='top'||q==='bottom')list=list.slice(0,10);
  const arrow=key=>(sk===key)?(sd>0?' ▲':' ▼'):'';
  const th=(key,label)=>`<th class="sortable" data-k="${esc(key)}">${esc(label)}${arrow(key)}</th>`;
  // SO-wise view keeps the two figures separate: "<metric> (ΣBO)" = sum of the SO's BO
  // rows, "<metric> (SO)" = the office's own summary-row figure ('—' when the report
  // has no such row). '(no SO)' buckets BO rows whose SO Name is blank in the report.
  const head=gby==='so'
    ?'<thead><tr>'+th('name','SO')+th('bos','BOs')+
     act.map(m=>th(m,m+' (ΣBO)')+th('so:'+m,m+' (SO)')).join('')+
     th('total','Total (ΣBO)')+th('so:total','Total (SO)')+'</tr></thead>'
    :'<thead><tr>'+th('name','BO')+th('so','SO')+act.map(m=>th(m,m)).join('')+th('total','Total')+'</tr></thead>';
  const body=gby==='so'
    ?list.map(g=>`<tr><td>${esc(g.name)}</td><td>${g.bos}</td>`+
      act.map(m=>`<td><b>${fmt(g.vals[m]||0)}</b></td><td>${g.soVals?fmt(g.soVals[m]||0):'—'}</td>`).join('')+
      `<td><b>${fmt(g.total)}</b></td><td>${g.soVals?fmt(soValOf(g,'total')):'—'}</td></tr>`).join('')
    :list.map(g=>`<tr><td>${esc(g.name)}</td><td>${esc(g.so||'')}</td>`+
      act.map(m=>`<td><b>${fmt(g.vals[m]||0)}</b></td>`).join('')+
      `<td><b>${fmt(g.total)}</b></td></tr>`).join('');
  const boMetricTotal=m=>boRows.reduce((s,r)=>s+(Number(r[m])||0),0);
  const soMetricTotal=m=>soRows.reduce((s,r)=>s+(Number(r[m])||0),0);
  const foot=gby==='so'
    ?'<tfoot><tr><td><b>Total</b></td><td><b>'+totalGroups+'</b></td>'+
     act.map(m=>`<td><b>${fmt(boMetricTotal(m))}</b></td><td><b>${fmt(soMetricTotal(m))}</b></td>`).join('')+
     `<td><b>${fmt(boRows.reduce((s,r)=>act.reduce((ss,m)=>ss+(Number(r[m])||0),0),0))}</b></td>`+
     `<td><b>${fmt(soRows.reduce((s,r)=>act.reduce((ss,m)=>ss+(Number(r[m])||0),0),0))}</b></td></tr></tfoot>`
    :'<tfoot><tr><td><b>Total</b></td><td><b>'+totalGroups+'</b></td>'+
     act.map(m=>`<td><b>${fmt(boMetricTotal(m))}</b></td>`).join('')+
     `<td><b>${fmt(boRows.reduce((s,r)=>act.reduce((ss,m)=>ss+(Number(r[m])||0),0),0))}</b></td></tr></tfoot>`;
  $('tbl').innerHTML=head+'<tbody>'+body+'</tbody>'+foot;
  $('tblTitle').textContent=(gby==='so'?'SO-wise':'BO-wise')+' cumulative performance';
  const skLabel=sk==='name'?(gby==='so'?'SO':'BO'):sk;
  $('tblInfo').textContent=`Showing ${list.length} of ${totalGroups} ${gby==='so'?'SOs':'BOs'}`+
    (q!=='all'?` (${q} 10)`:'')+` • ${boRows.length} BO rows${soRows.length?` + ${soRows.length} SO rows`:''} • sorted by ${skLabel} ${sd>0?'▲':'▼'}`;
  $('printHead').innerHTML=`<h1 style="font-size:16px">BO Performance Dashboard</h1>`+
    `<div style="font-size:12px">Mode: ${esc($('mode').value)} • From: ${esc($('from').value||'—')} • To: ${esc($('to').value||'—')}`+
    ` • BO filter: ${esc($('bo').value||'All')} • SO filter: ${esc($('so').value||'All')} • Group by: ${esc(gby)} • Data as of ${esc($('dataAsOf').textContent)}</div>`;
  let run={};const series={};act.forEach(m=>series[m]=[]);
  days.forEach(d=>{
    boRows.filter(r=>r.date===d).forEach(r=>act.forEach(m=>run[m]=(run[m]||0)+(Number(r[m])||0)));
    act.forEach(m=>series[m].push(run[m]||0));
  });
  const data={labels:days,datasets:act.map(m=>({label:m,data:series[m],tension:.3}))};
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
  const rows=all.filter(r=>r.kind!=='so'); // BO-wise pivot from BO rows only
  const out=all.map(r=>{
    const o={Date:r.date,Kind:r.kind||'bo',BO:r.bo,SO:r.so||'','SOL ID':r.sol||''};
    act.forEach(m=>o[m]=Number(r[m])||0);
    return o;
  });
  // BO-wise pivot sheet mirroring the on-screen table.
  const byBO={};
  rows.forEach(r=>{
    byBO[r.bo]??={BO:r.bo,SO:r.so||'',_t:0};
    if(!byBO[r.bo].SO&&r.so)byBO[r.bo].SO=r.so;
    act.forEach(m=>{const v=Number(r[m])||0;byBO[r.bo][m]=(byBO[r.bo][m]||0)+v;byBO[r.bo]._t+=v;});
  });
  const piv=Object.values(byBO).sort((a,b)=>a.BO.localeCompare(b.BO)).map(o=>{
    const p={BO:o.BO,SO:o.SO};
    act.forEach(m=>p[m]=o[m]||0);
    p.Total=o._t;
    return p;
  });
  const tot={BO:'Total',SO:''};
  act.forEach(m=>tot[m]=rows.reduce((s,r)=>s+(Number(r[m])||0),0));
  tot.Total=rows.reduce((s,r)=>act.reduce((ss,m)=>ss+(Number(r[m])||0),0),0);
  piv.push(tot);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(out),'Filtered Data');
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(piv),'BO-wise');
  XLSX.writeFile(wb,`BO_dashboard_${$('to').value||'export'}.xlsx`);
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
