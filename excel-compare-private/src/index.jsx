import React, {useEffect, useMemo, useState} from "react";
import {createRoot} from "react-dom/client";
import * as XLSX from "xlsx";

/* utils */
const esc = v => (v==null ? "" : String(v));
const trim = s => esc(s).trim();
const uniq = a => [...new Set(a)];
const unionHeaders = (a,b) => uniq([...(a||[]), ...(b||[])]);
const byKeyMap = (rows, key) => {
  const m = new Map();
  (rows||[]).forEach(r => { const k = esc(r[key]); if(k) m.set(k, r); });
  return m;
};

/* xlsx helpers */
async function readWorkbook(file){ const buf = await file.arrayBuffer(); return XLSX.read(buf,{type:"array"}); }
function sheetToRows(wb, sheetName){
  const name = (sheetName && wb.Sheets[sheetName]) ? sheetName : wb.SheetNames[0];
  const ws = wb.Sheets[name]; if(!ws) return {headers:[], rows:[], sheetName:name||""};
  const arr = XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
  if(!arr.length) return {headers:[], rows:[], sheetName:name};
  const headers = arr[0].map(h => trim(h));
  const rows = arr.slice(1).map(r => {
    const o={}; headers.forEach((h,i)=> o[h] = trim(r[i] ?? "")); return o;
  }).filter(o=> Object.values(o).some(v=>v!==""));
  return {headers, rows, sheetName:name};
}
function writeComparedWorkbook({fileAName,fileBName, headers, key, mapA, mapB, diffs}) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([...mapA.values()], {header:headers}), "FileA");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([...mapB.values()], {header:headers}), "FileB");
  const diffRows = diffs.map(d => ({ RowKey: d.rowKey, Column: d.column, Left_Before: d.left, Right_After: d.right }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(diffRows, {header:["RowKey","Column","Left_Before","Right_After"]}), "Diff");
  XLSX.writeFile(wb, `Compared_${fileAName}_vs_${fileBName}.xlsx`);
}

/* diff */
function computeDiff({Arows,Brows,headers,key,opts}) {
  const norm = rows => rows.map(r=>{
    const o={}; headers.forEach(h=>{
      let v = esc(r[h] ?? "");
      if(opts.trim) v = v.trim();
      if(opts.ci) v = v.toLowerCase();
      o[h]=v;
    }); return o;
  });

  const AA = norm(Arows), BB = norm(Brows);
  const mapA = byKeyMap(AA, key), mapB = byKeyMap(BB, key);

  const keysA = new Set(mapA.keys()), keysB = new Set(mapB.keys());
  const added = [...keysB].filter(k => !keysA.has(k));
  const removed = [...keysA].filter(k => !keysB.has(k));
  const common = [...keysA].filter(k => keysB.has(k));

  const diffs=[]; 
  common.forEach(k=>{
    const a = mapA.get(k), b = mapB.get(k);
    headers.forEach(h=>{
      if(h===key) return;
      const va = esc(a[h] ?? ""), vb = esc(b[h] ?? "");
      if(va !== vb){ diffs.push({rowKey:k, column:h, left:va, right:vb}); }
    });
  });

  return { mapA, mapB, added, removed, diffs };
}

/* small UI primitives */
const UploadCard = ({label,onFile}) => (
  <div className="card pad center">
    <div className="hint">{label}</div>
    <div className="uploader" onDragOver={e=>e.preventDefault()}
         onDrop={e=>{e.preventDefault(); const f=e.dataTransfer.files?.[0]; if(f) onFile(f);}}>
      <div style={{fontSize:40,color:"var(--primary)"}}>⤴️</div>
      <div className="hint mt8">Drag & drop, or</div>
      <label className="btn primary mt8">
        Browse Files
        <input type="file" accept=".xlsx,.xls,.csv" onChange={e=> e.target.files[0] && onFile(e.target.files[0])}/>
      </label>
      <div className="hint mt8">Accepted: .xlsx, .xls, .csv</div>
    </div>
  </div>
);

function HelpModal({onClose}){
  return (
    <div className="help-modal" onClick={onClose}>
      <div className="modal" onClick={e=>e.stopPropagation()}>
        <div className="hdr"><b>How to use</b><button className="btn outline" onClick={onClose}>Close</button></div>
        <div className="body">
          <ol>
            <li>Upload <b>File A</b> and <b>File B</b> (XLSX/CSV). Nothing is uploaded anywhere.</li>
            <li>Click <b>Configure & Compare</b>, pick a <b>Key</b> (or auto) and optional sheet names.</li>
            <li>Review results. Yellow = changed cells, Green = added rows, Red = deleted rows.</li>
            <li>Use <b>Download Compared File</b> to export A/B + Diff sheet.</li>
          </ol>
          <p className="hint mt12">Privacy: no network allowed (CSP). Works offline.</p>
        </div>
      </div>
    </div>
  );
}

function ConfigModal({open,onClose,headersA,headersB,sheetA,sheetB,setSheetA,setSheetB,key,setKey,opts,setOpts,onApply}){
  if(!open) return null;
  const options = [...new Set([...(headersA||[]), ...(headersB||[])])];
  return (
    <div className="help-modal" onClick={onClose}>
      <div className="modal" onClick={e=>e.stopPropagation()}>
        <div className="hdr"><b>Configure Comparison</b><button className="btn outline" onClick={onClose}>Cancel</button></div>
        <div className="body">
          <div className="mt8">
            <div className="hint">Key Identifier Column</div>
            <select value={key} onChange={e=>setKey(e.target.value)} style={{width:'100%',padding:'8px',border:'1px solid var(--line)',borderRadius:8}}>
              <option value="">(Auto: first column)</option>
              {options.map(o=><option key={o} value={o}>{o}</option>)}
            </select>
          </div>
          <div className="grid2 mt12">
            <div>
              <div className="hint">Sheet (File A)</div>
              <input value={sheetA} onChange={e=>setSheetA(e.target.value)} placeholder="e.g. Sheet1"
                     style={{width:'100%',padding:'8px',border:'1px solid var(--line)',borderRadius:8}}/>
            </div>
            <div>
              <div className="hint">Sheet (File B)</div>
              <input value={sheetB} onChange={e=>setSheetB(e.target.value)} placeholder="e.g. Sheet1"
                     style={{width:'100%',padding:'8px',border:'1px solid var(--line)',borderRadius:8}}/>
            </div>
          </div>
          <div className="row mt12">
            <label className="row"><input type="checkbox" checked={opts.trim} onChange={e=>setOpts({...opts,trim:e.target.checked})}/> Ignore surrounding spaces</label>
            <label className="row"><input type="checkbox" checked={opts.ci} onChange={e=>setOpts({...opts,ci:e.target.checked})}/> Case-insensitive</label>
          </div>
        </div>
        <div className="row" style={{justifyContent:'flex-end',padding:'12px',borderTop:'1px solid var(--line)'}}>
          <button className="btn primary" onClick={onApply}>Apply & Compare</button>
        </div>
      </div>
    </div>
  );
}

function ResultsView({fileAName,fileBName, headers, key, mapA, mapB, added, removed, diffs, onBack}){
  const [onlyDiff,setOnlyDiff] = useState(true);
  const [showChanged,setShowChanged] = useState(true);

  const allRowKeys = useMemo(()=>{
    const s=new Set([...mapA.keys(), ...mapB.keys()]);
    const arr=[...s].sort((a,b)=>esc(a).localeCompare(esc(b)));
    if(onlyDiff){
      const changedRows = new Set(diffs.map(d=>d.rowKey));
      return arr.filter(k => changedRows.has(k) || added.includes(k) || removed.includes(k));
    }
    return arr;
  },[mapA,mapB,diffs,added,removed,onlyDiff]);

  const download = ()=> writeComparedWorkbook({fileAName,fileBName, headers, key, mapA, mapB, diffs});

  return (
    <div className="container">
      <div className="row" style={{justifyContent:'space-between'}}>
        <div>
          <div style={{fontSize:22,fontWeight:700}}>Results</div>
          <div className="hint mt8">{fileAName} vs {fileBName}</div>
        </div>
        <div className="row">
          <button className="btn primary" onClick={download}>Download Compared File (.xlsx)</button>
          <button className="btn outline" onClick={onBack}>Back</button>
        </div>
      </div>

      <div className="kpis">
        <div className="ok"><b>{added.length}</b> Rows Added</div>
        <div className="warn"><b>{diffs.length}</b> Cells Changed</div>
        <div className="bad"><b>{removed.length}</b> Rows Deleted</div>
      </div>

      <div className="row mt12">
        <label className="row"><input type="checkbox" checked={onlyDiff} onChange={e=>setOnlyDiff(e.target.checked)}/> Show Only Differences</label>
        <label className="row"><input type="checkbox" checked={showChanged} onChange={e=>setShowChanged(e.target.checked)}/> Highlight Changed</label>
      </div>

      <div className="grid2 mt16">
        <div className="card">
          <div className="row" style={{justifyContent:'space-between',padding:'12px',borderBottom:'1px solid var(--line)'}}>
            <b>File A (Original)</b><span className="hint">Key: {key}</span>
          </div>
          <div className="table-wrap">
            <table>
              <thead><tr>{headers.map(h=><th key={h}>{h}</th>)}</tr></thead>
              <tbody>
              {allRowKeys.map(k=>{
                const r = mapA.get(k);
                const isRemoved = !mapB.has(k);
                return (
                  <tr key={k} className={isRemoved?'row-removed':''}>
                    {headers.map(h=>{
                      const v = r?.[h] ?? "";
                      const changed = showChanged && mapB.has(k) && (h!==key) && (esc(v)!==esc(mapB.get(k)?.[h] ?? ""));
                      return <td key={h} className={changed?'chg':''}>{esc(v)}</td>;
                    })}
                  </tr>
                );
              })}
              </tbody>
            </table>
          </div>
        </div>
        <div className="card">
          <div className="row" style={{justifyContent:'space-between',padding:'12px',borderBottom:'1px solid var(--line)'}}>
            <b>File B (Compared)</b><span className="hint">Key: {key}</span>
          </div>
          <div className="table-wrap">
            <table>
              <thead><tr>{headers.map(h=><th key={h}>{h}</th>)}</tr></thead>
              <tbody>
              {allRowKeys.map(k=>{
                const r = mapB.get(k);
                const isAdded = !mapA.has(k);
                return (
                  <tr key={k} className={isAdded?'row-added':''}>
                    {headers.map(h=>{
                      const v = r?.[h] ?? "";
                      const changed = showChanged && mapA.has(k) && (h!==key) && (esc(v)!==esc(mapA.get(k)?.[h] ?? ""));
                      return <td key={h} className={changed?'chg':''}>{esc(v)}</td>;
                    })}
                  </tr>
                );
              })}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      <div className="hint mt16">Runs entirely locally. CSP blocks any network access.</div>
    </div>
  );
}

function App(){
  const [helpOpen,setHelpOpen]=useState(false);
  const [fileA,setFileA]=useState(null), [fileB,setFileB]=useState(null);
  const [nameA,setNameA]=useState(""), [nameB,setNameB]=useState("");
  const [headersA,setHeadersA]=useState([]), [headersB,setHeadersB]=useState([]);
  const [sheetA,setSheetA]=useState(""), [sheetB,setSheetB]=useState("");
  const [parsedA,setParsedA]=useState(null), [parsedB,setParsedB]=useState(null);
  const [key,setKey]=useState(""), [opts,setOpts]=useState({trim:false,ci:false});
  const [stage,setStage]=useState("upload");
  const [headers,setHeaders]=useState([]), [mapA,setMapA]=useState(new Map()), [mapB,setMapB]=useState(new Map());
  const [added,setAdded]=useState([]), [removed,setRemoved]=useState([]), [diffs,setDiffs]=useState([]);

  useEffect(()=>{ const btn=document.getElementById('help-btn'); const h=()=>setHelpOpen(true); btn?.addEventListener('click',h); return ()=>btn?.removeEventListener('click',h); },[]);

  useEffect(()=>{ (async ()=>{
    if(fileA){ const wb=await readWorkbook(fileA); const {headers,rows}=sheetToRows(wb, sheetA||undefined); setHeadersA(headers); setParsedA({headers,rows}); setNameA(fileA.name); }
  })(); },[fileA,sheetA]);

  useEffect(()=>{ (async ()=>{
    if(fileB){ const wb=await readWorkbook(fileB); const {headers,rows}=sheetToRows(wb, sheetB||undefined); setHeadersB(headers); setParsedB({headers,rows}); setNameB(fileB.name); }
  })(); },[fileB,sheetB]);

  const canCompare = !!(fileA && fileB && parsedA && parsedB);
  const [showConfig,setShowConfig]=useState(false);

  async function applyAndCompare(){
    if(!canCompare) return;
    const allHeaders = unionHeaders(parsedA.headers, parsedB.headers);
    const k = key || parsedA.headers[0] || allHeaders[0];
    const {mapA,mapB,added,removed,diffs} = computeDiff({
      Arows: parsedA.rows, Brows: parsedB.rows, headers: allHeaders, key:k, opts
    });
    setHeaders(allHeaders); setMapA(mapA); setMapB(mapB); setAdded(added); setRemoved(removed); setDiffs(diffs); setStage("results"); setShowConfig(false);
  }

  if(stage==="results"){
    return <ResultsView fileAName={nameA} fileBName={nameB} headers={headers} key={key||headers[0]} mapA={mapA} mapB={mapB} added={added} removed={removed} diffs={diffs} onBack={()=>setStage("upload")}/>;
  }

  return (
    <div className="container">
      <h1 style={{fontSize:28,fontWeight:800}}>Excel Compare: Spot Spreadsheet Differences Instantly</h1>
      <p className="hint mt8">Quickly upload, compare, and analyze your data.</p>

      <div className="grid2 mt16">
        <UploadCard label="UPLOAD FILE A" onFile={setFileA}/>
        <UploadCard label="UPLOAD FILE B" onFile={setFileB}/>
      </div>

      <div className="row mt12">
        <button className="btn primary" disabled={!canCompare} onClick={()=>setShowConfig(true)} style={{opacity:canCompare?1:.5,cursor:canCompare?'pointer':'not-allowed'}}>Configure & Compare</button>
        {fileA && <span className="hint">A: {nameA}</span>}
        {fileB && <span className="hint">B: {nameB}</span>}
      </div>

      <p className="hint mt16">Your data stays on your device. No uploads.</p>

      <ConfigModal
        open={showConfig} onClose={()=>setShowConfig(false)}
        headersA={headersA} headersB={headersB}
        sheetA={sheetA} sheetB={sheetB} setSheetA={setSheetA} setSheetB={setSheetB}
        key={key} setKey={setKey} opts={opts} setOpts={setOpts}
        onApply={applyAndCompare}
      />

      {helpOpen && <HelpModal onClose={()=>setHelpOpen(false)}/>}
    </div>
  );
}

createRoot(document.getElementById("app")).render(<App/>);
