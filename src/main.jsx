import React, {useMemo, useState} from 'react';
import {createRoot} from 'react-dom/client';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { Upload, FileSpreadsheet, FileText, Search, Download, Users, CheckCircle2, AlertCircle, FolderArchive, RefreshCw, ShieldCheck, FileCheck2 } from 'lucide-react';
import './style.css';

const SHEET = 'REG-10j Individual Records';
const norm = v => String(v ?? '').trim().toLowerCase().replace(/\s+/g,' ');
const fmtDate = value => {
  if (value === null || value === undefined || value === '') return '';
  if (typeof value === 'number') {
    const parts=XLSX.SSF.parse_date_code(value);
    if (!parts) return String(value);
    return `${String(parts.d).padStart(2,'0')}-${new Date(Date.UTC(parts.y,parts.m-1,1)).toLocaleString('en-GB',{month:'short',timeZone:'UTC'})}-${parts.y}`;
  }
  let d = value;
  if (!(d instanceof Date)) d = new Date(d);
  if (!d || isNaN(d.getTime())) return String(value);
  return `${String(d.getUTCDate()).padStart(2,'0')}-${d.toLocaleString('en-GB',{month:'short',timeZone:'UTC'})}-${d.getUTCFullYear()}`;
};
function parseWorkbook(buffer) {
  const wb = XLSX.read(buffer,{type:'array',cellDates:false});
  const ws = wb.Sheets[SHEET];
  if (!ws) throw new Error(`Sheet "${SHEET}" was not found.`);
  const rows = XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:true});
  const headerIndex = rows.findIndex(r => r.some(c => norm(c)==='employee code') && r.some(c => norm(c)==='curriculum ref.'));
  if (headerIndex < 0) throw new Error('Could not locate the employee/training column headers in REG-10j Individual Records.');
  const headers = rows[headerIndex].map(norm);
  const idx = name => headers.indexOf(norm(name));
  const required = ['employee code','employee name','curriculum ref.','date delivered','delivered by'];
  const missing = required.filter(k=>idx(k)<0);
  if (missing.length) throw new Error('Missing Excel columns: '+missing.join(', '));
  const records=[];
  for (const r of rows.slice(headerIndex+1)) {
    const code=String(r[idx('employee code')]??'').trim();
    const ref=String(r[idx('curriculum ref.')]??'').trim();
    if (!code || !ref) continue;
    records.push({code,name:String(r[idx('employee name')]??'').trim(),ref,date:fmtDate(r[idx('date delivered')]),deliveredBy:String(r[idx('delivered by')]??'').trim()});
  }
  return records;
}
function employeeIdFromFilename(name) {
  const m = name.match(/ACF[-_]QMS[-_]FRM[-_]51[_-](\d+)[_-]/i) || name.match(/(?:^|[_-])(\d{2,})[_-][^/]+\.docx$/i);
  return m ? m[1] : null;
}
function textOf(el) { return [...el.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main','t')].map(n=>n.textContent).join(' ').trim(); }
function setRunFont(run, doc) {
  const ns='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  let rPr=run.getElementsByTagNameNS(ns,'rPr')[0];
  if (!rPr) {
    rPr=doc.createElementNS(ns,'w:rPr');
    run.insertBefore(rPr,run.firstChild);
  }
  let rFonts=rPr.getElementsByTagNameNS(ns,'rFonts')[0];
  if (!rFonts) {
    rFonts=doc.createElementNS(ns,'w:rFonts');
    rPr.insertBefore(rFonts,rPr.firstChild);
  }
  rFonts.setAttributeNS(ns,'w:ascii','Calibri');
  rFonts.setAttributeNS(ns,'w:hAnsi','Calibri');
  rFonts.setAttributeNS(ns,'w:eastAsia','Calibri');
  rFonts.setAttributeNS(ns,'w:cs','Calibri');
  for (const tag of ['sz','szCs']) {
    let size=rPr.getElementsByTagNameNS(ns,tag)[0];
    if (!size) size=rPr.appendChild(doc.createElementNS(ns,'w:'+tag));
    size.setAttributeNS(ns,'w:val','14');
  }
}
function setCellText(cell, value, doc) {
  const ns='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const texts=[...cell.getElementsByTagNameNS(ns,'t')];
  if (texts.length) {
    texts[0].textContent=value;
    if (value.startsWith(' ') || value.endsWith(' ')) texts[0].setAttributeNS('http://www.w3.org/XML/1998/namespace','xml:space','preserve');
    for (const t of texts.slice(1)) t.textContent='';
    for (const r of new Set(texts.map(t=>t.parentNode))) setRunFont(r,doc);
  } else {
    let p=cell.getElementsByTagNameNS(ns,'p')[0];
    if (!p) { p=doc.createElementNS(ns,'w:p'); cell.appendChild(p); }
    let r=doc.createElementNS(ns,'w:r'), t=doc.createElementNS(ns,'w:t'); t.textContent=value;
    setRunFont(r,doc);
    r.appendChild(t); p.appendChild(r);
  }
}
async function updateDocx(file, records) {
  const zip=await JSZip.loadAsync(file);
  const docXml=zip.file('word/document.xml');
  if (!docXml) throw new Error('Invalid DOCX: word/document.xml not found.');
  const xml=await docXml.async('string');
  const doc=new DOMParser().parseFromString(xml,'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error('Could not parse Word document XML.');
  const ns='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const tables=[...doc.getElementsByTagNameNS(ns,'tbl')];
  let target=null, colRef=-1, colDate=-1, colBy=-1, dataStart=-1;
  for (const table of tables) {
    const rows=[...table.getElementsByTagNameNS(ns,'tr')];
    if (!rows.length) continue;
    const cells=rows.slice(0,3).map(row=>[...row.getElementsByTagNameNS(ns,'tc')]);
    const joined=cells.map(cs=>cs.map(textOf));
    const headerFlat=joined.flat().map(norm);
    const ref=headerFlat.findIndex(x=>x==='ref');
    const date=headerFlat.findIndex(x=>x.includes('date') && x.includes('delivered'));
    const by=headerFlat.findIndex(x=>x.includes('delivered') && x.includes('by'));
    if (ref>=0 && date>=0 && by>=0) {
      // Header cells are normally in the first table row; retain column positions from that row.
      const headerRowIndex=joined.findIndex(a=>a.some(v=>norm(v)==='ref'));
      const headerCells=cells[headerRowIndex]||cells[0];
      const localRef=headerCells.findIndex(c=>norm(textOf(c))==='ref');
      const localDate=headerCells.findIndex(c=>norm(textOf(c)).includes('date')&&norm(textOf(c)).includes('delivered'));
      const localBy=headerCells.findIndex(c=>norm(textOf(c)).includes('delivered')&&norm(textOf(c)).includes('by'));
      if (localRef>=0 && localDate>=0 && localBy>=0) { target=rows; colRef=localRef; colDate=localDate; colBy=localBy; dataStart=headerRowIndex+1; break; }
      // The supplied template uses a two-line header where the header labels still belong to individual cells.
      const headerCells2=cells[0];
      const all=headerCells2.map(textOf).map(norm);
      const r=all.findIndex(x=>x==='ref'), d=all.findIndex(x=>x.includes('date')&&x.includes('delivered')), b=all.findIndex(x=>x.includes('delivered')&&x.includes('by'));
      if (r>=0&&d>=0&&b>=0) {target=rows;colRef=r;colDate=d;colBy=b;dataStart=1;break;}
    }
  }
  if (!target) throw new Error('Could not identify the Required Training table headers (Ref, Date delivered, Delivered by).');
  const byRef=new Map(records.map(r=>[norm(r.ref),r]));
  let updated=0, matched=0;
  for (const row of target.slice(dataStart)) {
    const cells=[...row.getElementsByTagNameNS(ns,'tc')];
    if (cells.length<=Math.max(colRef,colDate,colBy)) continue;
    const ref=norm(textOf(cells[colRef]));
    if (!ref || !byRef.has(ref)) continue;
    const rec=byRef.get(ref); matched++;
    setCellText(cells[colDate],rec.date,doc);
    setCellText(cells[colBy],rec.deliveredBy,doc);
    updated++;
  }
  if (!matched) throw new Error('No matching curriculum references were found in the Word training table.');
  zip.file('word/document.xml',new XMLSerializer().serializeToString(doc));
  return {blob:await zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'}),updated};
}
function App(){
 const [excel,setExcel]=useState(null),[records,setRecords]=useState([]),[docs,setDocs]=useState([]),[query,setQuery]=useState(''),[busy,setBusy]=useState(false),[status,setStatus]=useState(''),[results,setResults]=useState([]),[error,setError]=useState('');
 const employees=useMemo(()=>{const m=new Map();records.forEach(r=>{if(!m.has(r.code))m.set(r.code,{code:r.code,name:r.name,records:[]});m.get(r.code).records.push(r)});return [...m.values()].sort((a,b)=>a.name.localeCompare(b.name));},[records]);
 const filtered=employees.filter(e=>norm(e.name+' '+e.code).includes(norm(query)));
 async function loadExcel(file){setError('');setStatus('');try{const data=parseWorkbook(await file.arrayBuffer());setExcel(file);setRecords(data);setResults([]);setStatus(`Loaded ${data.length} training rows for ${new Set(data.map(x=>x.code)).size} employees.`);}catch(e){setError(e.message);}}
 function loadDocs(fileList){setError('');const list=[...fileList].filter(f=>f.name.toLowerCase().endsWith('.docx'));setDocs(list);setStatus(`${list.length} Word documents selected.`);}
 async function process(selected){
   if(!records.length||!docs.length){setError('Upload the Excel workbook and select employee Word documents first.');return;}
   setBusy(true);setError('');setResults([]);const out=[];
   try{
     for(let i=0;i<selected.length;i++){
       const emp=selected[i];setStatus(`Processing ${i+1} of ${selected.length}: ${emp.name}`);
       const matches=docs.filter(f=>employeeIdFromFilename(f.name)===emp.code);
       if(!matches.length){out.push({code:emp.code,name:emp.name,ok:false,message:'No Word document matched employee ID in filename.'});continue;}
       const file=matches[0];
       try{const res=await updateDocx(file,emp.records);out.push({code:emp.code,name:emp.name,ok:true,fileName:file.name.replace(/\.docx$/i,'_UPDATED.docx'),blob:res.blob,updated:res.updated});}
       catch(e){out.push({code:emp.code,name:emp.name,ok:false,message:e.message});}
     }
     setResults(out);setStatus(`Finished: ${out.filter(x=>x.ok).length} updated, ${out.filter(x=>!x.ok).length} needing attention.`);
   }catch(e){setError(e.message);}finally{setBusy(false);}
 }
 function downloadOne(item){const url=URL.createObjectURL(item.blob),a=document.createElement('a');a.href=url;a.download=item.fileName;a.click();URL.revokeObjectURL(url);}
 async function downloadAll(){const zip=new JSZip();results.filter(r=>r.ok).forEach(r=>zip.file(r.fileName,r.blob));const blob=await zip.generateAsync({type:'blob'});const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download='Employee-Training-Records-Updated.zip';a.click();URL.revokeObjectURL(url);}
 const done=results.filter(x=>x.ok).length;
 return <div className="app"><header className="top"><div className="brand"><div className="brandIcon"><FileCheck2/></div><div><b>Training Records</b><small>Employee document automation</small></div></div><div className="local"><span/> Browser-only processing</div></header>
 <main><section className="hero"><div><div className="eyebrow">ACUF​​ORE · QUALITY MANAGEMENT</div><h1>Employee Training<br/><em>Record Automation</em></h1><p>Update employee Word records using training delivery details from your Excel register. Your files stay in your browser.</p></div><div className="heroArt"><div className="artSheet"><FileSpreadsheet size={34}/><span>REG-10j</span></div><div className="artArrow">→</div><div className="artDoc"><FileText size={38}/><span>FRM-51</span></div></div></section>
 <section className="stats"><div><Users/><b>{employees.length||'—'}</b><span>Employees loaded</span></div><div><FileSpreadsheet/><b>{records.length||'—'}</b><span>Training rows</span></div><div><FileText/><b>{docs.length||'—'}</b><span>Word files selected</span></div><div><CheckCircle2/><b>{done}</b><span>Documents updated</span></div></section>
 <section className="steps"><div className="stepTitle"><span>01</span><div><h2>Load source files</h2><p>Choose the Excel register and the employee Word documents.</p></div></div><div className="uploadGrid">
 <label className="uploadCard"><input type="file" accept=".xlsx,.xls" onChange={e=>e.target.files[0]&&loadExcel(e.target.files[0])}/><div className="uploadIcon green"><FileSpreadsheet/></div><div><b>{excel?excel.name:'Upload Excel workbook'}</b><small>{excel?'Workbook loaded · '+records.length+' training rows':'Select ACF-QMS-REG-10f Training Plan and Record'}</small></div><span className="choose">{excel?'Replace':'Browse'}</span></label>
 <label className="uploadCard"><input type="file" accept=".docx" multiple webkitdirectory="" directory="" onChange={e=>loadDocs(e.target.files)}/><div className="uploadIcon blue"><FolderArchive/></div><div><b>Select employee documents folder</b><small>{docs.length?docs.length+' DOCX files selected':'Choose the folder containing employee .docx files'}</small></div><span className="choose">Browse</span></label></div>
 <div className="hint"><ShieldCheck size={16}/> Files are processed locally in this browser. Nothing is sent to a backend server.</div></section>
 <section className="steps"><div className="stepTitle"><span>02</span><div><h2>Find employees</h2><p>Match employee IDs against document names such as ACF-QMS-FRM-51_176_Cibisakkaravarthi_K.docx.</p></div></div>
 <div className="searchBar"><Search/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Search employee name or employee ID..."/><span>{filtered.length} employees</span></div>
 <div className="employeeList">{filtered.slice(0,8).map(emp=><div className="employee" key={emp.code}><div className="avatar">{emp.name.split(' ').slice(0,2).map(x=>x[0]).join('').toUpperCase()}</div><div className="empInfo"><b>{emp.name}</b><small>ID {emp.code} · {emp.records.length} training items</small></div><button disabled={busy||!docs.length} onClick={()=>process([emp])}>Update record <span>→</span></button></div>)}
 {filtered.length>8&&<div className="more">Showing 8 of {filtered.length}. Refine your search to find another employee.</div>}
 {!employees.length&&<div className="empty">Upload the Excel workbook to load employees.</div>}
 </div>
 <div className="bulk"><div><b>Bulk update</b><small>Process all employees that have a matching Word document.</small></div><button className="primary" disabled={busy||!employees.length||!docs.length} onClick={()=>process(employees)}>{busy?<><RefreshCw className="spin"/> Processing…</>:<>Update all employees <span>→</span></>}</button></div></section>
 {status&&<div className="notice"><CheckCircle2 size={17}/>{status}</div>}{error&&<div className="error"><AlertCircle size={17}/>{error}</div>}
 {results.length>0&&<section className="steps results"><div className="stepTitle"><span>03</span><div><h2>Processing results</h2><p>{done} successful · {results.length-done} need attention</p></div></div><div className="resultHead"><b>Employee</b><b>Result</b><b>Action</b></div>{results.map((r,i)=><div className="resultRow" key={r.code+'-'+i}><div><b>{r.name}</b><small>ID {r.code}</small></div><div>{r.ok?<span className="success">Updated {r.updated} rows</span>:<span className="failed">{r.message}</span>}</div><div>{r.ok&&<button className="download" onClick={()=>downloadOne(r)}><Download size={15}/> DOCX</button>}</div></div>)}{done>0&&<button className="primary downloadAll" onClick={downloadAll}><Download size={17}/> Download all updated documents (ZIP)</button>}</section>}
 <footer><span>ACF-QMS-FRM-51 · Individual Training Record</span><span>Only “Date delivered” and “Delivered by” are modified.</span></footer></main></div>
}
createRoot(document.getElementById('root')).render(<App/>);
