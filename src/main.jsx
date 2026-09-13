import React, {useEffect, useMemo, useState} from 'react';
import {createRoot} from 'react-dom/client';
import {PublicClientApplication} from '@azure/msal-browser';
import * as XLSX from 'xlsx';
import {seedData} from './seedData';
import './styles.css';

const KEY='the-boss-mpg-v1';
const CLOUD_FILE='the-boss-mpg-data.json';
const microsoftClientId = import.meta.env.VITE_MICROSOFT_CLIENT_ID;

const productionRedirectUri =
  "https://the-boss-mpg-tracker.vercel.app/auth.html";

const localRedirectUri =
  "http://localhost:5173/auth.html";

const redirectUri = window.location.hostname === "localhost"
  ? localRedirectUri
  : productionRedirectUri;

const msal = microsoftClientId
  ? new PublicClientApplication({
      auth: {
        clientId: microsoftClientId,
        authority: "https://login.microsoftonline.com/common",
        redirectUri,
        postLogoutRedirectUri:
          "https://the-boss-mpg-tracker.vercel.app"
      },
      cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: false
      }
    })
  : null;

const loginRequest = {
  scopes: ["User.Read", "Files.ReadWrite"],
  redirectUri,
  prompt: "select_account"
};
const money=new Intl.NumberFormat('en-US',{style:'currency',currency:'USD'});
const num=(n,d=2)=>Number.isFinite(n)?n.toFixed(d):'—';
const uid=()=>`${Date.now()}-${Math.random().toString(16).slice(2)}`;
function initial(){try{const x=JSON.parse(localStorage.getItem(KEY)); if(x?.fuel&&x?.def)return x;}catch{} return {
  fuel: [],
  def: []
};}
async function graphToken(account) {
  try {
    const result = await msal.acquireTokenSilent({
      scopes: ["User.Read", "Files.ReadWrite"],
      account
    });

    return result.accessToken;
  } catch (error) {
    const result = await msal.acquireTokenPopup({
      scopes: ["User.Read", "Files.ReadWrite"],
      account,
      redirectUri
    });

    return result.accessToken;
  }
}
async function cloudRequest(account,method='GET',body){const token=await graphToken(account);const response=await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${CLOUD_FILE}:/content`,{method,headers:{Authorization:`Bearer ${token}`,...(body?{'Content-Type':'application/json'}:{})},body:body?JSON.stringify(body):undefined});if(!response.ok)throw Error(response.status===404?'missing':'request');return response.status===204?null:response.json();}
function calc(rows){return [...rows].sort((a,b)=>a.date.localeCompare(b.date)).map((r,i,all)=>{const prev=all[i-1];const distance=Number(r.distance)||((prev&&r.odometer>prev.odometer)?r.odometer-prev.odometer:0);return {...r,distance,efficiency:r.gallons?distance/r.gallons:0,unitCost:r.gallons?r.cost/r.gallons:0,costMile:distance?r.cost/distance:0};});}
function Summary({rows,type}){const totalMiles=rows.reduce((s,r)=>s+r.distance,0), gallons=rows.reduce((s,r)=>s+r.gallons,0), cost=rows.reduce((s,r)=>s+r.cost,0), valid=rows.filter(r=>r.efficiency>0), average=valid.length?valid.reduce((s,r)=>s+r.efficiency,0)/valid.length:0, best=valid.length?Math.max(...valid.map(r=>r.efficiency)):0, worst=valid.length?Math.min(...valid.map(r=>r.efficiency)):0;return <div className={`cards ${type==='fuel'?'dashboard':''}`}>{type==='fuel'?<><article><span>Average MPG</span><b>{num(average)}</b></article><article><span>Best MPG</span><b>{num(best)}</b></article><article><span>Worst MPG</span><b>{num(worst)}</b></article><article><span>Fuel spend</span><b>{money.format(cost)}</b></article></>:<><article><span>Total miles</span><b>{num(totalMiles,0)}</b></article><article><span>Miles / DEF gal</span><b>{num(gallons?totalMiles/gallons:0)}</b></article><article><span>Total gallons</span><b>{num(gallons,2)}</b></article><article><span>Total cost</span><b>{money.format(cost)}</b></article></>}</div>}
function TrendChart({rows,title,field,color,format}){const values=rows.map(r=>Number(r[field])||0), max=Math.max(...values,1), min=Math.min(...values,0), range=max-min||1, width=640, height=220, pad={top:18,right:18,bottom:34,left:46};const points=rows.map((r,i)=>{const x=rows.length===1?width/2:pad.left+i*(width-pad.left-pad.right)/(rows.length-1);const y=pad.top+(max-(Number(r[field])||0))*(height-pad.top-pad.bottom)/range;return {...r,x,y,value:Number(r[field])||0};});const area=points.length>1?`${pad.left},${height-pad.bottom} ${points.map(p=>`${p.x},${p.y}`).join(' ')} ${points.at(-1).x},${height-pad.bottom}`:'';return <article className="chart"><div className="chartHead"><h3>{title}</h3><span>{rows.length?'Per fill-up':'No entries yet'}</span></div>{rows.length?<svg viewBox={`0 0 ${width} ${height}`} role="img" aria-label={title}><line className="axis" x1={pad.left} x2={pad.left} y1={pad.top} y2={height-pad.bottom}/><line className="axis" x1={pad.left} x2={width-pad.right} y1={height-pad.bottom} y2={height-pad.bottom}/><polygon className="chartArea" points={area} fill={color}/><polyline className="chartLine" points={points.map(p=>`${p.x},${p.y}`).join(' ')} stroke={color}/>{points.map(p=><circle key={p.id} cx={p.x} cy={p.y} r="4" fill={color}><title>{`${p.date}: ${format(p.value)}`}</title></circle>)}<text className="chartLabel" x={pad.left} y={height-10}>{rows[0].date}</text><text className="chartLabel" textAnchor="end" x={width-pad.right} y={height-10}>{rows.at(-1).date}</text><text className="chartValue" x={pad.left-8} y={pad.top+4} textAnchor="end">{format(max)}</text><text className="chartValue" x={pad.left-8} y={height-pad.bottom+4} textAnchor="end">{format(min)}</text></svg>:<div className="chartEmpty">Add a fill-up to start tracking this trend.</div>}</article>}
function Trends({rows,type}){return <section className="trends"><TrendChart rows={rows} title={type==='fuel'?'MPG over time':'DEF efficiency over time'} field="efficiency" color="#4fd1a8" format={v=>type==='fuel'?`${num(v)} MPG`:`${num(v)} mi/gal`}/><TrendChart rows={rows} title="Cost over time" field="cost" color="#f0b35b" format={v=>money.format(v)}/></section>}
function App(){
 const [data,setData]=useState(initial); const [tab,setTab]=useState('fuel'); const [msg,setMsg]=useState(''); const [historyOpen,setHistoryOpen]=useState(true); const [account,setAccount]=useState(null); const [syncStatus,setSyncStatus]=useState(msal?'OneDrive not connected':'Local storage only');
 const [form,setForm]=useState({date:new Date().toISOString().slice(0,10),odometer:'',gallons:'',cost:''});
 const rows=useMemo(()=>calc(data[tab]),[data,tab]);
 useEffect(()=>{if(!msal)return;let active=true;(async()=>{try{await msal.initialize();const cached=msal.getAllAccounts()[0];if(!cached)return;setAccount(cached);const remote=await cloudRequest(cached);if(active&&remote?.fuel&&remote?.def){setData(remote);localStorage.setItem(KEY,JSON.stringify(remote));setSyncStatus('Synced from OneDrive');}}catch{if(active)setSyncStatus('OneDrive sync unavailable');}})();return()=>{active=false};},[]);
 const save=n=>{setData(n);localStorage.setItem(KEY,JSON.stringify(n));if(account){setSyncStatus('Saving to OneDrive…');cloudRequest(account,'PUT',n).then(()=>setSyncStatus('Synced to OneDrive')).catch(()=>setSyncStatus('Saved locally; OneDrive sync failed'));}};
 const connectOneDrive=async()=>{if(!msal)return;try{setSyncStatus('Signing in…');await msal.initialize();const result=await msal.loginPopup(loginRequest);const signedIn=result.account;setAccount(signedIn);try{const remote=await cloudRequest(signedIn);if(remote?.fuel&&remote?.def){setData(remote);localStorage.setItem(KEY,JSON.stringify(remote));setSyncStatus('Synced from OneDrive');return;}}catch(error){
  if(error.message !== 'missing'){
    throw error;
  }

  const emptyData = {
    fuel: [],
    def: []
  };

  await cloudRequest(signedIn, 'PUT', emptyData);

  setData(emptyData);
  localStorage.setItem(KEY, JSON.stringify(emptyData));

  setSyncStatus('Created new OneDrive profile');

  return;
}} catch (error) {
  console.error("OneDrive connection error:", error);

  setSyncStatus(
    `OneDrive failed: ${
      error?.errorCode ||
      error?.message ||
      "Unknown error"
    }`
  );
}};
 const add=e=>{e.preventDefault();const odometer=+form.odometer,gallons=+form.gallons,cost=+form.cost;if(!form.date||odometer<=0||gallons<=0||cost<0)return setMsg('Enter a valid date, odometer, gallons, and cost.');const prev=rows.at(-1);const distance=prev&&odometer>prev.odometer?odometer-prev.odometer:0;save({...data,[tab]:[...data[tab],{id:uid(),type:tab,date:form.date,odometer,gallons,cost,distance}]});setForm(f=>({...f,odometer:'',gallons:'',cost:''}));setMsg(distance?'Entry saved.':'Entry saved. Add a later odometer reading to calculate distance.');};
 const remove=id=>save({...data,[tab]:data[tab].filter(x=>x.id!==id)});
 const exportFile=()=>{const wb=XLSX.utils.book_new();for(const t of ['fuel','def']){const out=calc(data[t]).map(r=>({Type:t==='fuel'?'Diesel':'DEF',Date:r.date,Odometer:r.odometer,'Miles since prior fill':r.distance,Gallons:r.gallons,Cost:r.cost,'Cost/Gallon':r.unitCost,'Miles/Gallon':r.efficiency,'Cost/Mile':r.costMile}));XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(out),t==='fuel'?'Diesel':'DEF');}XLSX.writeFile(wb,'The-Boss-MPG-Backup.xlsx');};
 const importFile=async e=>{const file=e.target.files[0];if(!file)return;try{const wb=XLSX.read(await file.arrayBuffer(),{cellDates:true});const imported={fuel:[],def:[]};const parse=(sheet,type)=>{const a=XLSX.utils.sheet_to_json(sheet,{defval:''});a.forEach((r,i)=>{const date=r.Date||r.DATE;const odo=+(r.Odometer||r['EndODO']||r['END ODO']);const gallons=+(r.Gallons||r['TOTAL GALLONS']);const cost=+(r.Cost||r['TOTAL FUEL COST']||r['TOTAL DEF COST']);const distance=+(r['Miles since prior fill']||r['TRIP METER']||r['Trip miles']);if(date&&gallons){const ds=date instanceof Date?date.toISOString().slice(0,10):String(date);imported[type].push({id:`import-${type}-${i}-${uid()}`,type,date:ds,odometer:odo||0,gallons,cost,distance});}})};wb.SheetNames.forEach(n=>parse(wb.Sheets[n],n.toLowerCase().includes('def')?'def':'fuel')); if(!imported.fuel.length&&!imported.def.length)throw Error();save({fuel:imported.fuel.length?imported.fuel:data.fuel,def:imported.def.length?imported.def:data.def});setMsg('Spreadsheet imported.');}catch{setMsg('Could not read that layout. Use the exported backup format or the original tracker workbook.');}e.target.value='';};
 return <main><header><div><p className="eyebrow">THE BOSS</p><h1>MPG Tracker</h1><p>Diesel and DEF fill-up history, all in one place.</p></div><div className="actions"><button className="button secondary" onClick={connectOneDrive} disabled={!msal}>{msal?(account?'OneDrive connected':'Connect OneDrive'):'Set Microsoft client ID'}</button><label className="button secondary">Import Excel<input type="file" accept=".xlsx,.xls" onChange={importFile}/></label><button className="button secondary" onClick={exportFile}>Export backup</button><span className="syncStatus">{syncStatus}</span></div></header>
 <nav><button className={tab==='fuel'?'active':''} onClick={()=>setTab('fuel')}>Diesel</button><button className={tab==='def'?'active':''} onClick={()=>setTab('def')}>DEF</button></nav>
 <Summary rows={rows} type={tab}/>
 <Trends rows={rows} type={tab}/>
 <section className="panel"><h2>Add {tab==='fuel'?'diesel':'DEF'} fill-up</h2><form onSubmit={add}><label>Date<input type="date" value={form.date} onChange={e=>setForm({...form,date:e.target.value})}/></label><label>Truck odometer<input inputMode="decimal" placeholder="24,750" value={form.odometer} onChange={e=>setForm({...form,odometer:e.target.value.replace(/,/g,'')})}/></label><label>Gallons<input inputMode="decimal" placeholder="18.5" value={form.gallons} onChange={e=>setForm({...form,gallons:e.target.value})}/></label><label>Fuel cost<input inputMode="decimal" placeholder="105.00" value={form.cost} onChange={e=>setForm({...form,cost:e.target.value})}/></label><button className="button" type="submit">Save fill-up</button></form>{msg&&<p className="message">{msg}</p>}</section>
 <section className="panel tablePanel"><div className="sectionHead"><h2>History</h2><div><span>{rows.length} entries</span><button className="button secondary historyToggle" type="button" onClick={()=>setHistoryOpen(open=>!open)} aria-expanded={historyOpen}>{historyOpen?'Collapse':'Expand'}</button></div></div>{historyOpen&&<div className="tableWrap"><table><thead><tr><th>Date</th><th>Odometer</th><th>Miles</th><th>Gallons</th><th>Cost</th><th>$/gal</th><th>{tab==='fuel'?'MPG':'Miles/DEF gal'}</th><th></th></tr></thead><tbody>{[...rows].reverse().map(r=><tr key={r.id}><td>{r.date}</td><td>{num(r.odometer,0)}</td><td>{num(r.distance,0)}</td><td>{num(r.gallons)}</td><td>{money.format(r.cost)}</td><td>{money.format(r.unitCost)}</td><td className="strong">{num(r.efficiency)}</td><td><button className="delete" onClick={()=>remove(r.id)} aria-label="Delete entry">×</button></td></tr>)}</tbody></table></div>}</section>
 <footer>Data is saved in this browser. Export a backup regularly.</footer></main>}
createRoot(document.getElementById('root')).render(<App/>);
