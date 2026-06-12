/*
 * SISTEM KOORDINASI KAMAR BEDAH — RS Panti Rini
 * Versi FINAL · Zero Bug · Zero Demo Data · Semua Real
 * PIN default: 1234
 * Modifikasi: Supabase Backup, PERAWAT Instrumen, Statistik
 */
import { useState, useEffect, useRef, useCallback, Component } from "react";
import * as XLSX from "xlsx";
import { createClient, type RealtimeChannel } from "@supabase/supabase-js";
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ReferenceLine, ResponsiveContainer } from "recharts";

/* ─── CONFIG ────────────────────────────────────────────────────────── */
const CONFIG = {
  HOSPITAL: "Rumah Sakit Panti Rini",
  DEFAULT_PIN: "1234",
};
const HOSPITAL = CONFIG.HOSPITAL;
const LOCK_MS  = 10 * 60 * 1000;

/* ─── CONSTANTS ─────────────────────────────────────────────────────── */
const ROOMS = ["Kamar Operasi 1","Kamar Operasi 2","Kamar Operasi 3","Kamar Operasi 4","Kamar Operasi 5"];
const BT    = ["A+","A-","B+","B-","AB+","AB-","O+","O-","Tidak Diketahui"];
const ST    = { surgeon:"Dokter Bedah", anesthesiologist:"Dokter Anestesi", circulating:"PERAWAT Instrumen", anesthesia_nurse:"Perawat Anestesi", onloop:"Perawat Onloop", katim:"RR / Katim" };
const OT    = { elektif:{label:"Elektif",c:"#1565C0",bg:"#E3F2FD"}, semi:{label:"Semi-Elektif",c:"#E65100",bg:"#FFF3E0"}, cyto:{label:"⚠ CYTO",c:"#B71C1C",bg:"#FFCDD2"} };
const STS   = { scheduled:{l:"Terjadwal",c:"#1565C0",bg:"#E3F2FD"}, ongoing:{l:"Berlangsung",c:"#00897B",bg:"#E0F2F1"}, done:{l:"Selesai",c:"#2E7D32",bg:"#E8F5E9"}, batal:{l:"Batal/Tunda",c:"#C62828",bg:"#FFEBEE"} };
const C     = { p:"#00695C",pL:"#00897B",pBg:"#E0F2F1", d:"#C62828",dBg:"#FFEBEE",dL:"#EF9A9A", w:"#E65100",wBg:"#FFF3E0", i:"#1565C0",iBg:"#E3F2FD", s:"#2E7D32",sBg:"#E8F5E9", wa:"#25D366",waBg:"#DCFCE7", g:"#546E7A",gBg:"#ECEFF1", t:"#1A2332",tL:"#607080", b:"#DDE3EA",white:"#FFF",bg:"#F4F7FA" };
const EOP   = { patient:"",age:"",rm:"",opType:"elektif",diagnosis:"",procedure:"",ruangAsal:"",room:ROOMS[0],date:"",time:"",surgeon:"",anesthesiologist:"",assistantNurse:"",circulatingNurse:"",anesthesiaNurse:"",onloopNurse:"",rrKatim:"",allergy:"Tidak Ada",specialNeeds:"",bloodType:"O+" };
const PAGE_SIZE = 10;
const DEFAULT_RECIPIENT = "Suster Thresmiati CB, bu Niken, pak Jaka dan teman sejawat, mohon ijin laporan kamar bedah:";

/* ─── HELPERS ───────────────────────────────────────────────────────── */
const gId       = () => `${Date.now()}-${Math.random().toString(36).slice(2,6)}`;
const todayDate = () => new Date().toISOString().split("T")[0];
const tmrwDate  = () => { const d=new Date(); d.setDate(d.getDate()+1); return d.toISOString().split("T")[0]; };
const fD        = (d: string) => { if(!d)return"-"; try{ return new Date(d+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}); }catch{ return d; } };
const fNow      = () => new Date().toLocaleString("id-ID",{day:"numeric",month:"long",year:"numeric",hour:"2-digit",minute:"2-digit"});
const fTR       = (t: string) => t?t.replace(":",".") : "";
const gWord     = () => { const t=new Date().getHours()*60+new Date().getMinutes(); return t<660?"pagi":t<900?"siang":t<1110?"sore":"malam"; };
const sanitize  = (s: any) => String(s||"").replace(/<[^>]*>/g,"").replace(/['"`;]/g,"").trim();
const maskName  = (n: string,a: boolean) => !a||!n ? n??"" : String(n).replace(/\S+/g,(w:string)=>w[0]+"*".repeat(Math.max(1,w.length-1)));
const maskRM    = (r: string,a: boolean) => !a||!r ? r??"" : String(r).slice(0,2)+"****"+String(r).slice(-2);
const parseDateCSV = (s: string) => { if(!s)return""; const m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); return m?`${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`:s; };

/* ─── REAL FILE HELPERS ─────────────────────────────────────────────── */
const downloadBlob = (content: string, filename: string, mime="text/csv") => {
  const blob = new Blob(["\uFEFF"+content], {type:mime});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href=url; a.download=filename;
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
};
const parseCSV = (text: string) => text.split("\n").filter(r=>r.trim())
  .map(r=>r.split(",").map(c=>c.trim().replace(/^["']|["']$/g,"")));

/* ─── SUPABASE BACKUP ───────────────────────────────────────────────── */
interface SupabaseConfig { url: string; anonKey: string; autoBackup: boolean; backupInterval: number; lastBackup: string | null; realtimeBackup: boolean; autoExcelBackup: boolean; lastExcelBackup: string | null; lastExcelBackupTs: number | null; }
const defaultSupaCfg: SupabaseConfig = { url:"", anonKey:"", autoBackup:false, backupInterval:60, lastBackup:null, realtimeBackup:false, autoExcelBackup:false, lastExcelBackup:null, lastExcelBackupTs:null };

/* ─── DROPBOX BACKUP ────────────────────────────────────────────────── */
interface DropboxConfig { token: string; path: string; autoBackup: boolean; backupInterval: number; realtimeBackup: boolean; lastBackup: string | null; autoExcelBackup: boolean; lastExcelBackup: string | null; lastExcelBackupTs: number | null; }
const defaultDbxCfg: DropboxConfig = { token:"", path:"/KamarBedahPantiRini/backup.json", autoBackup:false, backupInterval:30, realtimeBackup:false, lastBackup:null, autoExcelBackup:false, lastExcelBackup:null, lastExcelBackupTs:null };

async function dropboxUpload(cfg: DropboxConfig, data: any): Promise<{ok:boolean;msg:string}> {
  if(!cfg.token) return {ok:false, msg:"Access Token Dropbox belum diisi"};
  try {
    const arg = JSON.stringify({path: cfg.path, mode:"overwrite", autorename:false, mute:true});
    const body = JSON.stringify({exportedAt: new Date().toISOString(), ...data});
    const res = await fetch("https://content.dropboxapi.com/2/files/upload", {
      method:"POST",
      headers:{"Authorization":"Bearer "+cfg.token,"Content-Type":"application/octet-stream","Dropbox-API-Arg":arg},
      body,
    });
    if(!res.ok){ const t=await res.text(); return {ok:false, msg:`Error ${res.status}: ${t.slice(0,120)}`}; }
    return {ok:true, msg:`✓ Tersimpan ke Dropbox: ${cfg.path}`};
  } catch(e:any){ return {ok:false, msg:"Gagal menghubungi Dropbox: "+(e?.message||"network error")}; }
}

async function dropboxDownload(cfg: DropboxConfig): Promise<{ok:boolean;data?:any;msg:string}> {
  if(!cfg.token) return {ok:false, msg:"Access Token Dropbox belum diisi"};
  try {
    const arg = JSON.stringify({path: cfg.path});
    const res = await fetch("https://content.dropboxapi.com/2/files/download", {
      method:"POST",
      headers:{"Authorization":"Bearer "+cfg.token,"Dropbox-API-Arg":arg},
    });
    if(!res.ok){ const t=await res.text(); return {ok:false, msg:`Error ${res.status}: ${t.slice(0,120)}`}; }
    const data = await res.json();
    return {ok:true, data, msg:"✓ Data berhasil dimuat dari Dropbox"};
  } catch(e:any){ return {ok:false, msg:"Gagal mengunduh dari Dropbox: "+(e?.message||"network error")}; }
}

/* Upload Excel per-kategori ke Dropbox */
async function dropboxUploadExcel(token: string, path: string, wb: any): Promise<{ok:boolean;msg:string}> {
  if(!token) return {ok:false, msg:"Access Token Dropbox belum diisi"};
  try {
    const buf = XLSX.write(wb, {type:"array", bookType:"xlsx"});
    const arg = JSON.stringify({path, mode:"overwrite", autorename:false, mute:true});
    const res = await fetch("https://content.dropboxapi.com/2/files/upload", {
      method:"POST",
      headers:{"Authorization":"Bearer "+token,"Content-Type":"application/octet-stream","Dropbox-API-Arg":arg},
      body: new Uint8Array(buf),
    });
    if(!res.ok){ const t=await res.text(); return {ok:false, msg:`Error ${res.status}: ${t.slice(0,120)}`}; }
    return {ok:true, msg:`✓ Excel tersimpan ke Dropbox: ${path}`};
  } catch(e:any){ return {ok:false, msg:"Gagal upload Excel ke Dropbox: "+(e?.message||"network error")}; }
}

/* Download sebagai Word (.doc via HTML) */
function downloadAsWord(title: string, rows: string[][], filename: string) {
  const thead = rows[0].map(h=>`<th style="background:#004D40;color:#fff;padding:6px 10px;font-size:11px">${h}</th>`).join("");
  const tbody = rows.slice(1).map((r,i)=>`<tr style="background:${i%2===0?"#F0FFF8":"#fff"}">${r.map(c=>`<td style="padding:5px 10px;border-bottom:1px solid #e0e0e0;font-size:11px">${c||""}</td>`).join("")}</tr>`).join("");
  const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8"><title>${title}</title></head><body><h2 style="color:#004D40;font-family:Arial">${title}</h2><p style="color:#555;font-size:12px;font-family:Arial">Dicetak: ${fNow()} — ${HOSPITAL}</p><table style="border-collapse:collapse;width:100%;font-family:Arial"><thead><tr>${thead}</tr></thead><tbody>${tbody}</tbody></table></body></html>`;
  const blob = new Blob([html], {type:"application/msword"});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href=url; a.download=filename;
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}

async function supabaseBackup(cfg: SupabaseConfig, data: any): Promise<{ok: boolean; msg: string}> {
  if(!cfg.url || !cfg.anonKey) return {ok:false, msg:"URL dan Anon Key Supabase belum diisi"};
  try {
    const endpoint = cfg.url.replace(/\/$/, "") + "/rest/v1/kamar_bedah_backup";
    const payload = { created_at: new Date().toISOString(), data: JSON.stringify(data) };
    const res = await fetch(endpoint, {
      method: "POST",
      headers: {
        "apikey": cfg.anonKey,
        "Authorization": "Bearer " + cfg.anonKey,
        "Content-Type": "application/json",
        "Prefer": "return=minimal"
      },
      body: JSON.stringify(payload)
    });
    if(res.ok || res.status === 201) return {ok:true, msg:"✓ Backup ke Supabase berhasil — "+fNow()};
    const txt = await res.text().catch(()=>"");
    return {ok:false, msg:"Supabase error " + res.status + ": " + txt.slice(0,120)};
  } catch(e: any) {
    return {ok:false, msg:"Koneksi gagal: " + (e?.message||"unknown error")};
  }
}

async function supabaseRestore(cfg: SupabaseConfig): Promise<{ok: boolean; msg: string; data?: any}> {
  if(!cfg.url || !cfg.anonKey) return {ok:false, msg:"URL dan Anon Key Supabase belum diisi"};
  try {
    const endpoint = cfg.url.replace(/\/$/, "") + "/rest/v1/kamar_bedah_backup?order=created_at.desc&limit=1";
    const res = await fetch(endpoint, {
      method: "GET",
      headers: {
        "apikey": cfg.anonKey,
        "Authorization": "Bearer " + cfg.anonKey,
        "Content-Type": "application/json"
      }
    });
    if(!res.ok) {
      const txt = await res.text().catch(()=>"");
      return {ok:false, msg:"Supabase error " + res.status + ": " + txt.slice(0,120)};
    }
    const rows = await res.json();
    if(!Array.isArray(rows) || rows.length === 0) return {ok:false, msg:"Tidak ada data backup di Supabase"};
    const parsed = JSON.parse(rows[0].data || "{}");
    return {ok:true, msg:"✓ Data backup ditemukan — "+rows[0].created_at?.slice(0,19)?.replace("T"," "), data:parsed};
  } catch(e: any) {
    return {ok:false, msg:"Koneksi gagal: " + (e?.message||"unknown error")};
  }
}

/* ─── MESSAGE BUILDERS ──────────────────────────────────────────────── */
function msgSurgeon(name: string, ops: any[]) {
  const s = [...ops].sort((a,b)=>a.time.localeCompare(b.time));
  const ua = [...new Set(s.map(o=>o.anesthesiologist).filter(Boolean))];
  const N = sanitize(name).toUpperCase();
  if(ua.length===1){
    const A = sanitize(ua[0]).toUpperCase();
    return `Selamat ${gWord()} *${N}*,\nuntuk operasi besok bius dengan *${A}* ada ${s.length} mulai jam ${fTR(s[0]?.time)}:\n\n${s.map((o,i)=>`${i+1}. rencana ${sanitize(o.procedure)} atas nama ${sanitize(o.patient)}${o.age?` ${o.age} thn`:""}`).join("\n")}\n\nTerima kasih\n\nKamar Bedah ${HOSPITAL}`;
  }
  return `Selamat ${gWord()} *${N}*,\nuntuk operasi besok ada ${s.length} mulai jam ${fTR(s[0]?.time)}:\n\n${s.map((o,i)=>`${i+1}. rencana ${sanitize(o.procedure)} atas nama ${sanitize(o.patient)}${o.age?` ${o.age} thn`:""}${o.anesthesiologist?` (bius: *${sanitize(o.anesthesiologist).toUpperCase()}*)`:""}`).join("\n")}\n\nTerima kasih\n\nKamar Bedah ${HOSPITAL}`;
}
function msgAnest(name: string, ops: any[]) {
  const s = [...ops].sort((a,b)=>a.time.localeCompare(b.time));
  const byTime: Record<string,any[]> = {};
  s.forEach(o=>{ if(!byTime[o.time])byTime[o.time]=[]; byTime[o.time].push(o); });
  const secs = Object.entries(byTime).sort(([a],[b])=>a.localeCompare(b)).map(([time,oa])=>{
    const bySurg: Record<string,any[]> ={};
    oa.forEach(o=>{ if(!bySurg[o.surgeon])bySurg[o.surgeon]=[]; bySurg[o.surgeon].push(o); });
    return `*Jam ${fTR(time)}*\n`+Object.entries(bySurg).map(([sg,so])=>
      `pasien *${sanitize(sg).toUpperCase()}* ada ${so.length}:\n${so.map((op,i)=>`${i+1}. ${sanitize(op.patient)}${op.age?` ${op.age} thn`:""} rencana ${sanitize(op.procedure)}`).join("\n")}`
    ).join("\n\n");
  }).join("\n\n");
  return `Selamat ${gWord()} ${sanitize(name)},\nuntuk acara besok ${fD(tmrwDate())} ada ${s.length} pasien:\n\n${secs}\n\nTerima kasih dokter\n\nKamar Bedah ${HOSPITAL}`;
}
function msgCyto(op: any) {
  const alg = op.allergy&&op.allergy!=="Tidak Ada" ? `⚠ ALERGI: ${sanitize(op.allergy)}\n` : "";
  const ra  = op.ruangAsal ? `Ruang Asal : ${sanitize(op.ruangAsal)}\n` : "";
  return `⚠️ LAPORAN ACARA OPERASI CYTO / EMERGENCY ⚠️\n━━━━━━━━━━━━━━━━━━━\nKamar Bedah ${HOSPITAL}\n━━━━━━━━━━━━━━━━━━━\n${fNow()}\n\nMOHON SEGERA KE KAMAR OPERASI\n\n${alg}${ra}Pasien   : ${sanitize(op.patient)}${op.age?` (${op.age} th)`:""}\nRM/Gol   : ${op.rm||"-"} / ${op.bloodType}\nDiagnosa : ${sanitize(op.diagnosis)}\nTindakan : ${sanitize(op.procedure)}\nJam      : ${op.time} WIB — ${op.room}\n\nOperator : ${op.surgeon||"-"}\nAnestesi : ${op.anesthesiologist||"-"}\nAsisten  : ${op.assistantNurse||"-"}\nInstrumen: ${op.circulatingNurse||"-"}\nP.Anest  : ${op.anesthesiaNurse||"-"}\nOnloop   : ${op.onloopNurse||"-"}\nRR/Katim : ${op.rrKatim||"-"}\n━━━━━━━━━━━━━━━━━━━\n⚡ HARAP RESPON SEGERA ⚡\nKamar Bedah ${HOSPITAL}`;
}
function buildLaporan({greeting,recipients,keterangan,todayOps,tomorrowOps,anestStandby,anestCyto,nurseStandby,pembawaHP}: any) {
  const active = [...todayOps].filter((o:any)=>o.status!=="batal").sort((a:any,b:any)=>a.time.localeCompare(b.time));
  const batal  = todayOps.filter((o:any)=>o.status==="batal");
  const fT = (op: any) => { const t=[]; if(op.assistantNurse)t.push(`asisten: ${op.assistantNurse}`); if(op.onloopNurse)t.push(`Onloop: ${op.onloopNurse}`); if(op.rrKatim)t.push(`RR/Katim: ${op.rrKatim}`); return t.length?` (${t.join(", ")})`:""};
  const done=new Set<string>(), grps: any[][] =[];
  active.forEach((op:any)=>{ if(!done.has(op.id)){ const g=active.filter((o:any)=>!done.has(o.id)&&o.time===op.time&&o.surgeon===op.surgeon); g.forEach((x:any)=>done.add(x.id)); grps.push(g); }});
  const tLines = grps.map(g=>{ const f=g[0],cm=f.opType==="cyto"?"⚠️ [CYTO] ⚠️ ":"",base=`${cm}Jam ${fTR(f.time)} ${f.surgeon} dengan ${f.anesthesiologist}`;
    return g.length===1?`_${base} operasi ${g[0].procedure} a/n ${g[0].patient}${g[0].age?`, ${g[0].age}th`:""}${fT(g[0])}_`:`_${base} operasi ${g.length}:_\n${g.map((o:any,i:number)=>`_${i+1}. ${o.procedure} a/n ${o.patient}${o.age?`, ${o.age}th`:""}${fT(o)}_`).join("\n")}`;
  });
  const batalLine = batal.length?`\n_Batal/Tunda: ${batal.map((o:any)=>o.patient+(o.cancelReason?` (${o.cancelReason})`:""  )).join("; ")}_`:"";
  const todaySec = active.length?`*Hari ini acara operasi ada ${active.length} berjalan lancar\n\n${tLines.join("\n\n")}${batalLine}*`:`*Hari ini tidak ada operasi berjalan${batalLine}*`;
  const pagi = tomorrowOps.filter((o:any)=>o.time<"14:00").sort((a:any,b:any)=>a.time.localeCompare(b.time));
  const sore = tomorrowOps.filter((o:any)=>o.time>="14:00").sort((a:any,b:any)=>a.time.localeCompare(b.time));
  const blk = (ops: any[]) => {
    if(!ops.length)return"";
    const bT: Record<string,any[]> ={};
    ops.forEach(o=>{ if(!bT[o.time])bT[o.time]=[]; bT[o.time].push(o); });
    return Object.entries(bT).sort(([a],[b])=>a.localeCompare(b)).map(([time,oa])=>{
      const uA=[...new Set(oa.map(o=>o.anesthesiologist).filter(Boolean))];
      const uS=[...new Set(oa.map(o=>o.surgeon))];
      if(uA.length===1&&uS.length>1) return `Jam ${fTR(time)}, Bersama *${(uA[0] as string).toUpperCase()}*:\n`+uS.map(sg=>{ const so=oa.filter(o=>o.surgeon===sg); return so.length===1?`- ${sg} rencana ${so[0].procedure} a/n ${so[0].patient}${so[0].age?`, ${so[0].age}th`:""}`:`- ${sg} rencana ${so.length} operasi:\n${so.map((op:any,i:number)=>`  ${i+1}. ${op.procedure} a/n ${op.patient}${op.age?`, ${op.age}th`:""}`).join("\n")}`; }).join("\n");
      return oa.map(op=>`Jam ${fTR(op.time)} ${op.surgeon} dengan *${op.anesthesiologist}* rencana ${op.procedure} a/n ${op.patient}${op.age?`, ${op.age}th`:""}`).join("\n\n");
    }).join("\n\n");
  };
  let tmrwSec = `Rencana operasi besok hari ${fD(tmrwDate())}:\n`;
  if(pagi.length) tmrwSec += `\n*PAGI*\nRencana operasi ${pagi.length}:\n\n${blk(pagi)}`;
  if(sore.length) tmrwSec += `\n\n*SORE*\n${blk(sore)}`;
  if(!pagi.length&&!sore.length) tmrwSec += "\nBelum ada rencana operasi";
  const aSec  = anestStandby.filter(Boolean).length ? `\n*Siaga Anaesthesi*\n${anestStandby.filter(Boolean).map((n:string,i:number)=>`${i+1}. ${n}`).join("\n")}` : "";
  const cSec  = anestCyto.filter(Boolean).length    ? `\n\n*Siaga Anaesthesi Cyto*\n${anestCyto.filter(Boolean).map((n:string,i:number)=>`${i+1}. ${n}`).join("\n")}` : "";
  const nSec  = nurseStandby.filter(Boolean).length ? `\n\n*Perawat siaga*\n${nurseStandby.filter(Boolean).map((n:string,i:number)=>`${i+1}. ${n}`).join("\n")}` : "";
  const phSec = pembawaHP?.trim() ? `\n\nPembawa HP: ${pembawaHP}` : "";
  const ket   = keterangan?.trim() ? `\n${keterangan}\n` : "";
  return `Selamat ${greeting} ${recipients}\n${ket}\n${todaySec}\n\n\n${tmrwSec}\n${aSec}${cSec}${nSec}${phSec}\n\nDemikian laporan kamar bedah, terima kasih🙏`;
}

/* ─── UI PRIMITIVES ─────────────────────────────────────────────────── */
const iS: React.CSSProperties = { width:"100%", padding:"11px 14px", borderRadius:10, border:`1.5px solid ${C.b}`, fontSize:14, lineHeight:1.6, background:"#FAFBFD", boxSizing:"border-box", fontFamily:"inherit", color:C.t, outline:"none" };

function LF({label,req,err,children}: {label?: string; req?: boolean; err?: string; children: React.ReactNode}) {
  return (
    <div style={{marginBottom:14}}>
      {label && <div style={{fontSize:12,fontWeight:700,color:C.g,marginBottom:5,textTransform:"uppercase",letterSpacing:.4}}>
        {label}{req&&<span style={{color:C.d}}> ✱</span>}
      </div>}
      {children}
      {err && <div style={{fontSize:11,color:C.d,marginTop:4}}>⚠ {err}</div>}
    </div>
  );
}
function TF({label,req,err,foc,onFoc,onBlr,...p}: any) {
  return (
    <LF label={label} req={req} err={err}>
      <input style={{...iS, borderColor:err?C.d:foc?C.p:C.b, boxShadow:foc?`0 0 0 3px ${C.p}22`:err?`0 0 0 3px ${C.d}22`:"none"} as React.CSSProperties}
        onFocus={onFoc} onBlur={onBlr} autoComplete="off" spellCheck={false} {...p}/>
    </LF>
  );
}
function TA({label,minH=72,...p}: any) {
  const [f,sf] = useState(false);
  return (
    <LF label={label}>
      <textarea style={{...iS,resize:"vertical",minHeight:minH,lineHeight:1.7, borderColor:f?C.p:C.b, boxShadow:f?`0 0 0 3px ${C.p}22`:""} as React.CSSProperties}
        onFocus={()=>sf(true)} onBlur={()=>sf(false)} spellCheck={false} {...p}/>
    </LF>
  );
}
function SF({label,req,err,options,empty="-- Pilih --",...p}: any) {
  const [f,sf] = useState(false);
  return (
    <LF label={label} req={req} err={err}>
      <div style={{position:"relative"}}>
        <select style={{...iS,appearance:"none",cursor:"pointer",paddingRight:32, borderColor:err?C.d:f?C.p:C.b, boxShadow:f?`0 0 0 3px ${C.p}22`:""} as React.CSSProperties}
          onFocus={()=>sf(true)} onBlur={()=>sf(false)} {...p}>
          <option value="">{empty}</option>
          {options.map((o: any,i: number)=><option key={i} value={typeof o==="string"?o:o.v}>{typeof o==="string"?o:o.l}</option>)}
        </select>
        <div style={{position:"absolute",right:12,top:"50%",transform:"translateY(-50%)",pointerEvents:"none",color:C.g,fontSize:11}}>▼</div>
      </div>
    </LF>
  );
}
const Btn = ({children,onClick,color=C.p,outline,sm,full,disabled,style={}}: any) => (
  <button onClick={onClick} disabled={disabled} style={{padding:sm?"7px 12px":"10px 18px",borderRadius:10,border:`1.5px solid ${disabled?"#ccc":color}`,background:outline?C.white:disabled?"#eee":color,color:outline?color:C.white,fontSize:sm?12:14,fontWeight:700,cursor:disabled?"not-allowed":"pointer",opacity:disabled?.45:1,width:full?"100%":"auto",fontFamily:"inherit",transition:"opacity .15s",...style}}>{children}</button>
);
const WaBt = ({children,onClick,sm,full,disabled,style={}}: any) => (
  <button onClick={onClick} disabled={disabled} style={{padding:sm?"7px 12px":"10px 18px",borderRadius:10,border:`1.5px solid ${disabled?"#ccc":C.wa}`,background:disabled?"#eee":C.wa,color:C.white,fontSize:sm?12:14,fontWeight:700,cursor:disabled?"not-allowed":"pointer",width:full?"100%":"auto",fontFamily:"inherit",display:"flex",alignItems:"center",justifyContent:"center",gap:5,...style}}>✓ {children}</button>
);
const Card = ({children,style={},hi}: any) => (
  <div style={{background:C.white,borderRadius:14,padding:"18px 20px",marginBottom:12,boxShadow:"0 2px 8px rgba(0,0,0,.06)",border:`1px solid ${hi||C.b}`,...style}}>{children}</div>
);
const Row  = ({title,right}: any) => (
  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
    <div style={{fontSize:16,fontWeight:800,color:C.t}}>{title}</div>
    {right}
  </div>
);
const Bdg  = ({label,color,bg}: any) => <span style={{background:bg,color,fontSize:10,fontWeight:700,padding:"3px 9px",borderRadius:20,border:`1px solid ${color}33`,whiteSpace:"nowrap"}}>{label}</span>;
const SH   = ({label,color=C.p}: any) => <div style={{fontSize:12,fontWeight:800,color,letterSpacing:.6,marginBottom:14,paddingBottom:8,borderBottom:`2px solid ${color}22`,textTransform:"uppercase"}}>{label}</div>;
const Toggle = ({value,onChange,label,sub,color=C.p}: any) => (
  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 0",borderBottom:`1px solid ${C.b}`}}>
    <div><div style={{fontSize:14,fontWeight:700,color:C.t}}>{label}</div>{sub&&<div style={{fontSize:11,color:C.tL,marginTop:2}}>{sub}</div>}</div>
    <button onClick={()=>onChange(!value)} style={{width:50,height:28,borderRadius:14,background:value?color:"#DDE3EA",border:"none",cursor:"pointer",position:"relative",transition:"background .2s",flexShrink:0}}>
      <div style={{position:"absolute",top:4,left:value?24:4,width:20,height:20,borderRadius:10,background:"#fff",boxShadow:"0 1px 4px rgba(0,0,0,.25)",transition:"left .2s"}}/>
    </button>
  </div>
);

/* ─── MINI BAR CHART ─────────────────────────────────────────────────── */
function MiniBar({data, colorFn, maxVal}: {data: {label: string; value: number; color?: string}[]; colorFn?: (i: number) => string; maxVal?: number}) {
  const max = maxVal ?? Math.max(...data.map(d=>d.value), 1);
  return (
    <div style={{display:"flex",flexDirection:"column",gap:6}}>
      {data.map((d,i)=>(
        <div key={i} style={{display:"flex",alignItems:"center",gap:8}}>
          <div style={{width:110,fontSize:11,color:C.tL,textAlign:"right",flexShrink:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{d.label}</div>
          <div style={{flex:1,background:"#F0F4F8",borderRadius:6,height:20,position:"relative",overflow:"hidden"}}>
            <div style={{position:"absolute",left:0,top:0,bottom:0,width:`${Math.round((d.value/max)*100)}%`,background:d.color||(colorFn?colorFn(i):C.p),borderRadius:6,transition:"width .4s"}}/>
          </div>
          <div style={{width:28,fontSize:12,fontWeight:700,color:C.t,textAlign:"left",flexShrink:0}}>{d.value}</div>
        </div>
      ))}
    </div>
  );
}

/* ─── ERROR BOUNDARY ────────────────────────────────────────────────── */
class ErrorBoundary extends Component<{children: React.ReactNode},{err:boolean;msg:string}> {
  state = {err:false,msg:""};
  static getDerivedStateFromError(e: any){ return {err:true,msg:e?.message||"Error"}; }
  render(){
    if(this.state.err) return (
      <div style={{padding:24,textAlign:"center",background:C.dBg,borderRadius:14,margin:"8px 0",border:`1px solid ${C.dL}`}}>
        <div style={{fontSize:32,marginBottom:8}}>⚠️</div>
        <div style={{fontSize:14,color:C.d,fontWeight:600,marginBottom:8}}>{this.state.msg}</div>
        <Btn onClick={()=>this.setState({err:false})}>Coba Lagi</Btn>
      </div>
    );
    return this.props.children;
  }
}

/* ─── PIN SCREEN ────────────────────────────────────────────────────── */
function PinScreen({onVerify,savedPin}: any) {
  const [mode,setMode] = useState("login");
  const [pin,setPin]   = useState(""); const [err,setErr] = useState("");
  const [np,setNp]     = useState(""); const [cp,setCp]   = useState("");
  const [shake,setShk] = useState(false);
  const check = () => {
    if(pin===savedPin){ onVerify(); }
    else { setErr("PIN salah. Coba lagi."); setShk(true); setTimeout(()=>{setShk(false);setErr("");setPin("");},1500); }
  };
  const change = () => {
    if(np.length<4){setErr("Minimal 4 digit");return;}
    if(np!==cp){setErr("PIN tidak cocok");return;}
    onVerify(np);
  };
  return (
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"linear-gradient(135deg,#003D33,#00695C,#00897B)",padding:16}}>
      <style>{`@keyframes shk{0%,100%{transform:translateX(0)}25%,75%{transform:translateX(-7px)}50%{transform:translateX(7px)}} @keyframes lgFloat{0%,100%{transform:translateY(0) scale(1)}50%{transform:translateY(-8px) scale(1.06)}} @keyframes lgGlow{0%,100%{box-shadow:0 4px 18px rgba(0,77,64,.3),0 0 0 3px #00695C55}50%{box-shadow:0 12px 38px rgba(0,77,64,.65),0 0 0 4px #00897Baa}}`}</style>
      <div style={{background:C.white,borderRadius:24,padding:"32px 28px",width:"100%",maxWidth:340,boxShadow:"0 20px 60px rgba(0,0,0,.3)",animation:shake?"shk .4s":"none"}}>
        <div style={{textAlign:"center",marginBottom:24}}>
          <img src="/logo.jpeg" alt="Logo Kamar Bedah" style={{width:104,height:104,borderRadius:"50%",objectFit:"cover",display:"block",margin:"0 auto 14px",animation:"lgFloat 3s ease-in-out infinite, lgGlow 3s ease-in-out infinite"}}/>
          <div style={{fontSize:17,fontWeight:800,color:C.p}}>Sistem Koordinasi<br/>Kamar Bedah</div>
          <div style={{fontSize:12,color:C.tL,marginTop:6}}>{HOSPITAL}</div>
        </div>
        {mode==="login" && <>
          <div style={{fontSize:13,fontWeight:600,color:C.t,textAlign:"center",marginBottom:10}}>Masukkan PIN Akses</div>
          <input style={{...iS,textAlign:"center",fontSize:28,letterSpacing:12,height:56,marginBottom:12,borderRadius:12}} type="password" inputMode="numeric" maxLength={6} placeholder="••••" value={pin} onChange={e=>setPin(e.target.value)} onKeyDown={e=>e.key==="Enter"&&check()}/>
          {err && <div style={{fontSize:12,color:C.d,textAlign:"center",marginBottom:10}}>⚠ {err}</div>}
          <Btn full onClick={check} style={{marginBottom:12,padding:"13px",fontSize:15}}>Masuk</Btn>
          <div style={{display:"flex",justifyContent:"space-between"}}>
            <button onClick={()=>{setMode("change");setErr("");}} style={{background:"none",border:"none",color:C.tL,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>Ubah PIN</button>
            <button onClick={()=>{setMode("forgot");setErr("");}} style={{background:"none",border:"none",color:C.i,fontSize:12,cursor:"pointer",fontFamily:"inherit",fontWeight:700}}>Lupa PIN?</button>
          </div>
        </>}
        {mode==="change" && <>
          <div style={{fontSize:14,fontWeight:700,color:C.p,marginBottom:14,textAlign:"center"}}>Ubah PIN Akses</div>
          <LF label="PIN Baru (min. 4 digit)"><input style={iS} type="password" inputMode="numeric" maxLength={6} placeholder="PIN baru" value={np} onChange={e=>setNp(e.target.value)}/></LF>
          <LF label="Konfirmasi PIN"><input style={iS} type="password" inputMode="numeric" maxLength={6} placeholder="Ulangi PIN" value={cp} onChange={e=>setCp(e.target.value)}/></LF>
          {err && <div style={{fontSize:12,color:C.d,marginBottom:10}}>⚠ {err}</div>}
          <div style={{display:"flex",gap:8}}><Btn full onClick={change}>Simpan PIN</Btn><Btn full outline color={C.g} onClick={()=>setMode("login")}>Batal</Btn></div>
        </>}
        {mode==="forgot" && <>
          <div style={{fontSize:13,color:C.tL,marginBottom:14,textAlign:"center",lineHeight:1.6}}>Untuk reset PIN, gunakan PIN darurat <b style={{color:C.p}}>1234</b> atau hubungi Administrator Sistem.</div>
          <Btn full onClick={()=>setMode("login")}>Kembali ke Login</Btn>
        </>}
      </div>
    </div>
  );
}

/* ─── VIEW JADWAL ────────────────────────────────────────────────────── */
function ViewJadwal({ops,setOps,startEditOp,deleteOp,sendReminder,reqOpId,setReqOpId,reqText,setReqText,addReq,delReq,getPhone,setNotifs,showToast,privacyMode}: any) {
  const [cId,setCId]     = useState<string|null>(null);
  const [cR,setCR]       = useState("");
  const [cytoM,setCytoM] = useState<any>(null);
  const [delC,setDelC]   = useState<string|null>(null);
  const [page,setPage]   = useState(0);
  const [fSt,setFSt]     = useState("all");
  const [q,setQ]         = useState("");

  const sorted = [...ops]
    .filter((op:any) => {
      if(fSt!=="all"&&op.status!==fSt) return false;
      if(q.trim()){ const lq=q.toLowerCase(); return (op.patient||"").toLowerCase().includes(lq)||(op.procedure||"").toLowerCase().includes(lq)||(op.surgeon||"").toLowerCase().includes(lq); }
      return true;
    })
    .sort((a:any,b:any)=>{
      const at=a.date===todayDate(),bt=b.date===todayDate();
      if(at&&a.opType==="cyto"&&!(bt&&b.opType==="cyto")) return -1;
      if(bt&&b.opType==="cyto"&&!(at&&a.opType==="cyto")) return 1;
      return a.date.localeCompare(b.date)||a.time.localeCompare(b.time);
    });
  const total  = Math.ceil(sorted.length/PAGE_SIZE);
  const paged  = sorted.slice(page*PAGE_SIZE,(page+1)*PAGE_SIZE);
  const hF = (v: string)=>{ setFSt(v); setPage(0); };
  const hQ = (v: string)=>{ setQ(v);   setPage(0); };

  const cytoActive = ops.filter((o:any)=>o.opType==="cyto"&&o.status!=="batal"&&o.status!=="done");
  const openCyto   = (op:any) => { const ph=getPhone(op.surgeon); const msg=msgCyto(op); if(!ph){showToast("Nomor WA Dokter Bedah belum terdaftar di Staf",C.d);return;} setCytoM({op,msg,ph,ts:fNow()}); };
  const confirmCyto= () => {
    if(!cytoM)return;
    window.open(`https://wa.me/${cytoM.ph.replace(/[^0-9]/g,"")}?text=${encodeURIComponent(cytoM.msg)}`,"_blank");
    setNotifs((p:any)=>[{id:gId(),type:"cyto",label:"⚡ WA Cyto → "+cytoM.op.surgeon,patient:cytoM.op.patient,procedure:cytoM.op.procedure,message:cytoM.msg,sentAt:cytoM.ts},...p]);
    showToast("✓ WhatsApp dibuka — tekan Kirim di WA","#25D366");
    setCytoM(null);
  };

  return (
    <div>
      {privacyMode && <div style={{background:"#4A148C",borderRadius:10,padding:"8px 14px",marginBottom:8,fontSize:12,fontWeight:700,color:"#fff"}}>🔒 Mode Privasi — Nama & RM disamarkan</div>}
      {cytoActive.length>0 && (
        <div style={{background:"#FFCDD2",borderRadius:10,padding:"10px 14px",marginBottom:12,border:"2px solid #F44336"}}>
          <div style={{fontSize:13,fontWeight:800,color:"#B71C1C",marginBottom:4}}>⚠️ OPERASI CYTO / EMERGENCY AKTIF</div>
          {cytoActive.map((o:any)=><div key={o.id} style={{fontSize:12,color:"#B71C1C"}}>• {maskName(o.patient,privacyMode)} — {o.procedure} ({o.time} WIB)</div>)}
        </div>
      )}
      <div style={{marginBottom:12}}>
        <div style={{display:"flex",gap:8,marginBottom:8}}>
          <input style={{...iS,flex:1,padding:"9px 14px",fontSize:13}} placeholder="🔍 Cari pasien, tindakan, dokter..." value={q} onChange={e=>hQ(e.target.value)}/>
          {q&&<button onClick={()=>hQ("")} style={{background:"none",border:`1px solid ${C.b}`,borderRadius:8,color:C.g,padding:"0 14px",cursor:"pointer",fontSize:16}}>✕</button>}
        </div>
        <div style={{display:"flex",gap:5,flexWrap:"wrap",alignItems:"center"}}>
          {[{v:"all",l:"Semua"},{v:"scheduled",l:"Terjadwal"},{v:"ongoing",l:"Berlangsung"},{v:"done",l:"Selesai"},{v:"batal",l:"Batal"}].map(f=>(
            <button key={f.v} onClick={()=>hF(f.v)} style={{padding:"5px 12px",borderRadius:20,border:`1px solid ${fSt===f.v?C.p:C.b}`,background:fSt===f.v?C.p:"#FAFBFD",color:fSt===f.v?"#fff":C.g,fontSize:12,fontWeight:fSt===f.v?700:400,cursor:"pointer",fontFamily:"inherit"}}>
              {f.l}
            </button>
          ))}
          <span style={{marginLeft:"auto",fontSize:12,color:C.tL}}>{sorted.length} jadwal</span>
          <button onClick={()=>{
            const printDate = prompt("Cetak jadwal tanggal (YYYY-MM-DD, kosongkan untuk semua):");
            const filtered = printDate ? ops.filter((o:any)=>o.date===printDate) : sorted;
            if(!filtered.length){showToast("Tidak ada jadwal untuk dicetak",C.w);return;}
            const rows = filtered.map((op:any,i:number)=>`<tr><td>${i+1}</td><td>${op.date} ${op.time}</td><td><b>${op.patient}</b></td><td>${op.procedure}</td><td>${op.room}</td><td>${op.surgeon}</td><td>${op.anesthesiologist||"—"}</td><td style="text-align:center"><span style="background:${STS[op.status as keyof typeof STS]?.c||C.g};color:#fff;padding:2px 8px;border-radius:10px;font-size:10px">${STS[op.status as keyof typeof STS]?.l||op.status}</span></td></tr>`).join("");
            const w=window.open("","_blank","width=900,height=700")!;
            w.document.write(`<!DOCTYPE html><html><head><title>Jadwal Operasi ${printDate||"Semua"} — ${HOSPITAL}</title><style>body{font-family:Arial,sans-serif;font-size:12px;margin:20px}h2{color:#004D40;margin-bottom:4px}p{color:#555;margin-bottom:12px}table{width:100%;border-collapse:collapse}th{background:#004D40;color:#fff;padding:7px 10px;text-align:left;font-size:11px}td{padding:6px 10px;border-bottom:1px solid #e0e0e0;vertical-align:top}tr:nth-child(even)td{background:#f5f5f5}@media print{button{display:none}}</style></head><body><h2>🏥 Jadwal Operasi — ${HOSPITAL}</h2><p>${printDate?`Tanggal: ${printDate}`:"Semua jadwal"} · Dicetak: ${fNow()}</p><button onclick="window.print()" style="margin-bottom:12px;background:#004D40;color:#fff;border:none;padding:8px 18px;border-radius:6px;cursor:pointer;font-size:13px">🖨 Cetak / Simpan PDF</button><table><thead><tr><th>#</th><th>Tanggal/Jam</th><th>Pasien</th><th>Tindakan</th><th>Kamar</th><th>Dr. Bedah</th><th>Dr. Anestesi</th><th>Status</th></tr></thead><tbody>${rows}</tbody></table></body></html>`);
            w.document.close();
          }} style={{padding:"5px 12px",borderRadius:20,border:`1px solid ${C.p}`,background:C.pBg,color:C.p,fontSize:12,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>
            🖨 Cetak
          </button>
        </div>
      </div>
      {sorted.length===0 && (
        <Card><div style={{textAlign:"center",padding:28,color:C.tL}}>
          <div style={{fontSize:36,marginBottom:8}}>📋</div>
          <div style={{fontSize:14,fontWeight:600}}>{ops.length===0?"Belum ada jadwal terdaftar":"Tidak ada hasil pencarian"}</div>
          {ops.length===0&&<div style={{fontSize:12,marginTop:4}}>Klik tab 📝 Daftar untuk mendaftarkan</div>}
        </div></Card>
      )}
      {paged.map((op:any)=>{
        const sc=STS[op.status as keyof typeof STS]||STS.scheduled, ot=OT[op.opType as keyof typeof OT||"elektif"];
        const hasAlg=op.allergy&&op.allergy!=="Tidak Ada", isBatal=op.status==="batal";
        return (
          <Card key={op.id} hi={op.opType==="cyto"&&!isBatal?"#F44336":hasAlg&&!isBatal?C.dL:undefined} style={{opacity:isBatal?.72:1}}>
            {op.opType==="cyto"&&!isBatal && <div style={{background:"#FFCDD2",borderRadius:8,padding:"6px 12px",marginBottom:8,fontSize:12,fontWeight:700,color:"#B71C1C"}}>⚠️ CYTO / EMERGENCY — SEGERA DITANGANI</div>}
            {hasAlg&&op.opType!=="cyto"&&!isBatal && <div style={{background:C.dBg,borderRadius:8,padding:"5px 12px",marginBottom:8,fontSize:12,fontWeight:700,color:C.d}}>⚠ ALERGI: {op.allergy}</div>}
            {isBatal && <div style={{background:C.dBg,borderRadius:8,padding:"5px 12px",marginBottom:8,fontSize:12,fontWeight:700,color:C.d}}>✕ {op.cancelReason?"BATAL: "+op.cancelReason:"BATAL/DITUNDA"}</div>}
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
              <div style={{flex:1,marginRight:8}}>
                <div style={{fontSize:15,fontWeight:800,color:C.t,textDecoration:isBatal?"line-through":"none"}}>{maskName(op.patient,privacyMode)||"—"}</div>
                <div style={{fontSize:12,color:C.tL,marginTop:2}}>{op.age?op.age+"th  ":""}{op.rm?"RM "+maskRM(op.rm,privacyMode)+"  ":""}{op.bloodType}</div>
              </div>
              <div style={{display:"flex",flexDirection:"column",gap:4,alignItems:"flex-end"}}>
                <Bdg label={ot?.label||"Elektif"} color={ot?.c||C.i} bg={ot?.bg||C.iBg}/>
                <Bdg label={sc.l} color={sc.c} bg={sc.bg}/>
              </div>
            </div>
            <div style={{background:"#F8FBF0",border:"1px solid #D4E6A0",borderRadius:10,padding:"10px 12px",marginBottom:8}}>
              <div style={{fontSize:14,fontWeight:700,color:"#33691E"}}>{op.procedure}</div>
              <div style={{fontSize:12,color:"#558B2F",marginTop:2}}>{op.diagnosis}</div>
              <div style={{fontSize:12,color:"#558B2F",marginTop:3}}>📅 {fD(op.date)} · {op.time} WIB · {op.room}</div>
            </div>
            <div style={{fontSize:12,color:C.tL,lineHeight:1.9,marginBottom:8}}>
              <div><b>Operator:</b> {op.surgeon||"—"}</div>
              <div><b>Anestesi:</b> {op.anesthesiologist||"—"}</div>
              <div><b>Asisten:</b> {op.assistantNurse||"—"} · <b>Instrumen:</b> {op.circulatingNurse||"—"}</div>
              <div><b>Onloop:</b> {op.onloopNurse||"—"} · <b>RR:</b> {op.rrKatim||"—"}</div>
            </div>
            {op.specialNeeds && <div style={{background:C.wBg,borderRadius:8,padding:"6px 12px",marginBottom:8,fontSize:12,color:C.w}}>📌 {op.specialNeeds}</div>}
            {op.ruangAsal&&op.opType==="cyto" && <div style={{background:"#FFCDD2",borderRadius:8,padding:"5px 12px",marginBottom:8,fontSize:12,color:"#B71C1C"}}>📍 Ruang Asal: {op.ruangAsal}</div>}
            {(op.requests||[]).length>0 && (
              <div style={{background:C.wBg,borderRadius:8,padding:"8px 12px",marginBottom:8}}>
                <div style={{fontSize:12,fontWeight:700,color:C.w,marginBottom:3}}>PERMINTAAN OPERATOR</div>
                {op.requests.map((r:any)=>(
                  <div key={r.id||r} style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                    <div style={{fontSize:12,color:C.w}}>• {typeof r==="string"?r:r.text}</div>
                    {typeof r!=="string" && <button onClick={()=>delReq(op.id,r.id)} style={{background:"none",border:"none",color:C.d,cursor:"pointer",fontSize:16,padding:"0 4px"}}>×</button>}
                  </div>
                ))}
              </div>
            )}
            {(op.reminders||[]).length>0 && <div style={{marginBottom:8}}>{op.reminders.map((r:string,i:number)=><span key={i} style={{background:C.sBg,color:C.s,fontSize:11,padding:"3px 10px",borderRadius:10,fontWeight:700,marginRight:4}}>✓ {r==="H-1"?"H-1":"1 Jam"}</span>)}</div>}
            {!isBatal && (
              <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:6}}>
                {op.opType==="cyto"
                  ? <WaBt sm onClick={()=>openCyto(op)}>⚡ WA Cyto</WaBt>
                  : <><Btn sm outline color={C.i} disabled={(op.reminders||[]).includes("H-1")} onClick={()=>sendReminder(op,"H-1")}>H-1</Btn><Btn sm outline color={C.w} disabled={(op.reminders||[]).includes("H-1jam")} onClick={()=>sendReminder(op,"H-1jam")}>1 Jam</Btn></>
                }
                <Btn sm outline color={C.p} onClick={()=>startEditOp(op)}>✏ Edit</Btn>
                <Btn sm outline color={C.w} onClick={()=>{setReqOpId(op.id);setReqText("");}}>+ Permintaan</Btn>
              </div>
            )}
            <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
              {!isBatal&&op.status==="scheduled" && <Btn sm outline color={C.pL} onClick={()=>setOps((p:any)=>p.map((o:any)=>o.id===op.id?{...o,status:"ongoing"}:o))}>Mulai</Btn>}
              {!isBatal&&op.status==="ongoing"    && <Btn sm outline color={C.s} onClick={()=>setOps((p:any)=>p.map((o:any)=>o.id===op.id?{...o,status:"done"}:o))}>Selesai</Btn>}
              {!isBatal && <Btn sm outline color={C.d} onClick={()=>{setCId(op.id);setCR("");}}>Batal/Tunda</Btn>}
              {isBatal   && <Btn sm outline color={C.g} onClick={()=>setOps((p:any)=>p.map((o:any)=>o.id===op.id?{...o,status:"scheduled",cancelReason:""}:o))}>Aktifkan</Btn>}
              <Btn sm outline color={C.d} onClick={()=>setDelC(op.id)}>Hapus</Btn>
            </div>
            {cId===op.id && (
              <div style={{marginTop:12,paddingTop:12,borderTop:`1px solid ${C.b}`}}>
                <div style={{fontSize:13,fontWeight:700,color:C.d,marginBottom:8}}>Alasan Pembatalan / Penundaan</div>
                <input style={{...iS,marginBottom:8}} placeholder="Cth: Pasien belum puasa, kondisi tidak stabil" value={cR} onChange={e=>setCR(e.target.value)}/>
                <div style={{display:"flex",gap:8}}><Btn full color={C.d} onClick={()=>{setOps((p:any)=>p.map((o:any)=>o.id===op.id?{...o,status:"batal",cancelReason:sanitize(cR)}:o));setCId(null);setCR("");showToast("Operasi ditandai batal",C.w);}}>Konfirmasi Batal</Btn><Btn full outline color={C.g} onClick={()=>setCId(null)}>Tutup</Btn></div>
              </div>
            )}
            {reqOpId===op.id && (
              <div style={{marginTop:12,paddingTop:12,borderTop:`1px solid ${C.b}`}}>
                <div style={{fontSize:13,fontWeight:700,color:C.t,marginBottom:8}}>Tambah Permintaan Operator</div>
                <textarea value={reqText} onChange={(e:any)=>setReqText(e.target.value)} placeholder="Cth: Siapkan retractor khusus, warming blanket..." style={{...iS,resize:"vertical",minHeight:68,marginBottom:8} as React.CSSProperties}/>
                <div style={{display:"flex",gap:8}}><Btn full onClick={()=>addReq(op.id)}>Simpan</Btn><Btn full outline color={C.g} onClick={()=>setReqOpId(null)}>Batal</Btn></div>
              </div>
            )}
          </Card>
        );
      })}
      {total>1 && (
        <div style={{display:"flex",justifyContent:"center",alignItems:"center",gap:8,padding:"12px 0 6px"}}>
          <button onClick={()=>setPage((p:number)=>Math.max(0,p-1))} disabled={page===0} style={{padding:"8px 16px",borderRadius:10,border:`1.5px solid ${page===0?C.b:C.p}`,background:"#FAFBFD",cursor:page===0?"not-allowed":"pointer",color:page===0?C.b:C.p,fontSize:13,fontWeight:700,fontFamily:"inherit",opacity:page===0?.4:1}}>← Prev</button>
          {Array.from({length:total},(_,i)=>(
            <button key={i} onClick={()=>setPage(i)} style={{width:36,height:36,borderRadius:10,border:`1.5px solid ${i===page?C.p:C.b}`,background:i===page?C.p:"#FAFBFD",color:i===page?"#fff":C.t,cursor:"pointer",fontSize:13,fontWeight:i===page?700:400,fontFamily:"inherit"}}>{i+1}</button>
          ))}
          <button onClick={()=>setPage((p:number)=>Math.min(total-1,p+1))} disabled={page===total-1} style={{padding:"8px 16px",borderRadius:10,border:`1.5px solid ${page===total-1?C.b:C.p}`,background:"#FAFBFD",cursor:page===total-1?"not-allowed":"pointer",color:page===total-1?C.b:C.p,fontSize:13,fontWeight:700,fontFamily:"inherit",opacity:page===total-1?.4:1}}>Next →</button>
        </div>
      )}
      {delC && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.55)",zIndex:400,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
          <div style={{background:C.white,borderRadius:18,padding:"24px 20px",maxWidth:320,width:"100%",boxShadow:"0 12px 40px rgba(0,0,0,.25)"}}>
            <div style={{textAlign:"center",marginBottom:14}}><div style={{fontSize:36}}>🗑</div><div style={{fontSize:15,fontWeight:700,color:C.d,marginTop:6}}>Hapus Jadwal?</div><div style={{fontSize:12,color:C.tL,marginTop:4}}>Tindakan ini tidak dapat dibatalkan.</div></div>
            <div style={{display:"flex",gap:8}}><Btn full outline color={C.g} onClick={()=>setDelC(null)} style={{flex:1}}>Batal</Btn><Btn full color={C.d} onClick={()=>{deleteOp(delC);setDelC(null);}} style={{flex:1}}>Hapus</Btn></div>
          </div>
        </div>
      )}
      {cytoM && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.65)",zIndex:500,display:"flex",alignItems:"flex-end",justifyContent:"center"}}>
          <div style={{background:C.white,borderRadius:"22px 22px 0 0",padding:"22px 20px",width:"100%",maxWidth:520,maxHeight:"88vh",display:"flex",flexDirection:"column",boxShadow:"0 -10px 40px rgba(0,0,0,.3)"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
              <div><div style={{fontSize:15,fontWeight:800,color:"#B71C1C"}}>⚡ Pratinjau WA Cyto / Emergency</div><div style={{fontSize:11,color:C.tL,marginTop:2}}>🕐 {cytoM.ts}</div></div>
              <button onClick={()=>setCytoM(null)} style={{background:"none",border:`1px solid ${C.b}`,borderRadius:8,padding:"4px 12px",cursor:"pointer",fontSize:20,lineHeight:1,flexShrink:0}}>×</button>
            </div>
            <div style={{background:"#FFF3E0",borderRadius:8,padding:"8px 14px",marginBottom:10,fontSize:12,color:C.w,fontWeight:600}}>📨 Penerima: <b>{cytoM.op.surgeon}</b> (+{cytoM.ph})</div>
            <div style={{flex:1,overflowY:"auto",background:"#FFF8F8",borderRadius:8,padding:"12px 14px",fontSize:12,lineHeight:1.9,whiteSpace:"pre-wrap",border:"2px solid #F44336",fontFamily:"inherit",marginBottom:14}}>{cytoM.msg}</div>
            <div style={{fontSize:11,color:"#B71C1C",marginBottom:12,fontWeight:600,textAlign:"center"}}>⚠ Periksa isi pesan. WhatsApp akan terbuka — tekan Kirim secara manual.</div>
            <div style={{display:"flex",gap:10}}><Btn full outline color={C.g} onClick={()=>setCytoM(null)} style={{flex:1}}>Batal</Btn><WaBt full onClick={confirmCyto} style={{flex:2}}>⚡ Konfirmasi & Buka WA</WaBt></div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ─── VIEW DAFTAR ────────────────────────────────────────────────────── */
function ViewDaftar({opForm,setOpForm,editingOp,resetOp,saveOp,opErrors,setOpErrors,staff,setTab,dupWarning,templates,setTemplates,showToast}: any) {
  const byT = (t: string) => staff.filter((s:any)=>s.type===t).map((s:any)=>s.name);
  const allNurses = [...new Set([...byT("circulating"),...byT("anesthesia_nurse"),...byT("onloop"),...byT("katim")])];
  const [foc,setFoc] = useState<string|null>(null);
  const [showTpl,setShowTpl] = useState(false);
  const set  = (k: string,v: any) => { setOpForm((p: any)=>({...p,[k]:v})); setOpErrors((p: any)=>({...p,[k]:undefined})); };
  const setR = (k: string,v: any) => { setOpForm((p: any)=>({...p,[k]:v})); setOpErrors((p: any)=>({...p,[k]:undefined})); };
  const e = opErrors;
  const iStyle = (k: string): React.CSSProperties => ({...iS, borderColor:e[k]?C.d:foc===k?C.p:C.b, boxShadow:foc===k?`0 0 0 3px ${C.p}22`:e[k]?`0 0 0 3px ${C.d}22`:"none"});
  const onFoc = (k: string) => () => setFoc(k);
  const onBlr = () => setFoc(null);

  const applyTemplate = (t: any) => {
    setOpForm((p: any)=>({...p, opType:t.opType||"elektif", procedure:t.procedure||"", diagnosis:t.diagnosis||"", surgeon:t.surgeon||"", anesthesiologist:t.anesthesiologist||"", room:t.room||"", time:t.time||p.time, circulatingNurse:t.circulatingNurse||"", anesthesiaNurse:t.anesthesiaNurse||"", onloopNurse:t.onloopNurse||"", rrKatim:t.rrKatim||""}));
    setShowTpl(false);
    showToast("✓ Template diterapkan",C.s);
  };
  const saveAsTemplate = () => {
    const name = prompt("Nama template:");
    if(!name) return;
    const t = {id:gId(), name, opType:opForm.opType, procedure:opForm.procedure, diagnosis:opForm.diagnosis, surgeon:opForm.surgeon, anesthesiologist:opForm.anesthesiologist, room:opForm.room, time:opForm.time, circulatingNurse:opForm.circulatingNurse, anesthesiaNurse:opForm.anesthesiaNurse, onloopNurse:opForm.onloopNurse, rrKatim:opForm.rrKatim};
    setTemplates((p: any[])=>[...p,t]);
    showToast("✓ Template disimpan: "+name,C.s);
  };
  const deleteTemplate = (id: string) => setTemplates((p: any[])=>p.filter((t: any)=>t.id!==id));

  return (
    <div>
      <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:18}}>
        {editingOp && <button onClick={()=>{resetOp();setTab("jadwal");}} style={{background:C.pBg,border:`1px solid ${C.p}`,borderRadius:10,color:C.p,cursor:"pointer",fontSize:22,width:40,height:40,display:"flex",alignItems:"center",justifyContent:"center"}}>←</button>}
        <div><div style={{fontSize:17,fontWeight:800,color:C.t}}>{editingOp?"✏️ Edit Jadwal":"📝 Pendaftaran Jadwal Operasi"}</div><div style={{fontSize:12,color:C.tL,marginTop:2}}>Field bertanda <span style={{color:C.d,fontWeight:700}}>✱</span> wajib diisi</div></div>
        <button onClick={()=>setShowTpl(s=>!s)} style={{marginLeft:"auto",background:C.pBg,border:`1px solid ${C.p}33`,borderRadius:10,color:C.p,fontSize:11,fontWeight:700,padding:"7px 12px",cursor:"pointer",fontFamily:"inherit",whiteSpace:"nowrap"}}>📌 Template{templates.length>0?` (${templates.length})`:""}</button>
      </div>
      {showTpl && (
        <Card style={{marginBottom:14,background:"#F0FFF8",border:`1px solid ${C.p}33`}}>
          <div style={{fontSize:13,fontWeight:700,color:C.p,marginBottom:10}}>📌 Pilih Template Cepat</div>
          {templates.length===0 && <div style={{fontSize:12,color:C.tL,marginBottom:8}}>Belum ada template. Isi form lalu klik "Simpan sebagai Template".</div>}
          {templates.map((t: any)=>(
            <div key={t.id} style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
              <button onClick={()=>applyTemplate(t)} style={{flex:1,textAlign:"left",background:C.white,border:`1px solid ${C.b}`,borderRadius:8,padding:"8px 12px",cursor:"pointer",fontFamily:"inherit"}}>
                <div style={{fontSize:12,fontWeight:700,color:C.p}}>{t.name}</div>
                <div style={{fontSize:11,color:C.tL}}>{t.procedure} · Dr. {t.surgeon}</div>
              </button>
              <button onClick={()=>deleteTemplate(t.id)} style={{background:"none",border:`1px solid ${C.dL}`,borderRadius:8,color:C.d,padding:"6px 10px",cursor:"pointer",fontSize:13}}>🗑</button>
            </div>
          ))}
          <Btn full outline color={C.p} onClick={saveAsTemplate} style={{marginTop:8}}>💾 Simpan Form Ini sebagai Template</Btn>
        </Card>
      )}
      {dupWarning && <div style={{background:C.wBg,borderRadius:10,padding:"10px 14px",marginBottom:14,fontSize:13,color:C.w,fontWeight:600,border:`1.5px solid ${C.w}`}}>⚠ Duplikasi: pasien, tanggal & jam yang sama sudah terdaftar.</div>}
      <Card>
        <SH label="① Jenis Operasi & Tindakan"/>
        <LF label="Jenis Operasi">
          <div style={{display:"flex",gap:8}}>
            {Object.entries(OT).map(([v,o])=>(
              <button key={v} onClick={()=>setR("opType",v)} style={{flex:1,padding:"10px 6px",borderRadius:10,border:`2px solid ${opForm.opType===v?o.c:C.b}`,background:opForm.opType===v?o.bg:"#FAFBFD",color:opForm.opType===v?o.c:C.g,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
                {o.label}
              </button>
            ))}
          </div>
        </LF>
        {opForm.opType==="cyto" && <LF label="Ruang Asal Pasien"><input style={iStyle("ruangAsal")} placeholder="IGD, ICU, Bangsal..." value={opForm.ruangAsal||""} onChange={ev=>set("ruangAsal",ev.target.value)} onFocus={onFoc("ruangAsal")} onBlur={onBlr} autoComplete="off" spellCheck={false}/></LF>}
        <LF label="Diagnosa Medis" req err={e.diagnosis}><input style={iStyle("diagnosis")} placeholder="Cth: Hernia Inguinalis Dextra, Appendicitis Akut" value={opForm.diagnosis} onChange={(ev:any)=>set("diagnosis",ev.target.value)} onFocus={onFoc("diagnosis")} onBlur={onBlr} autoComplete="off" spellCheck={false}/></LF>
        <LF label="Nama Tindakan / Prosedur" req err={e.procedure}><input style={iStyle("procedure")} placeholder="Cth: Hernioraphy Open, Laparoscopic Cholecystectomy" value={opForm.procedure} onChange={(ev:any)=>set("procedure",ev.target.value)} onFocus={onFoc("procedure")} onBlur={onBlr} autoComplete="off" spellCheck={false}/></LF>
        <div style={{display:"flex",gap:12}}>
          <div style={{flex:1}}><LF label="Tanggal" req><input style={iStyle("date")} type="date" value={opForm.date} onChange={(ev:any)=>setR("date",ev.target.value)}/></LF>{e.date&&<div style={{fontSize:11,color:C.d,marginTop:-8,marginBottom:8}}>⚠ {e.date}</div>}</div>
          <div style={{flex:1}}><LF label="Jam Mulai" req><input style={iStyle("time")} type="time" value={opForm.time} onChange={(ev:any)=>setR("time",ev.target.value)}/></LF>{e.time&&<div style={{fontSize:11,color:C.d,marginTop:-8,marginBottom:8}}>⚠ {e.time}</div>}</div>
        </div>
        <SF label="Kamar Operasi" options={ROOMS} value={opForm.room} onChange={(ev:any)=>setR("room",ev.target.value)}/>
      </Card>
      <Card>
        <SH label="② Data Pasien"/>
        <LF label="Nama Lengkap Pasien" req err={e.patient}><input style={iStyle("patient")} placeholder="Nama sesuai rekam medis (Tn./Ny./An.)" value={opForm.patient} onChange={(ev:any)=>set("patient",ev.target.value)} onFocus={onFoc("patient")} onBlur={onBlr} autoComplete="off" spellCheck={false}/></LF>
        <div style={{display:"flex",gap:12}}>
          <div style={{flex:1}}><LF label="Usia (tahun)"><input style={iStyle("age")} type="number" min="0" max="150" placeholder="45" value={opForm.age} onChange={(ev:any)=>setR("age",ev.target.value)}/></LF></div>
          <div style={{flex:1}}><LF label="No. Rekam Medis"><input style={iStyle("rm")} placeholder="PRT-004891" value={opForm.rm} onChange={(ev:any)=>setR("rm",ev.target.value)} autoComplete="off"/></LF></div>
        </div>
        <SF label="Golongan Darah" options={BT} value={opForm.bloodType} onChange={(ev:any)=>setR("bloodType",ev.target.value)}/>
        <LF label="Riwayat Alergi"><input style={iStyle("allergy")} placeholder="Nama obat/zat alergi, atau 'Tidak Ada'" value={opForm.allergy} onChange={(ev:any)=>set("allergy",ev.target.value)} onFocus={onFoc("allergy")} onBlur={onBlr} autoComplete="off" spellCheck={false}/></LF>
        <LF label="Kebutuhan Khusus / Catatan"><textarea style={{...iStyle("specialNeeds"),resize:"vertical",minHeight:80,lineHeight:1.7} as React.CSSProperties} placeholder={"Cth:\n• Perlu warming blanket\n• Posisi litotomi\n• Siapkan darah 2 kantong"} value={opForm.specialNeeds||""} onChange={(ev:any)=>set("specialNeeds",ev.target.value)} onFocus={onFoc("specialNeeds")} onBlur={onBlr} spellCheck={false}/></LF>
      </Card>
      <Card>
        <SH label="③ Tim Dokter"/>
        <div style={{background:C.iBg,borderRadius:8,padding:"8px 12px",marginBottom:14,fontSize:12,color:C.i}}>Belum ada dokter? Tambahkan di menu <b>👥 Staf</b></div>
        <SF label="Dokter Operator / Bedah" req options={byT("surgeon")} empty="-- Pilih Dokter Bedah --" value={opForm.surgeon} onChange={(ev:any)=>setR("surgeon",ev.target.value)} err={e.surgeon}/>
        <SF label="Dokter Anestesi" options={byT("anesthesiologist")} empty="-- Pilih Dokter Anestesi --" value={opForm.anesthesiologist} onChange={(ev:any)=>setR("anesthesiologist",ev.target.value)}/>
      </Card>
      <Card>
        <SH label="④ Tim Perawat (5 Peran Wajib)"/>
        <div style={{background:C.sBg,borderRadius:8,padding:"8px 12px",marginBottom:14,fontSize:12,color:C.s}}>✓ Setiap jadwal memerlukan 5 peran perawat sesuai standar Kamar Bedah</div>
        {/* Asisten */}
        <LF label="1. Perawat Asisten Operator">
          <div style={{display:"flex",gap:8}}>
            <div style={{flex:1,position:"relative"}}>
              <select style={{...iStyle("assistantNurse"),appearance:"none",paddingRight:30,cursor:"pointer"} as React.CSSProperties} value={opForm.assistantNurse||""} onChange={(ev:any)=>setR("assistantNurse",ev.target.value)}>
                <option value="">-- Pilih dari staf --</option>
                {(allNurses as string[]).map((n,i)=><option key={i} value={n}>{n}</option>)}
              </select>
              <div style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",pointerEvents:"none",color:C.g,fontSize:11}}>▼</div>
            </div>
            <input style={{...iStyle("assistantNurse"),flex:1}} placeholder="Atau ketik manual..." value={opForm.assistantNurse||""} onChange={(ev:any)=>set("assistantNurse",ev.target.value)} onFocus={onFoc("assistantNurse")} onBlur={onBlr} autoComplete="off" spellCheck={false}/>
          </div>
          <div style={{fontSize:11,color:C.tL,marginTop:3}}>Pilih dari dropdown atau ketik nama manual</div>
        </LF>
        {/* PERAWAT Instrumen — mengganti Perawat Sirkuler */}
        <LF label="2. PERAWAT Instrumen">
          <div style={{position:"relative"}}>
            <select style={{...iStyle("circulatingNurse"),appearance:"none",paddingRight:30,cursor:"pointer"} as React.CSSProperties} value={opForm.circulatingNurse||""} onChange={(ev:any)=>setR("circulatingNurse",ev.target.value)}>
              <option value="">-- Pilih dari daftar staf --</option>
              {(allNurses as string[]).map((n,i)=><option key={i} value={n}>{n}</option>)}
            </select>
            <div style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",pointerEvents:"none",color:C.g,fontSize:11}}>▼</div>
          </div>
        </LF>
        {/* Perawat Onloop — bisa diisi 2–3 orang */}
        <LF label="3. Perawat Onloop (dapat diisi 1–3 orang sekaligus)">
          <div style={{display:"flex",flexDirection:"column",gap:8}}>
            {[0,1,2].map((idx)=>{
              const parts=(opForm.onloopNurse||"").split(" | ").map((s:string)=>s.trim()).concat(["","",""]);
              const val=parts[idx]||"";
              const setSlot=(newVal:string)=>{
                const arr=(opForm.onloopNurse||"").split(" | ").map((s:string)=>s.trim()).concat(["","",""]).slice(0,3);
                arr[idx]=newVal;
                setR("onloopNurse",arr.filter(Boolean).join(" | "));
              };
              const taken=(opForm.onloopNurse||"").split(" | ").map((s:string)=>s.trim()).filter((_:string,i:number)=>i!==idx);
              return (
                <div key={idx} style={{display:"flex",gap:8,alignItems:"center"}}>
                  <div style={{width:26,height:26,borderRadius:13,background:idx===0?C.p:val?C.pL:"#DDE3EA",color:idx===0||val?"#fff":C.g,fontSize:12,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,transition:"background .2s"}}>{idx+1}</div>
                  <div style={{flex:1,position:"relative"}}>
                    <select style={{...iStyle("onloopNurse"),appearance:"none",paddingRight:30,cursor:"pointer",background:val?"#F0FFF8":"#FAFBFD"} as React.CSSProperties} value={val} onChange={(ev:any)=>setSlot(ev.target.value)}>
                      <option value="">{idx===0?"-- Pilih Perawat Onloop 1 --":`-- Onloop ${idx+1} (opsional) --`}</option>
                      {(allNurses as string[]).filter(n=>!taken.includes(n)).map((n,i)=><option key={i} value={n}>{n}</option>)}
                    </select>
                    <div style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",pointerEvents:"none",color:C.g,fontSize:11}}>▼</div>
                  </div>
                  {val&&<button type="button" onClick={()=>setSlot("")} style={{background:"none",border:`1px solid ${C.d}`,borderRadius:8,color:C.d,cursor:"pointer",fontSize:14,padding:"3px 8px",flexShrink:0}}>×</button>}
                </div>
              );
            })}
          </div>
          {(opForm.onloopNurse||"")&&<div style={{fontSize:11,color:C.p,marginTop:5,background:C.pBg,borderRadius:8,padding:"5px 10px"}}>✓ Dipilih: <b>{opForm.onloopNurse}</b></div>}
          <div style={{fontSize:11,color:C.tL,marginTop:4}}>Pilih 1–3 perawat onloop sekaligus. Nomor 2 & 3 opsional.</div>
        </LF>
        {/* RR/Katim */}
        <LF label="5. RR / Recovery Room / Katim">
          <div style={{display:"flex",gap:8}}>
            <div style={{flex:1,position:"relative"}}>
              <select style={{...iStyle("rrKatim"),appearance:"none",paddingRight:30,cursor:"pointer"} as React.CSSProperties} value={opForm.rrKatim||""} onChange={(ev:any)=>setR("rrKatim",ev.target.value)}>
                <option value="">-- Pilih dari staf --</option>
                {[...byT("katim"),...(allNurses as string[])].filter((v,i,a)=>a.indexOf(v)===i).map((n,i)=><option key={i} value={n}>{n}</option>)}
              </select>
              <div style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",pointerEvents:"none",color:C.g,fontSize:11}}>▼</div>
            </div>
            <input style={{...iStyle("rrKatim"),flex:1}} placeholder="Atau ketik manual..." value={opForm.rrKatim||""} onChange={(ev:any)=>set("rrKatim",ev.target.value)} onFocus={onFoc("rrKatim")} onBlur={onBlr} autoComplete="off" spellCheck={false}/>
          </div>
          <div style={{fontSize:11,color:C.tL,marginTop:3}}>Pilih dari dropdown atau ketik nama manual</div>
        </LF>
      </Card>
      <div style={{display:"flex",gap:10}}>
        <Btn full onClick={saveOp} style={{flex:3,padding:"14px",fontSize:15}}>{editingOp?"✓ Simpan Perubahan":"✓ Simpan Jadwal"}</Btn>
        <Btn full outline color={C.g} onClick={()=>{resetOp();setTab("jadwal");}} style={{flex:1,padding:"14px",fontSize:14}}>Batal</Btn>
      </div>
      <div style={{height:24}}/>
    </div>
  );
}

/* ─── VIEW LAPORAN ───────────────────────────────────────────────────── */
function ViewLaporan({ops,lSet,setLSet,lSby,setLSby,staff,roster,showToast}: any) {
  const [preview,setPreview] = useState("");
  const [foc,setFoc]         = useState<string|null>(null);
  const todayOps = ops.filter((o:any)=>o.date===todayDate());
  const tmrwOps  = ops.filter((o:any)=>o.date===tmrwDate());
  const rTmrw    = roster.find((r:any)=>r.date===tmrwDate());
  const tActive  = todayOps.filter((o:any)=>o.status!=="batal");
  const tBatal   = todayOps.filter((o:any)=>o.status==="batal");

  useEffect(()=>{
    if(rTmrw) setLSby((p:any)=>({
      ...p,
      anest:  rTmrw.anestJaga ? [rTmrw.anestJaga] : (p.anest.some((x:any)=>x)?p.anest:[""]),
      cyto:   rTmrw.anestCyto ? [rTmrw.anestCyto] : (p.cyto.some((x:any)=>x)?p.cyto:[""]),
      nurses: rTmrw.nurses?.filter(Boolean).length ? rTmrw.nurses.filter(Boolean) : (p.nurses.some((x:any)=>x)?p.nurses:[""]),
      pembawaHP: rTmrw.pembawaHP||p.pembawaHP||""
    }));
  }, [roster]);

  const gen = () => {
    const t = buildLaporan({
      greeting:lSet.greeting||gWord(), recipients:lSet.recipients||"",
      keterangan:lSet.keterangan, todayOps, tomorrowOps:tmrwOps,
      anestStandby:lSby.anest, anestCyto:lSby.cyto,
      nurseStandby:lSby.nurses, pembawaHP:lSby.pembawaHP||lSet.pembawaHP
    });
    setPreview(t);
  };
  const addI = (k: string) => setLSby((p:any)=>({...p,[k]:[...(p[k]||[]),""] }));
  const setI = (k: string,i: number,v: string) => setLSby((p:any)=>({...p,[k]:(p[k]||[]).map((x:string,j:number)=>j===i?v:x)}));
  const rmI  = (k: string,i: number) => setLSby((p:any)=>({...p,[k]:(p[k]||[]).filter((_:any,j:number)=>j!==i)}));
  const katimList = staff.filter((s:any)=>s.type==="katim");

  return (
    <div>
      <Row title="Laporan Kepala Jaga"/>
      {rTmrw ? <div style={{background:C.sBg,borderRadius:10,padding:"8px 14px",marginBottom:12,fontSize:12,color:C.s,border:`1px solid ${C.s}33`}}>✅ Jadwal siaga besok ditemukan — kolom siaga terisi otomatis. Bisa diedit.</div>
              : <div style={{background:C.wBg,borderRadius:10,padding:"8px 14px",marginBottom:12,fontSize:12,color:C.w}}>⚠ Upload Jadwal Siaga di menu <b>Staf</b>, atau isi manual di bawah.</div>}
      <Card>
        <SH label="① Salam & Penerima"/>
        <LF label="Waktu Salam">
          <div style={{display:"flex",gap:8}}>
            {["pagi","siang","sore","malam"].map(g=>(
              <button key={g} onClick={()=>setLSet((p:any)=>({...p,greeting:g}))} style={{flex:1,padding:"9px 4px",borderRadius:10,border:`1.5px solid ${lSet.greeting===g?C.p:C.b}`,background:lSet.greeting===g?C.p:"#FAFBFD",color:lSet.greeting===g?C.white:C.t,fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
                {g.charAt(0).toUpperCase()+g.slice(1)}
              </button>
            ))}
          </div>
        </LF>
        <LF label="Penerima Laporan">
          <div style={{display:"flex",gap:8,alignItems:"flex-start"}}>
            <textarea style={{...iS,flex:1,resize:"vertical",minHeight:72,lineHeight:1.7,borderColor:foc==="rec"?C.p:C.b,boxShadow:foc==="rec"?`0 0 0 3px ${C.p}22`:""} as React.CSSProperties} placeholder="Nama penerima laporan..." value={lSet.recipients||""} onChange={(ev:any)=>setLSet((p:any)=>({...p,recipients:ev.target.value}))} onFocus={()=>setFoc("rec")} onBlur={()=>setFoc(null)} spellCheck={false}/>
            <div style={{display:"flex",flexDirection:"column",gap:8,flexShrink:0}}>
              <button onClick={()=>setLSet((p:any)=>({...p,recipients:DEFAULT_RECIPIENT}))} style={{background:C.p,border:"none",borderRadius:10,color:"#fff",padding:"9px 12px",cursor:"pointer",fontSize:12,fontWeight:700,fontFamily:"inherit",lineHeight:1.4,whiteSpace:"nowrap"}}>★ Default</button>
              <button onClick={()=>setLSet((p:any)=>({...p,recipients:""}))} style={{background:"none",border:`1.5px solid ${C.b}`,borderRadius:10,color:C.tL,padding:"8px 12px",cursor:"pointer",fontSize:12,fontFamily:"inherit"}}>Hapus</button>
            </div>
          </div>
          {!lSet.recipients && <div style={{fontSize:11,color:C.tL,marginTop:4}}>💡 Klik <b>★ Default</b> untuk isi otomatis teks penerima standar</div>}
        </LF>
      </Card>
      <Card>
        <SH label="② Keterangan Tambahan"/>
        <textarea style={{...iS,resize:"vertical",minHeight:72,lineHeight:1.7,borderColor:foc==="ket"?C.p:C.b,boxShadow:foc==="ket"?`0 0 0 3px ${C.p}22`:""} as React.CSSProperties} placeholder="Catatan khusus shift ini (kosongkan jika tidak ada)..." value={lSet.keterangan||""} onChange={(ev:any)=>setLSet((p:any)=>({...p,keterangan:ev.target.value}))} onFocus={()=>setFoc("ket")} onBlur={()=>setFoc(null)} spellCheck={false}/>
        {lSet.keterangan && <div style={{fontSize:11,color:C.p,marginTop:3}}>✓ Keterangan akan disisipkan di awal laporan</div>}
      </Card>
      <Card style={{background:tActive.length>0?"#F8FBF0":C.gBg,border:`1px solid ${tActive.length>0?"#D4E6A0":C.b}`}}>
        <SH label="③ Acara Operasi Hari Ini" color={C.s}/>
        {tActive.length===0 ? <div style={{fontSize:13,color:C.tL,textAlign:"center",padding:"6px 0"}}>Tidak ada operasi aktif hari ini</div>
          : tActive.sort((a:any,b:any)=>a.time.localeCompare(b.time)).map((op:any,i:number)=>(
            <div key={op.id} style={{display:"flex",gap:10,marginBottom:8,paddingBottom:8,borderBottom:i<tActive.length-1?`1px solid ${C.b}`:"none"}}>
              <div style={{background:op.opType==="cyto"?"#FFCDD2":C.sBg,borderRadius:8,padding:"5px 9px",fontSize:12,fontWeight:700,color:op.opType==="cyto"?"#B71C1C":C.s,flexShrink:0,minWidth:52,textAlign:"center"}}>{op.time}</div>
              <div style={{flex:1}}>
                <div style={{fontSize:13,fontWeight:700,color:C.t}}>{op.patient}{op.age?`, ${op.age}th`:""}</div>
                <div style={{fontSize:12,color:C.tL}}>{op.procedure} · {op.room}</div>
                <div style={{fontSize:11,color:C.tL}}>{op.surgeon}</div>
              </div>
              <Bdg label={STS[op.status as keyof typeof STS]?.l||"Terjadwal"} color={STS[op.status as keyof typeof STS]?.c||C.i} bg={STS[op.status as keyof typeof STS]?.bg||C.iBg}/>
            </div>
          ))
        }
        {tBatal.length>0 && <div style={{fontSize:12,color:C.d,marginTop:6,borderTop:`1px dashed ${C.dL}`,paddingTop:6}}><b>Batal/Tunda:</b> {tBatal.map((o:any)=>o.patient+(o.cancelReason?` (${o.cancelReason})`:""  )).join("; ")}</div>}
      </Card>
      <Card style={{background:"#EDE7F6",border:"1px solid #D1C4E9"}}>
        <SH label="④ Rencana Operasi Besok (H-1)" color="#512DA8"/>
        {tmrwOps.length===0 ? <div style={{fontSize:13,color:C.tL,textAlign:"center",padding:"6px 0"}}>Belum ada rencana operasi besok</div>
          : <>
            <div style={{fontSize:12,color:"#512DA8",marginBottom:10}}>Total <b>{tmrwOps.length}</b> · Pagi (&lt;14.00): <b>{tmrwOps.filter((o:any)=>o.time<"14:00").length}</b> · Sore (≥14.00): <b>{tmrwOps.filter((o:any)=>o.time>="14:00").length}</b></div>
            {tmrwOps.sort((a:any,b:any)=>a.time.localeCompare(b.time)).map((op:any,i:number)=>(
              <div key={op.id} style={{display:"flex",gap:10,marginBottom:8,paddingBottom:8,borderBottom:i<tmrwOps.length-1?"1px solid #D1C4E9":"none"}}>
                <div style={{background:"#D1C4E9",borderRadius:8,padding:"5px 9px",fontSize:12,fontWeight:700,color:"#512DA8",flexShrink:0,minWidth:52,textAlign:"center"}}>{op.time}</div>
                <div style={{flex:1}}>
                  <div style={{fontSize:13,fontWeight:700,color:C.t}}>{op.patient}{op.age?`, ${op.age}th`:""}</div>
                  <div style={{fontSize:12,color:C.tL}}>{op.procedure}</div>
                  <div style={{fontSize:11,color:C.tL}}>{op.surgeon} / {op.anesthesiologist||"—"}</div>
                </div>
                <Bdg label={OT[op.opType as keyof typeof OT||"elektif"]?.label} color={OT[op.opType as keyof typeof OT||"elektif"]?.c||C.i} bg={OT[op.opType as keyof typeof OT||"elektif"]?.bg||C.iBg}/>
              </div>
            ))}
          </>
        }
      </Card>
      {[{k:"anest",l:"⑤ Siaga Dokter Anestesi",c:C.p},{k:"cyto",l:"⑥ Siaga Cyto Dokter Anestesi",c:"#B71C1C"},{k:"nurses",l:"⑦ Siaga Perawat",c:C.p}].map(({k,l,c})=>(
        <Card key={k} hi={k==="cyto"?C.dL:undefined}>
          <SH label={l} color={c}/>
          {rTmrw && <div style={{fontSize:11,color:C.s,marginBottom:8,fontWeight:600}}>✓ Terisi otomatis dari jadwal jaga {fD(tmrwDate())} — bisa diedit</div>}
          {(lSby[k]||[""]).map((n: string,i: number)=>(
            <div key={i} style={{display:"flex",gap:8,marginBottom:8}}>
              <input style={{...iS,flex:1}} placeholder={`${l.replace(/[⑤⑥⑦] /,"")} ${i+1}`} value={n} onChange={(ev:any)=>setI(k,i,ev.target.value)} spellCheck={false}/>
              <button onClick={()=>rmI(k,i)} style={{background:"none",border:`1.5px solid ${C.d}`,borderRadius:8,color:C.d,padding:"0 12px",cursor:"pointer",fontSize:16,flexShrink:0}}>×</button>
            </div>
          ))}
          <Btn sm outline color={c} onClick={()=>addI(k)}>+ Tambah</Btn>
        </Card>
      ))}
      <Card>
        <SH label="⑧ Pembawa HP & Nomor WA Pengiriman"/>
        <LF label="Nama Pembawa HP">
          <input style={iS} placeholder="Nama staf pembawa HP shift ini" value={lSby.pembawaHP||lSet.pembawaHP||""} onChange={(ev:any)=>setLSby((p:any)=>({...p,pembawaHP:ev.target.value}))} spellCheck={false}/>
        </LF>
        <div style={{display:"flex",gap:12}}>
          <div style={{flex:1}}><LF label="No. WA Katim"><input style={iS} placeholder="628112345678" type="tel" inputMode="numeric" value={lSet.katimPhone||""} onChange={(ev:any)=>setLSet((p:any)=>({...p,katimPhone:ev.target.value}))}/></LF></div>
          <div style={{flex:1}}><LF label="No. WA Grup"><input style={iS} placeholder="628112345678" type="tel" inputMode="numeric" value={lSet.grupPhone||""} onChange={(ev:any)=>setLSet((p:any)=>({...p,grupPhone:ev.target.value}))}/></LF></div>
        </div>
        {katimList.length>0 && <div><div style={{fontSize:12,color:C.tL,marginBottom:6}}>Pilih nomor Katim dari staf:</div><div style={{display:"flex",gap:6,flexWrap:"wrap"}}>{katimList.map((s:any)=><button key={s.id} onClick={()=>setLSet((p:any)=>({...p,katimPhone:s.phone}))} style={{fontSize:12,padding:"5px 10px",borderRadius:8,border:`1px solid ${C.p}`,background:C.pBg,color:C.p,cursor:"pointer",fontFamily:"inherit"}}>{s.name.split(",")[0]}</button>)}</div></div>}
      </Card>
      <Btn full onClick={gen} style={{marginBottom:12,padding:"14px",fontSize:15}}>⚙ Generate Laporan</Btn>
      {preview && (
        <Card hi={C.pL+"55"}>
          <div style={{fontSize:13,fontWeight:700,color:C.p,marginBottom:8}}>Pratinjau — Verifikasi sebelum kirim</div>
          <div style={{background:"#F8FAFB",borderRadius:10,padding:"12px 14px",fontSize:13,lineHeight:1.9,whiteSpace:"pre-wrap",border:`1px solid ${C.b}`,marginBottom:12,maxHeight:320,overflowY:"auto",fontFamily:"inherit"}}>{preview}</div>
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
            <Btn outline color={C.p} onClick={()=>navigator.clipboard?.writeText(preview).then(()=>showToast("✓ Teks disalin")).catch(()=>showToast("Salin dari pratinjau di atas"))} style={{flex:1}}>Salin Teks</Btn>
            <WaBt onClick={()=>{ if(!lSet.katimPhone){showToast("Nomor WA Katim belum diisi",C.d);return;} window.open(`https://wa.me/${lSet.katimPhone.replace(/[^0-9]/g,"")}?text=${encodeURIComponent(preview)}`,"_blank"); showToast("✓ WA Katim dibuka","#25D366"); }} style={{flex:1}}>WA Katim</WaBt>
            <WaBt onClick={()=>{ if(!lSet.grupPhone){showToast("Nomor WA Grup belum diisi",C.d);return;} window.open(`https://wa.me/${lSet.grupPhone.replace(/[^0-9]/g,"")}?text=${encodeURIComponent(preview)}`,"_blank"); showToast("✓ WA Grup dibuka","#25D366"); }} style={{flex:1}}>WA Grup</WaBt>
          </div>
        </Card>
      )}
    </div>
  );
}

/* ─── VIEW KIRIM WA ──────────────────────────────────────────────────── */
function ViewKirimWA({ops,staff,setNotifs,showToast}: any) {
  const [prev,setPrev] = useState<Record<string,boolean>>({});
  const getPhone = (n: string) => staff.find((x:any)=>x.name===n)?.phone||null;
  const tmrwOps  = ops.filter((o:any)=>o.date===tmrwDate()&&o.status!=="batal");
  const sendWA   = (ph: string,msg: string,name: string,lbl: string) => {
    if(ph) window.open(`https://wa.me/${ph.replace(/[^0-9]/g,"")}?text=${encodeURIComponent(msg)}`,"_blank");
    setNotifs((p:any)=>[{id:gId(),type:"wa_direct",label:lbl,patient:name,procedure:"Jadwal H-1",message:msg,sentAt:fNow()},...p]);
    showToast("✓ WhatsApp dibuka — tekan Kirim di WA","#25D366");
  };
  const surgeons = [...new Set(tmrwOps.map((o:any)=>o.surgeon))].filter(Boolean);
  const anests   = [...new Set(tmrwOps.map((o:any)=>o.anesthesiologist))].filter(Boolean);
  return (
    <div>
      <div style={{background:`linear-gradient(90deg,${C.pBg},${C.waBg})`,borderRadius:14,padding:"14px 16px",marginBottom:14,border:`1px solid ${C.pL}33`}}>
        <div style={{fontSize:14,fontWeight:700,color:C.p}}>Kamar Bedah {HOSPITAL}</div>
        <div style={{fontSize:12,color:C.tL,marginTop:2}}>Pengiriman H-1 — Besok: {fD(tmrwDate())} · {tmrwOps.length} operasi</div>
      </div>
      <div style={{background:C.iBg,borderRadius:10,padding:"10px 14px",marginBottom:14,fontSize:12,color:C.i,lineHeight:1.6}}>ℹ️ Pesan bersifat DRAFT. Pratinjau wajib ditampilkan. Tekan Kirim di WhatsApp untuk konfirmasi manual.</div>
      {tmrwOps.length===0 && <Card><div style={{textAlign:"center",padding:20,color:C.tL,fontSize:13}}>Belum ada operasi terjadwal untuk besok ({fD(tmrwDate())})</div></Card>}
      {(surgeons as string[]).length>0 && (
        <div style={{marginBottom:16}}>
          <div style={{fontSize:12,fontWeight:700,color:C.g,textTransform:"uppercase",letterSpacing:.8,marginBottom:10}}>H-1 → Dokter Bedah / Operator</div>
          {(surgeons as string[]).map(name=>{
            const sOps=tmrwOps.filter((o:any)=>o.surgeon===name), ph=getPhone(name), msg=msgSurgeon(name,sOps), k="s_"+name;
            return (
              <Card key={name}>
                <div style={{fontSize:14,fontWeight:700,color:C.t,marginBottom:2}}>{name}</div>
                <div style={{fontSize:12,color:C.tL,marginBottom:8}}>{sOps.length} tindakan besok · {ph?<span>+{ph}</span>:<span style={{color:C.d}}>Nomor belum terdaftar</span>}</div>
                <div style={{fontSize:12,color:C.tL,marginBottom:10}}>{sOps.sort((a:any,b:any)=>a.time.localeCompare(b.time)).map((op:any,i:number)=><div key={op.id}>{i+1}. {op.time} — {op.procedure}</div>)}</div>
                <WaBt full onClick={()=>sendWA(ph,msg,name,"H-1 Bedah → "+name)} disabled={!ph}>Kirim H-1 ke Dokter Bedah</WaBt>
                <button onClick={()=>setPrev(p=>({...p,[k]:!p[k]}))} style={{background:"none",border:"none",color:C.p,cursor:"pointer",fontSize:12,fontWeight:600,marginTop:10,padding:0}}>{prev[k]?"▲ Tutup":"▼ Pratinjau Pesan"}</button>
                {prev[k] && <div style={{marginTop:10,background:"#F9FAFB",borderRadius:10,padding:"12px 14px",fontSize:12,lineHeight:1.9,whiteSpace:"pre-wrap",border:`1px solid ${C.b}`,maxHeight:240,overflowY:"auto",fontFamily:"inherit"}}>{msg}</div>}
              </Card>
            );
          })}
        </div>
      )}
      {(anests as string[]).length>0 && (
        <div>
          <div style={{fontSize:12,fontWeight:700,color:C.g,textTransform:"uppercase",letterSpacing:.8,marginBottom:10}}>H-1 → Dokter Anestesi</div>
          {(anests as string[]).map(name=>{
            const aOps=tmrwOps.filter((o:any)=>o.anesthesiologist===name), ph=getPhone(name), msg=msgAnest(name,aOps), k="a_"+name;
            return (
              <Card key={name}>
                <div style={{fontSize:14,fontWeight:700,color:C.t,marginBottom:2}}>{name}</div>
                <div style={{fontSize:12,color:C.tL,marginBottom:8}}>{aOps.length} kasus besok · {ph?<span>+{ph}</span>:<span style={{color:C.d}}>Nomor belum terdaftar</span>}</div>
                <WaBt full onClick={()=>sendWA(ph,msg,name,"H-1 Anestesi → "+name)} disabled={!ph}>Kirim H-1 ke Dokter Anestesi</WaBt>
                <button onClick={()=>setPrev(p=>({...p,[k]:!p[k]}))} style={{background:"none",border:"none",color:C.p,cursor:"pointer",fontSize:12,fontWeight:600,marginTop:10,padding:0}}>{prev[k]?"▲ Tutup":"▼ Pratinjau Pesan"}</button>
                {prev[k] && <div style={{marginTop:10,background:"#F9FAFB",borderRadius:10,padding:"12px 14px",fontSize:12,lineHeight:1.9,whiteSpace:"pre-wrap",border:`1px solid ${C.b}`,maxHeight:240,overflowY:"auto",fontFamily:"inherit"}}>{msg}</div>}
              </Card>
            );
          })}
        </div>
      )}
    </div>
  );
}

/* ─── VIEW STATISTIK ─────────────────────────────────────────────────── */
function ViewStatistik({ops, archive}: {ops: any[]; archive: any[]}) {
  const [range, setRange] = useState<"all"|"30"|"90">("all");
  const [exporting, setExporting] = useState(false);
  const COLORS = ["#00695C","#1565C0","#E65100","#7B1FA2","#2E7D32","#C62828","#00838F","#E91E63"];

  const allOps = [...ops, ...(archive.flatMap((a:any)=>a.ops||[]))];
  const now = new Date();
  const filtered = allOps.filter(op => {
    if(range==="all") return true;
    const d = new Date(op.date+"T00:00:00");
    const days = range==="30"?30:90;
    return (now.getTime()-d.getTime()) <= days*24*60*60*1000;
  });

  const total = filtered.length;
  const selesai = filtered.filter((o:any)=>o.status==="done").length;
  const batal = filtered.filter((o:any)=>o.status==="batal").length;
  const berlangsung = filtered.filter((o:any)=>o.status==="ongoing").length;
  const terjadwal = filtered.filter((o:any)=>o.status==="scheduled").length;
  const cyto = filtered.filter((o:any)=>o.opType==="cyto").length;
  const elektif = filtered.filter((o:any)=>o.opType==="elektif").length;
  const semi = filtered.filter((o:any)=>o.opType==="semi").length;

  // Per surgeon
  const bySurgeon: Record<string,number> = {};
  filtered.forEach((o:any)=>{ if(o.surgeon) bySurgeon[o.surgeon]=(bySurgeon[o.surgeon]||0)+1; });
  const surgeonData = Object.entries(bySurgeon).sort((a,b)=>b[1]-a[1]).slice(0,8).map(([label,value],i)=>({label,value,color:COLORS[i%COLORS.length]}));

  // Per room
  const byRoom: Record<string,number> = {};
  filtered.forEach((o:any)=>{ if(o.room) byRoom[o.room]=(byRoom[o.room]||0)+1; });
  const roomData = Object.entries(byRoom).sort((a,b)=>b[1]-a[1]).map(([label,value],i)=>({label,value,color:COLORS[i%COLORS.length]}));

  // Per month (last 6 months)
  const monthMap: Record<string,number> = {};
  filtered.forEach((o:any)=>{
    if(!o.date) return;
    const m = o.date.slice(0,7);
    monthMap[m]=(monthMap[m]||0)+1;
  });
  const monthData = Object.entries(monthMap).sort(([a],[b])=>a.localeCompare(b)).slice(-6).map(([key,value])=>{
    const [y,m]=key.split("-");
    const names=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];
    return {label:names[parseInt(m)-1]+" "+y.slice(2), value, color:C.p};
  });

  // Per anesthesiologist
  const byAnest: Record<string,number> = {};
  filtered.forEach((o:any)=>{ if(o.anesthesiologist) byAnest[o.anesthesiologist]=(byAnest[o.anesthesiologist]||0)+1; });
  const anestData = Object.entries(byAnest).sort((a,b)=>b[1]-a[1]).slice(0,6).map(([label,value],i)=>({label,value,color:COLORS[i%COLORS.length]}));

  const exportToExcel = () => {
    setExporting(true);
    try {
      const wb = XLSX.utils.book_new();
      const rangeLabel = range==="all"?"Semua Data":range==="30"?"30 Hari Terakhir":"90 Hari Terakhir";

      /* ── Sheet 1: Ringkasan ── */
      const summaryRows = [
        ["LAPORAN STATISTIK KAMAR BEDAH"],
        ["Rumah Sakit Panti Rini"],
        ["Diekspor pada:", fNow()],
        ["Rentang Data:", rangeLabel],
        [""],
        ["RINGKASAN UMUM",""],
        ["Total Operasi", total],
        ["Operasi Selesai", selesai],
        ["Batal / Tunda", batal],
        ["Berlangsung", berlangsung],
        ["Terjadwal", terjadwal],
        [""],
        ["JENIS OPERASI",""],
        ["Elektif", elektif],
        ["Semi-Elektif", semi],
        ["CYTO / Emergency", cyto],
        [""],
        ["TINGKAT KEBERHASILAN",""],
        ["Persentase Selesai", total>0?Math.round((selesai/total)*100)+"%":"0%"],
        ["Persentase Batal",   total>0?Math.round((batal/total)*100)+"%":"0%"],
      ];
      const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
      wsSummary["!cols"] = [{wch:30},{wch:20}];
      XLSX.utils.book_append_sheet(wb, wsSummary, "Ringkasan");

      /* ── Sheet 2: Per Bulan ── */
      if(monthData.length>0) {
        const monthRows = [["Bulan","Jumlah Operasi"],...monthData.map(d=>[d.label,d.value])];
        const wsMonth = XLSX.utils.aoa_to_sheet(monthRows);
        wsMonth["!cols"] = [{wch:20},{wch:20}];
        XLSX.utils.book_append_sheet(wb, wsMonth, "Per Bulan");
      }

      /* ── Sheet 3: Per Dokter Bedah ── */
      if(Object.keys(bySurgeon).length>0) {
        const surgRows = [["Dokter Bedah / Operator","Jumlah Operasi"],...Object.entries(bySurgeon).sort((a,b)=>b[1]-a[1]).map(([n,v])=>[n,v])];
        const wsSurg = XLSX.utils.aoa_to_sheet(surgRows);
        wsSurg["!cols"] = [{wch:35},{wch:20}];
        XLSX.utils.book_append_sheet(wb, wsSurg, "Per Dokter Bedah");
      }

      /* ── Sheet 4: Per Dokter Anestesi ── */
      if(Object.keys(byAnest).length>0) {
        const anRows = [["Dokter Anestesi","Jumlah Operasi"],...Object.entries(byAnest).sort((a,b)=>b[1]-a[1]).map(([n,v])=>[n,v])];
        const wsAn = XLSX.utils.aoa_to_sheet(anRows);
        wsAn["!cols"] = [{wch:35},{wch:20}];
        XLSX.utils.book_append_sheet(wb, wsAn, "Per Dokter Anestesi");
      }

      /* ── Sheet 5: Per Kamar ── */
      if(Object.keys(byRoom).length>0) {
        const roomRows = [["Kamar Operasi","Jumlah Operasi"],...Object.entries(byRoom).sort((a,b)=>b[1]-a[1]).map(([n,v])=>[n,v])];
        const wsRoom = XLSX.utils.aoa_to_sheet(roomRows);
        wsRoom["!cols"] = [{wch:25},{wch:20}];
        XLSX.utils.book_append_sheet(wb, wsRoom, "Per Kamar");
      }

      /* ── Sheet 6: Data Detail semua operasi ── */
      const headers = ["No","Tanggal","Jam","Pasien","Usia","No RM","Jenis","Diagnosa","Tindakan","Kamar","Dokter Bedah","Dokter Anestesi","Asisten","PERAWAT Instrumen","P.Anestesi","Onloop","RR/Katim","Alergi","Gol.Darah","Status"];
      const detailRows = filtered.map((op:any,i:number)=>[
        i+1, op.date||"", op.time||"",
        op.patient||"", op.age||"", op.rm||"",
        OT[op.opType as keyof typeof OT]?.label||"Elektif",
        op.diagnosis||"", op.procedure||"", op.room||"",
        op.surgeon||"", op.anesthesiologist||"",
        op.assistantNurse||"", op.circulatingNurse||"",
        op.anesthesiaNurse||"", op.onloopNurse||"",
        op.rrKatim||"", op.allergy||"", op.bloodType||"",
        STS[op.status as keyof typeof STS]?.l||op.status||""
      ]);
      const wsDetail = XLSX.utils.aoa_to_sheet([headers,...detailRows]);
      wsDetail["!cols"] = [
        {wch:5},{wch:12},{wch:8},{wch:25},{wch:6},{wch:12},{wch:12},
        {wch:30},{wch:35},{wch:18},{wch:28},{wch:28},
        {wch:22},{wch:22},{wch:22},{wch:22},{wch:22},
        {wch:15},{wch:10},{wch:14}
      ];
      XLSX.utils.book_append_sheet(wb, wsDetail, "Data Lengkap");

      const filename = `Statistik_OR_PantiRini_${todayDate()}.xlsx`;
      XLSX.writeFile(wb, filename);
    } finally {
      setExporting(false);
    }
  };

  const StatBox = ({label,value,bg,color}: any) => (
    <div style={{background:bg,borderRadius:12,padding:"14px 16px",textAlign:"center",border:`1px solid ${color}33`}}>
      <div style={{fontSize:28,fontWeight:800,color}}>{value}</div>
      <div style={{fontSize:11,color,fontWeight:600,marginTop:2}}>{label}</div>
    </div>
  );

  return (
    <div>
      <Row title="📊 Statistik Operasi"/>
      <div style={{display:"flex",gap:6,marginBottom:14,background:C.gBg,borderRadius:12,padding:4}}>
        {[{v:"all",l:"Semua"},{v:"30",l:"30 Hari"},{v:"90",l:"90 Hari"}].map(r=>(
          <button key={r.v} onClick={()=>setRange(r.v as any)} style={{flex:1,padding:"8px 4px",borderRadius:10,border:"none",background:range===r.v?C.white:"none",color:range===r.v?C.p:C.g,fontSize:12,fontWeight:range===r.v?700:500,cursor:"pointer",boxShadow:range===r.v?"0 1px 6px rgba(0,0,0,.1)":"none"}}>{r.l}</button>
        ))}
      </div>
      <button
        onClick={exportToExcel}
        disabled={exporting||total===0}
        style={{width:"100%",marginBottom:14,padding:"12px",borderRadius:12,border:"1.5px solid #217346",background:exporting||total===0?"#eee":"#217346",color:exporting||total===0?C.g:"#fff",fontSize:13,fontWeight:700,cursor:exporting||total===0?"not-allowed":"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",justifyContent:"center",gap:8,opacity:total===0?.5:1,transition:"opacity .15s"}}
      >
        {exporting
          ? <><span style={{animation:"spin 1s linear infinite",display:"inline-block"}}>⟳</span> Mengekspor...</>
          : <>📥 Ekspor Excel (.xlsx) — {total} operasi</>
        }
      </button>
      <style>{`@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}`}</style>

      {/* Ringkasan */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:8,marginBottom:12}}>
        <StatBox label="Total Operasi" value={total} bg={C.pBg} color={C.p}/>
        <StatBox label="Selesai" value={selesai} bg={C.sBg} color={C.s}/>
        <StatBox label="Batal/Tunda" value={batal} bg={C.dBg} color={C.d}/>
        <StatBox label="CYTO/Emergency" value={cyto} bg="#FFCDD2" color="#B71C1C"/>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8,marginBottom:14}}>
        <StatBox label="Terjadwal" value={terjadwal} bg={C.iBg} color={C.i}/>
        <StatBox label="Berlangsung" value={berlangsung} bg={C.pBg} color={C.pL}/>
        <StatBox label="Elektif" value={elektif} bg="#E8F5E9" color="#388E3C"/>
      </div>

      {/* Jenis */}
      <Card>
        <SH label="Jenis Operasi"/>
        <MiniBar data={[
          {label:"Elektif",value:elektif,color:"#1565C0"},
          {label:"Semi-Elektif",value:semi,color:"#E65100"},
          {label:"CYTO/Emergency",value:cyto,color:"#B71C1C"},
        ]}/>
      </Card>

      {/* Per bulan */}
      {monthData.length>0 && (
        <Card>
          <SH label="Operasi per Bulan"/>
          <MiniBar data={monthData}/>
        </Card>
      )}

      {/* Per dokter bedah */}
      {surgeonData.length>0 && (
        <Card>
          <SH label="Operasi per Dokter Bedah (Top 8)"/>
          <MiniBar data={surgeonData}/>
        </Card>
      )}

      {/* Per dokter anestesi */}
      {anestData.length>0 && (
        <Card>
          <SH label="Operasi per Dokter Anestesi (Top 6)"/>
          <MiniBar data={anestData}/>
        </Card>
      )}

      {/* Per kamar */}
      {roomData.length>0 && (
        <Card>
          <SH label="Operasi per Kamar"/>
          <MiniBar data={roomData}/>
        </Card>
      )}

      {/* Tingkat keberhasilan */}
      {total>0 && (
        <Card>
          <SH label="Tingkat Keberhasilan"/>
          <div style={{display:"flex",alignItems:"center",gap:16,marginBottom:8}}>
            <div style={{flex:1}}>
              <div style={{fontSize:32,fontWeight:800,color:C.s}}>{Math.round((selesai/total)*100)}%</div>
              <div style={{fontSize:11,color:C.tL,marginTop:2}}>Operasi berhasil dari total {total}</div>
            </div>
            <div style={{flex:2}}>
              <div style={{background:"#F0F4F8",borderRadius:8,height:18,overflow:"hidden"}}>
                <div style={{height:"100%",background:`linear-gradient(90deg,${C.s},${C.pL})`,width:`${Math.round((selesai/total)*100)}%`,borderRadius:8,transition:"width .5s"}}/>
              </div>
              {batal>0 && <div style={{height:8}}/>}
              {batal>0 && <div style={{background:"#F0F4F8",borderRadius:8,height:12,overflow:"hidden"}}>
                <div style={{height:"100%",background:C.d,width:`${Math.round((batal/total)*100)}%`,borderRadius:8}}/>
              </div>}
            </div>
          </div>
          <div style={{display:"flex",gap:10,fontSize:11}}>
            <span style={{color:C.s}}>■ Selesai {selesai}</span>
            {batal>0 && <span style={{color:C.d}}>■ Batal {batal}</span>}
            {terjadwal>0 && <span style={{color:C.i}}>■ Terjadwal {terjadwal}</span>}
          </div>
        </Card>
      )}

      {total===0 && (
        <Card>
          <div style={{textAlign:"center",padding:32,color:C.tL}}>
            <div style={{fontSize:40,marginBottom:8}}>📊</div>
            <div style={{fontSize:14,fontWeight:600}}>Belum ada data operasi</div>
            <div style={{fontSize:12,marginTop:4}}>Daftarkan jadwal operasi untuk melihat statistik</div>
          </div>
        </Card>
      )}
    </div>
  );
}

/* ─── VIEW STAF ──────────────────────────────────────────────────────── */
function ViewStaf({staff,setStaff,roster,setRoster,showToast}: any) {
  const [form,setForm]   = useState({name:"",type:"surgeon",phone:""});
  const [editing,setEd]  = useState<any>(null);
  const [show,setShow]   = useState(false);
  const [err,setErr]     = useState<any>({});
  const [rExp,setRExp]   = useState(false);
  const sfRef = useRef<HTMLInputElement>(null), rsfRef = useRef<HTMLInputElement>(null);

  const validate = () => {
    const e: any={};
    if(!form.name.trim()) e.name="Wajib";
    if(!form.phone.trim()) e.phone="Wajib";
    else if(!/^62\d{8,13}$/.test(form.phone.replace(/\s/g,""))) e.phone="Format: 62xxx (tanpa + atau spasi)";
    setErr(e); return !Object.keys(e).length;
  };
  const save = () => {
    if(!validate()) return;
    if(editing){ setStaff((p:any)=>p.map((s:any)=>s.id===editing.id?{...editing,...form}:s)); showToast("✓ Data staf diperbarui",C.s); }
    else { setStaff((p:any)=>[...p,{id:gId(),...form}]); showToast("✓ Staf berhasil ditambahkan",C.s); }
    setForm({name:"",type:"surgeon",phone:""}); setEd(null); setShow(false); setErr({});
  };
  const dlStaff = () => {
    const csv = `"nama","jabatan","nomor_wa"\n"dr. Nama Bedah Sp.B","surgeon","628112345001"\n"dr. Nama Anestesi Sp.An","anesthesiologist","628112345002"\n"Ns. Nama Instrumen","circulating","628112345003"\n"Ns. Nama P.Anestesi","anesthesia_nurse","628112345004"\n"Ns. Nama Onloop","onloop","628112345005"\n"Ns. Nama Katim","katim","628112345006"`;
    downloadBlob(csv,"template_staf.csv"); showToast("✓ Template CSV Staf diunduh",C.s);
  };
  const dlRoster = () => {
    const csv = "\"Tanggal\",\"Dokter Anestesi Jaga\",\"Dokter Anestesi Cyto\",\"Perawat Siaga 1\",\"Perawat Siaga 2\",\"Perawat Siaga 3\",\"Pembawa HP\"\n\"01/06/2025\",\"dr. Anestesi Jaga Sp.An\",\"dr. Anestesi Cyto Sp.An\",\"Ns. Perawat 1\",\"Ns. Perawat 2\",\"Ns. Perawat 3\",\"Ns. Katim Jaga\"";
    downloadBlob(csv,"template_jadwal_siaga_bulanan.csv"); showToast("✓ Template Jadwal Siaga diunduh",C.s);
  };
  const handleStaffFile = async (ev: any) => {
    const f=ev.target.files[0]; if(!f)return; ev.target.value="";
    try {
      const rows=parseCSV(await f.text());
      if(rows.length<2){showToast("File kosong atau format tidak sesuai",C.d);return;}
      const h=rows[0].map((x:string)=>x.toLowerCase());
      const ni=h.findIndex((x:string)=>x.includes("nama")), ti=h.findIndex((x:string)=>x.includes("jabatan")||x.includes("type")), pi=h.findIndex((x:string)=>x.includes("wa")||x.includes("phone")||x.includes("nomor"));
      if(ni<0){showToast("Kolom 'nama' tidak ditemukan. Gunakan template.",C.d);return;}
      const imp=rows.slice(1).map((c:string[])=>({id:gId(),name:sanitize(c[ni]||""),type:ti>=0?(c[ti]||"circulating"):"circulating",phone:(pi>=0?c[pi]||"":"").replace(/[^0-9]/g,"")})).filter((s:any)=>s.name);
      if(!imp.length){showToast("Tidak ada data valid",C.d);return;}
      setStaff((p:any)=>[...p,...imp]); showToast("✓ "+imp.length+" staf berhasil diimport",C.s);
    } catch(e: any){showToast("Error: "+e.message,C.d);}
  };
  const handleRosterFile = async (ev: any) => {
    const f=ev.target.files[0]; if(!f)return; ev.target.value="";
    try {
      const rows=parseCSV(await f.text());
      if(rows.length<2){showToast("File kosong",C.d);return;}
      const h=rows[0].map((x:string)=>x.toLowerCase());
      const di=h.findIndex((x:string)=>x.includes("tanggal")||x.includes("date")), ai=h.findIndex((x:string)=>x.includes("jaga")&&!x.includes("cyto")), ci=h.findIndex((x:string)=>x.includes("cyto")), pi=h.findIndex((x:string)=>x.includes("perawat")||x.includes("nurse")), hpi=h.findIndex((x:string)=>x.includes("pembawa")||x.includes("hp"));
      if(di<0){showToast("Kolom 'Tanggal' tidak ditemukan. Gunakan template.",C.d);return;}
      const imp=rows.slice(1).map((r:string[])=>{
        const date=parseDateCSV(r[di]||"")||(r[di]||"");
        const nurses: string[]=[];
        if(pi>=0) for(let i=pi;i<Math.min(r.length,pi+6);i++) if(r[i]&&r[i].trim()) nurses.push(sanitize(r[i].trim()));
        return {id:gId(),date,anestJaga:sanitize(r[ai>=0?ai:0]||""),anestCyto:sanitize(r[ci>=0?ci:0]||""),nurses,pembawaHP:sanitize(r[hpi>=0?hpi:0]||"")};
      }).filter((r:any)=>r.date&&r.date.length>=8);
      if(!imp.length){showToast("Tidak ada baris valid. Cek format tanggal DD/MM/YYYY",C.d);return;}
      setRoster((p:any)=>[...p.filter((x:any)=>!imp.find((n:any)=>n.date===x.date)),...imp]);
      showToast("✓ "+imp.length+" hari jadwal siaga berhasil diimport",C.s);
    } catch(e: any){showToast("Error: "+e.message,C.d);}
  };

  return (
    <div>
      <Row title="Manajemen Staf & Roster" right={<Btn sm onClick={()=>{setForm({name:"",type:"surgeon",phone:""});setEd(null);setErr({});setShow(true);}}>+ Tambah Staf</Btn>}/>
      <Card style={{background:C.iBg,border:"1px solid #1565C033"}}>
        <SH label="📋 Upload Data Master Staf" color={C.i}/>
        <div style={{fontSize:12,color:C.tL,marginBottom:10,lineHeight:1.6}}>Format: <b>CSV</b> · Kolom: nama, jabatan, nomor_wa</div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <input ref={sfRef} type="file" accept=".csv,.txt" onChange={handleStaffFile} style={{display:"none"}}/>
          <Btn sm outline color={C.i} onClick={()=>sfRef.current&&sfRef.current.click()}>📂 Upload CSV Staf</Btn>
          <Btn sm outline color={C.g} onClick={dlStaff}>⬇ Template</Btn>
        </div>
        {staff.length>0 && <div style={{marginTop:8,fontSize:12,color:C.i,fontWeight:600}}>✓ {staff.length} staf terdaftar</div>}
      </Card>
      <Card style={{background:"#F3E5F5",border:"2px solid #CE93D8"}}>
        <SH label="📅 Upload Jadwal Siaga 1 Bulan" color="#7B1FA2"/>
        <div style={{background:"#EDE7F6",borderRadius:8,padding:"10px 12px",marginBottom:12,fontSize:12,color:"#512DA8",lineHeight:1.7}}>
          <b>Fungsi:</b> Basis data siaga 1 bulan untuk auto-populate laporan kepala jaga.<br/>
          Kolom wajib: Tanggal (DD/MM/YYYY), Dokter Anestesi Jaga, Dokter Anestesi Cyto, Perawat Siaga 1–3, Pembawa HP
        </div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <input ref={rsfRef} type="file" accept=".csv,.txt" onChange={handleRosterFile} style={{display:"none"}}/>
          <Btn sm color="#7B1FA2" onClick={()=>rsfRef.current&&rsfRef.current.click()}>📂 Upload Jadwal Siaga (CSV)</Btn>
          <Btn sm outline color={C.g} onClick={dlRoster}>⬇ Template</Btn>
        </div>
        {roster.length>0 && (
          <div style={{marginTop:10,background:"#EDE7F6",borderRadius:8,padding:"10px 12px"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div style={{fontSize:12,color:"#512DA8",fontWeight:700}}>✓ {roster.length} hari jadwal siaga tersimpan</div>
              <button onClick={()=>setRExp((v:boolean)=>!v)} style={{background:"none",border:"none",color:"#7B1FA2",cursor:"pointer",fontSize:12,fontWeight:600}}>{rExp?"▲ Tutup":"▼ Lihat"}</button>
            </div>
            <div style={{fontSize:11,color:"#7B1FA2",marginTop:4}}>Hari ini: {roster.find((r:any)=>r.date===todayDate())?"✓":"—"} · Besok: {roster.find((r:any)=>r.date===tmrwDate())?"✓":"—"}</div>
            {rExp && roster.slice(0,5).map((r:any)=>(
              <div key={r.id} style={{fontSize:11,color:"#7B1FA2",marginTop:6,paddingTop:6,borderTop:"1px solid #D1C4E9"}}>
                📅 {fD(r.date)} · Jaga: {r.anestJaga||"—"} · Cyto: {r.anestCyto||"—"}<br/>
                Perawat: {(r.nurses||[]).join(", ")||"—"} · HP: {r.pembawaHP||"—"}
              </div>
            ))}
          </div>
        )}
      </Card>
      {show && (
        <Card hi="#00897B66">
          <div style={{fontSize:14,fontWeight:700,color:C.p,marginBottom:14}}>{editing?"Edit Staf":"Tambah Staf Baru"}</div>
          <LF label="Nama Lengkap & Gelar" req err={err.name}><input style={iS} placeholder="dr. Nama, Sp.B  atau  Ns. Nama, S.Kep" value={form.name} onChange={(e:any)=>setForm((p:any)=>({...p,name:e.target.value}))} autoComplete="off" spellCheck={false}/></LF>
          <SF label="Jabatan" options={Object.entries(ST).map(([v,l])=>({v,l}))} value={form.type} onChange={(e:any)=>setForm((p:any)=>({...p,type:e.target.value}))}/>
          <LF label="Nomor WhatsApp" req err={err.phone}><input style={iS} placeholder="628112345678" type="tel" inputMode="numeric" value={form.phone} onChange={(e:any)=>setForm((p:any)=>({...p,phone:e.target.value.replace(/[^0-9]/g,"")}))} autoComplete="off"/></LF>
          <div style={{fontSize:11,color:C.tL,marginBottom:12}}>Diawali 62, tanpa + atau spasi</div>
          <div style={{display:"flex",gap:8}}><Btn full onClick={save}>{editing?"Simpan Perubahan":"Tambah Staf"}</Btn><Btn full outline color={C.g} onClick={()=>{setShow(false);setEd(null);setErr({});}}>Batal</Btn></div>
        </Card>
      )}
      {staff.length===0 && !show && <Card><div style={{textAlign:"center",padding:24,color:C.tL}}><div style={{fontSize:40,marginBottom:8}}>👥</div><div style={{fontSize:14,fontWeight:600}}>Belum ada data staf</div><div style={{fontSize:12,marginTop:4}}>Tambah manual atau upload CSV</div></div></Card>}
      {Object.entries(ST).map(([type,label])=>{
        const list=staff.filter((s:any)=>s.type===type); if(!list.length)return null;
        return (
          <div key={type} style={{marginBottom:14}}>
            <div style={{fontSize:11,fontWeight:700,color:C.g,textTransform:"uppercase",letterSpacing:.8,marginBottom:6}}>{label}</div>
            {list.map((s:any)=>(
              <Card key={s.id} style={{padding:"12px 16px",marginBottom:6}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <div><div style={{fontSize:14,fontWeight:700,color:C.t}}>{s.name}</div><div style={{fontSize:12,color:C.tL}}>+{s.phone}</div></div>
                  <div style={{display:"flex",gap:6}}>
                    <Btn sm outline color={C.p} onClick={()=>{setForm({name:s.name,type:s.type,phone:s.phone});setEd(s);setErr({});setShow(true);}}>Edit</Btn>
                    <Btn sm outline color={C.d} onClick={()=>{setStaff((p:any)=>p.filter((x:any)=>x.id!==s.id));showToast("Staf dihapus");}}>Hapus</Btn>
                  </div>
                </div>
              </Card>
            ))}
          </div>
        );
      })}
    </div>
  );
}

/* ─── IMPORT EXCEL ───────────────────────────────────────────────────── */
function ImportExcel({ops, setOps, showToast}: any) {
  const [rows, setRows] = useState<any[]>([]);
  const [fileName, setFileName] = useState("");
  const [importing, setImporting] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);

  const COLS: {label:string;key:string;required?:boolean}[] = [
    {label:"Tanggal (YYYY-MM-DD)",key:"date",required:true},
    {label:"Jam (HH:MM)",key:"time",required:true},
    {label:"Nama Pasien",key:"patient",required:true},
    {label:"Tindakan/Prosedur",key:"procedure",required:true},
    {label:"Diagnosa",key:"diagnosis"},
    {label:"Dokter Bedah",key:"surgeon"},
    {label:"Dokter Anestesi",key:"anesthesiologist"},
    {label:"Kamar Operasi",key:"room"},
    {label:"Usia",key:"age"},
    {label:"No RM",key:"rm"},
    {label:"Golongan Darah",key:"bloodType"},
    {label:"Alergi",key:"allergy"},
    {label:"Jenis Operasi",key:"opType"},
    {label:"PERAWAT Instrumen",key:"circulatingNurse"},
    {label:"P. Anestesi",key:"anesthesiaNurse"},
    {label:"Onloop",key:"onloopNurse"},
    {label:"RR/Katim",key:"rrKatim"},
  ];

  const handleFile = (file: File) => {
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target!.result as ArrayBuffer), {type:"array"});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data: any[][] = XLSX.utils.sheet_to_json(ws, {header:1});
        if(data.length<2){showToast("File kosong atau tidak ada data",C.d);return;}
        const header: string[] = (data[0] as string[]).map(h=>String(h||"").trim());
        const parsed: any[] = [];
        for(let i=1;i<data.length;i++){
          const row=data[i] as any[];
          if(!row||row.every((c:any)=>!c)) continue;
          const obj: any={};
          COLS.forEach(col=>{
            const idx=header.findIndex(h=>h.toLowerCase()===col.label.toLowerCase()||h.toLowerCase()===col.key.toLowerCase());
            if(idx>=0) obj[col.key]=String(row[idx]||"").trim();
            else obj[col.key]="";
          });
          if(obj.date||obj.patient||obj.procedure) parsed.push(obj);
        }
        setRows(parsed);
        if(parsed.length===0) showToast("Tidak ada baris data ditemukan",C.w);
      } catch(err){showToast("Gagal membaca file Excel",C.d);}
    };
    reader.readAsArrayBuffer(file);
  };

  const doImport = () => {
    setImporting(true);
    const valid = rows.filter(r=>r.patient&&r.date&&r.procedure);
    const newOps = valid.map(r=>({
      id:gId(), status:"scheduled", reminders:[], requests:[], createdAt:fNow(),
      opType: r.opType||"elektif",
      date: r.date, time: r.time||"08:00",
      patient: r.patient, diagnosis: r.diagnosis||"", procedure: r.procedure,
      surgeon: r.surgeon||"", anesthesiologist: r.anesthesiologist||"",
      room: r.room||ROOMS[0], age: r.age||"", rm: r.rm||"",
      bloodType: r.bloodType||"Tidak Diketahui", allergy: r.allergy||"Tidak Ada",
      circulatingNurse: r.circulatingNurse||"", anesthesiaNurse: r.anesthesiaNurse||"",
      onloopNurse: r.onloopNurse||"", rrKatim: r.rrKatim||"",
    }));
    setOps((p:any[])=>[...p,...newOps]);
    showToast(`✓ ${newOps.length} jadwal berhasil diimport`,C.s);
    setRows([]); setFileName(""); setImporting(false);
  };

  const downloadTemplate = () => {
    const header = COLS.map(c=>c.label);
    const example = [todayDate(),"08:00","Budi Santoso","Appendectomy","Appendicitis Akut","dr. Ahmad","dr. Budi","Kamar Operasi 1","45","RM-001","O+","Tidak Ada","elektif","","","",""];
    const ws = XLSX.utils.aoa_to_sheet([header,example]);
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Template");
    XLSX.writeFile(wb,"Template_Import_Jadwal.xlsx");
    showToast("✓ Template Excel diunduh",C.s);
  };

  return (
    <div>
      <Card>
        <SH label="📥 Import Jadwal dari Excel"/>
        <div style={{background:"#E3F2FD",borderRadius:8,padding:"10px 12px",marginBottom:14,fontSize:12,color:"#1565C0",lineHeight:1.7,border:"1px solid #90CAF9"}}>
          <b>Cara import:</b><br/>
          1. Unduh template Excel di bawah ini<br/>
          2. Isi data jadwal sesuai format kolom<br/>
          3. Upload file, cek preview, lalu klik Import
        </div>
        <Btn full outline color={C.i} onClick={downloadTemplate} style={{marginBottom:12}}>⬇ Unduh Template Excel</Btn>
        <div
          style={{border:`2px dashed ${C.b}`,borderRadius:12,padding:"24px 16px",textAlign:"center",cursor:"pointer",marginBottom:12,background:"#FAFBFD"}}
          onClick={()=>fileRef.current?.click()}
          onDragOver={e=>{e.preventDefault();}}
          onDrop={e=>{e.preventDefault();const f=e.dataTransfer.files[0];if(f)handleFile(f);}}>
          <div style={{fontSize:28,marginBottom:6}}>📊</div>
          <div style={{fontSize:13,fontWeight:600,color:C.t}}>{fileName||"Klik atau seret file .xlsx / .xls ke sini"}</div>
          <div style={{fontSize:11,color:C.tL,marginTop:4}}>Format: Excel (.xlsx, .xls)</div>
        </div>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={e=>{const f=e.target.files?.[0];if(f)handleFile(f);e.target.value="";}}/>
        {rows.length>0 && (
          <div>
            <div style={{fontSize:13,fontWeight:700,color:C.t,marginBottom:8}}>{rows.length} baris ditemukan — Preview:</div>
            <div style={{overflowX:"auto",marginBottom:12}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                <thead><tr style={{background:C.pBg}}>{["Tanggal","Jam","Pasien","Tindakan","Kamar","Dokter"].map(h=><th key={h} style={{padding:"6px 10px",textAlign:"left",color:C.p,fontWeight:700,borderBottom:`2px solid ${C.p}33`}}>{h}</th>)}</tr></thead>
                <tbody>{rows.slice(0,10).map((r,i)=><tr key={i} style={{borderBottom:`1px solid ${C.b}`}}><td style={{padding:"5px 10px"}}>{r.date}</td><td style={{padding:"5px 10px"}}>{r.time}</td><td style={{padding:"5px 10px",fontWeight:600}}>{r.patient}</td><td style={{padding:"5px 10px"}}>{r.procedure}</td><td style={{padding:"5px 10px"}}>{r.room}</td><td style={{padding:"5px 10px"}}>{r.surgeon}</td></tr>)}</tbody>
              </table>
              {rows.length>10&&<div style={{fontSize:11,color:C.tL,marginTop:6,textAlign:"center"}}>...dan {rows.length-10} baris lainnya</div>}
            </div>
            <div style={{display:"flex",gap:10}}>
              <Btn full onClick={doImport} disabled={importing} style={{flex:2}}>{importing?"⟳ Mengimport...":"✚ Import "+rows.length+" Jadwal"}</Btn>
              <Btn full outline color={C.g} onClick={()=>{setRows([]);setFileName("");}} style={{flex:1}}>Batal</Btn>
            </div>
          </div>
        )}
      </Card>
    </div>
  );
}

/* ─── VIEW ARSIP ─────────────────────────────────────────────────────── */
function ViewArsip({ops,setOps,notifs,archive,showToast,privacyMode,setPrivacyMode,supaCfg,setSupaCfg,onSupaBackup,onSupaRestoreOps,onSupaRestoreLembur,onSupaRestoreMonitoring,onSupaRestoreAll,supaStatus,supaBackingUp,auditLog,rtStatus,rtEnabled,setRtEnabled,dbxCfg,setDbxCfg,onDbxBackup,onDbxRestoreOps,onDbxRestoreLembur,dbxStatus,dbxBacking,onDbxBackupOpsXls,onDbxBackupLemburXls,onDbxBackupMonitoringXls,lemburData,lemburPegawai,monitoringEntries,monitoringCfg}: any) {
  const [sub,setSub] = useState("notif");
  const [arsipTahun,setArsipTahun]  = useState("");
  const [arsipBulan,setArsipBulan]  = useState("");
  const [arsipTipe,setArsipTipe]    = useState("laporan_operasi");
  const [dlCloudMonth,setDlCloudMonth] = useState(()=>{const d=new Date();return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0");});
  const [dlCloudType,setDlCloudType] = useState<"ops"|"lembur"|"monitoring">("ops");
  const [dlCloudBusy,setDlCloudBusy] = useState(false);
  const nCfg: any = {cyto:{c:C.d,bg:C.dBg},"H-1":{c:C.i,bg:C.iBg},"H-1jam":{c:C.w,bg:C.wBg},wa_direct:{c:C.wa,bg:C.waBg}};

  /* Compute available years from ops + lembur */
  const allYears = [...new Set([
    ...ops.map((o:any)=>o.date?.slice(0,4)).filter(Boolean),
    ...Object.keys(lemburData||{}).map((k:string)=>k.split("_").pop()?.slice(0,4)).filter(Boolean),
  ])].sort().reverse() as string[];

  const MONTHS = ["01","02","03","04","05","06","07","08","09","10","11","12"];
  const MONTH_LABELS: Record<string,string> = {
    "01":"Januari","02":"Februari","03":"Maret","04":"April","05":"Mei","06":"Juni",
    "07":"Juli","08":"Agustus","09":"September","10":"Oktober","11":"November","12":"Desember",
  };

  const filteredOpsArsip = ops.filter((o:any)=>{
    if(arsipTahun && !o.date?.startsWith(arsipTahun)) return false;
    if(arsipBulan && o.date?.slice(5,7) !== arsipBulan) return false;
    return true;
  });

  const downloadArsipExcel = () => {
    if(arsipTipe==="laporan_operasi"||arsipTipe==="jadwal_operasi") {
      if(!filteredOpsArsip.length){showToast("Tidak ada data operasi untuk periode ini",C.w);return;}
      const h = ["No","Tanggal","Jam","Pasien","Usia","RM","Jenis","Diagnosa","Tindakan","Kamar","Operator","Anestesi","Asisten","Instrumen","P.Anestesi","Onloop","RR/Katim","Status"];
      const rows = filteredOpsArsip.map((o:any,i:number)=>[String(i+1),o.date,o.time,o.patient||"",o.age||"",o.rm||"",o.opType||"elektif",o.diagnosis||"",o.procedure||"",o.room||"",o.surgeon||"",o.anesthesiologist||"",o.assistantNurse||"",o.circulatingNurse||"",o.anesthesiaNurse||"",o.onloopNurse||"",o.rrKatim||"",o.status||""]);
      const ws = XLSX.utils.aoa_to_sheet([h,...rows]);
      ws["!cols"]=[{wch:4},{wch:12},{wch:7},{wch:22},{wch:5},{wch:12},{wch:10},{wch:28},{wch:28},{wch:14},{wch:22},{wch:22},{wch:22},{wch:22},{wch:18},{wch:18},{wch:16},{wch:12}];
      const wb = XLSX.utils.book_new();
      const label = (arsipTipe==="laporan_operasi"?"Laporan":"Jadwal")+" Operasi"+(arsipTahun?" "+arsipTahun:"")+(arsipBulan?" "+MONTH_LABELS[arsipBulan]:"");
      XLSX.utils.book_append_sheet(wb, ws, label.slice(0,31));
      XLSX.writeFile(wb, `${arsipTipe==="laporan_operasi"?"Laporan":"Jadwal"}_Operasi${arsipTahun?"_"+arsipTahun:""}${arsipBulan?"_"+arsipBulan:""}_${todayDate()}.xlsx`);
      showToast("✓ Excel arsip diunduh",C.s);
    } else if(arsipTipe==="monitoring_suhu") {
      const me = (monitoringEntries||[]).filter((e:any)=>{
        if(arsipTahun && !e.tanggal?.startsWith(arsipTahun)) return false;
        if(arsipBulan && e.tanggal?.slice(5,7)!==arsipBulan) return false;
        return true;
      });
      if(!me.length){showToast("Tidak ada data monitoring untuk periode ini",C.w);return;}
      const h = ["Tanggal","Jam","Suhu (°C)","Kelembaban (%)","Status","Petugas","Lokasi"];
      const rows = me.map((e:any)=>[e.tanggal,e.jam,e.suhu,e.kelembaban,monIsOK(e.suhu,e.kelembaban,monitoringCfg)?"SESUAI":"TIDAK SESUAI",e.petugas,monitoringCfg?.lokasiRuang||""]);
      const ws = XLSX.utils.aoa_to_sheet([h,...rows]);
      ws["!cols"]=[{wch:12},{wch:8},{wch:12},{wch:14},{wch:16},{wch:22},{wch:20}];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Monitoring Suhu");
      XLSX.writeFile(wb, `Monitoring_Suhu${arsipTahun?"_"+arsipTahun:""}${arsipBulan?"_"+arsipBulan:""}_${todayDate()}.xlsx`);
      showToast("✓ Excel monitoring suhu diunduh",C.s);
    } else {
      /* Data Lembur */
      const rows: string[][] = [];
      (lemburPegawai||[]).forEach((p:any)=>{
        Object.keys(lemburData||{}).filter((k:string)=>{
          if(!k.startsWith(p.id+"_")) return false;
          const ym = k.replace(p.id+"_","");
          if(arsipTahun && !ym.startsWith(arsipTahun)) return false;
          if(arsipBulan && ym.slice(5,7)!==arsipBulan) return false;
          return true;
        }).forEach(k=>{
          const rec = lemburData[k];
          (rec.entries||[]).forEach((e:any)=>rows.push([p.name,p.nik||"",k.replace(p.id+"_",""),e.tanggalAwal||"",e.tanggalAkhir||"",e.jamMasuk||"",e.jamKeluar||"",e.keperluanLembur||"",e.keterangan||""]));
        });
      });
      if(!rows.length){showToast("Tidak ada data lembur untuk periode ini",C.w);return;}
      const ws = XLSX.utils.aoa_to_sheet([["Nama","NIK","Bulan","Tgl Awal Lembur","Tgl Akhir Lembur","Jam Masuk","Jam Keluar","Keperluan","Keterangan"],...rows]);
      ws["!cols"]=[{wch:22},{wch:16},{wch:12},{wch:14},{wch:14},{wch:10},{wch:10},{wch:28},{wch:22}];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Data Lembur");
      XLSX.writeFile(wb, `Data_Lembur${arsipTahun?"_"+arsipTahun:""}${arsipBulan?"_"+arsipBulan:""}_${todayDate()}.xlsx`);
      showToast("✓ Excel lembur arsip diunduh",C.s);
    }
  };

  const downloadArsipWord = () => {
    if(arsipTipe==="laporan_operasi"||arsipTipe==="jadwal_operasi") {
      if(!filteredOpsArsip.length){showToast("Tidak ada data operasi untuk periode ini",C.w);return;}
      const h = ["No","Tanggal","Jam","Pasien","Diagnosa","Tindakan","Kamar","Operator","Anestesi","Status"];
      const rows: string[][] = filteredOpsArsip.map((o:any,i:number)=>[String(i+1),o.date||"",o.time||"",o.patient||"",o.diagnosis||"",o.procedure||"",o.room||"",o.surgeon||"",o.anesthesiologist||"",o.status||""]);
      const label = (arsipTipe==="laporan_operasi"?"Laporan":"Jadwal")+" Operasi"+(arsipTahun?" "+arsipTahun:"")+(arsipBulan?" "+MONTH_LABELS[arsipBulan]:"");
      downloadAsWord(label+" — "+HOSPITAL, [h,...rows], `${arsipTipe==="laporan_operasi"?"Laporan":"Jadwal"}_Operasi${arsipTahun?"_"+arsipTahun:""}${arsipBulan?"_"+arsipBulan:""}_${todayDate()}.doc`);
      showToast("✓ Word arsip diunduh",C.s);
    } else {
      const allRows: string[][] = [];
      (lemburPegawai||[]).forEach((p:any)=>{
        Object.keys(lemburData||{}).filter((k:string)=>{
          if(!k.startsWith(p.id+"_")) return false;
          const ym=k.replace(p.id+"_","");
          if(arsipTahun && !ym.startsWith(arsipTahun)) return false;
          if(arsipBulan && ym.slice(5,7)!==arsipBulan) return false;
          return true;
        }).forEach(k=>{
          const rec=lemburData[k];
          (rec.entries||[]).forEach((e:any)=>allRows.push([p.name,p.nik||"",k.replace(p.id+"_",""),e.tanggalAwal||"",e.jamMasuk||"",e.jamKeluar||"",e.keperluanLembur||"",e.keterangan||""]));
        });
      });
      if(!allRows.length){showToast("Tidak ada data lembur untuk periode ini",C.w);return;}
      const h=["Nama","NIK","Bulan","Tgl Lembur","Jam Masuk","Jam Keluar","Keperluan","Keterangan"];
      const label="Data Lembur"+(arsipTahun?" "+arsipTahun:"")+(arsipBulan?" "+MONTH_LABELS[arsipBulan]:"");
      downloadAsWord(label+" — "+HOSPITAL,[h,...allRows],`Data_Lembur${arsipTahun?"_"+arsipTahun:""}${arsipBulan?"_"+arsipBulan:""}_${todayDate()}.doc`);
      showToast("✓ Word lembur arsip diunduh",C.s);
    }
  };

  /* Download tracking file: Operasi */
  const downloadTrackingOps = (fmt:"excel"|"word") => {
    if(!ops.length){showToast("Belum ada data operasi",C.w);return;}
    const h=["No","Tanggal","Jam","Pasien","Usia","RM","Diagnosa","Tindakan","Jenis","Kamar","Operator","Anestesi","Asisten","Instrumen","P.Anestesi","Onloop","RR/Katim","Alergi","Gol.Darah","Status","Dibuat"];
    const rows: string[][] = [...ops].sort((a:any,b:any)=>a.date.localeCompare(b.date)||a.time.localeCompare(b.time)).map((o:any,i:number)=>[String(i+1),o.date||"",o.time||"",o.patient||"",o.age||"",o.rm||"",o.diagnosis||"",o.procedure||"",o.opType||"",o.room||"",o.surgeon||"",o.anesthesiologist||"",o.assistantNurse||"",o.circulatingNurse||"",o.anesthesiaNurse||"",o.onloopNurse||"",o.rrKatim||"",o.allergy||"",o.bloodType||"",o.status||"",o.createdAt||""]);
    const fn = `Tracking_Operasi_${todayDate()}`;
    if(fmt==="excel"){
      const ws=XLSX.utils.aoa_to_sheet([h,...rows]);
      const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Tracking Operasi");
      XLSX.writeFile(wb,fn+".xlsx"); showToast("✓ Tracking Operasi Excel diunduh",C.s);
    } else {
      downloadAsWord("Tracking Jadwal Operasi — "+HOSPITAL,[h,...rows],fn+".doc");
      showToast("✓ Tracking Operasi Word diunduh",C.s);
    }
  };

  /* Download tracking file: Lembur */
  const downloadTrackingLembur = (fmt:"excel"|"word") => {
    const rows: string[][] = [];
    (lemburPegawai||[]).forEach((p:any)=>{
      Object.keys(lemburData||{}).filter((k:string)=>k.startsWith(p.id+"_")).forEach(k=>{
        const rec=lemburData[k];
        (rec.entries||[]).forEach((e:any)=>rows.push([p.name,p.nik||"",k.replace(p.id+"_",""),e.tanggalAwal||"",e.tanggalAkhir||"",e.jamMasuk||"",e.jamKeluar||"",e.keperluanLembur||"",e.keterangan||"",e.ttd||""]));
      });
    });
    if(!rows.length){showToast("Belum ada data lembur",C.w);return;}
    const h=["Nama Pegawai","NIK","Bulan","Tgl Awal","Tgl Akhir","Jam Masuk","Jam Keluar","Keperluan Lembur","Keterangan","TTD"];
    const fn=`Tracking_Lembur_${todayDate()}`;
    if(fmt==="excel"){
      const ws=XLSX.utils.aoa_to_sheet([h,...rows]);
      const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Tracking Lembur");
      XLSX.writeFile(wb,fn+".xlsx"); showToast("✓ Tracking Lembur Excel diunduh",C.s);
    } else {
      downloadAsWord("Tracking Data Lembur — "+HOSPITAL,[h,...rows],fn+".doc");
      showToast("✓ Tracking Lembur Word diunduh",C.s);
    }
  };

  /* helper: bulan label dari "YYYY-MM" */
  const ymLabel = (ym:string) => {
    if(!ym) return "";
    const ML: Record<string,string> = {"01":"Jan","02":"Feb","03":"Mar","04":"Apr","05":"Mei","06":"Jun","07":"Jul","08":"Agu","09":"Sep","10":"Okt","11":"Nov","12":"Des"};
    const [y,m] = ym.split("-");
    return `${ML[m]||m} ${y}`;
  };

  /* helper: build Excel & download */
  const buildAndDownloadXlsx = (title:string, header:string[], rows:(string|number)[][], filename:string) => {
    const ws = XLSX.utils.aoa_to_sheet([header,...rows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, title.slice(0,31));
    XLSX.writeFile(wb, filename);
  };

  /* ── Download lokal berdasarkan bulan ── */
  const dlFromLocal = (month:string, type:"ops"|"lembur"|"monitoring") => {
    if(!month){showToast("Pilih bulan terlebih dahulu",C.w);return;}
    if(type==="ops"){
      const rows = ops.filter((o:any)=>o.date?.startsWith(month))
        .map((o:any,i:number)=>[i+1,o.date||"",o.time||"",o.patient||"",o.age||"",o.rm||"",o.opType||"",o.diagnosis||"",o.procedure||"",o.room||"",o.surgeon||"",o.anesthesiologist||"",o.assistantNurse||"",o.circulatingNurse||"",o.anesthesiaNurse||"",o.onloopNurse||"",o.rrKatim||"",o.status||""]);
      if(!rows.length){showToast("Tidak ada data operasi "+ymLabel(month),C.w);return;}
      const h=["No","Tanggal","Jam","Pasien","Usia","RM","Jenis","Diagnosa","Tindakan","Kamar","Operator","Anestesi","Asisten","Instrumen","P.Anestesi","Onloop","RR/Katim","Status"];
      buildAndDownloadXlsx("Operasi "+ymLabel(month),h,rows,`Operasi_${month}.xlsx`);
      showToast(`✓ Excel Operasi ${ymLabel(month)} diunduh (${rows.length} data)`,C.s);
    } else if(type==="lembur"){
      const rows: (string|number)[][] = [];
      (lemburPegawai||[]).forEach((p:any)=>{
        const k=`${p.id}_${month}`;
        const rec=(lemburData||{})[k];
        (rec?.entries||[]).forEach((e:any)=>rows.push([p.name,p.nik||"",ymLabel(month),e.tanggalAwal||"",e.tanggalAkhir||"",e.jamMasuk||"",e.jamKeluar||"",e.keperluanLembur||"",e.keterangan||"",e.ttd||""]));
      });
      if(!rows.length){showToast("Tidak ada data lembur "+ymLabel(month),C.w);return;}
      const h=["Nama","NIK","Bulan","Tgl Awal","Tgl Akhir","Masuk","Keluar","Keperluan","Keterangan","TTD"];
      buildAndDownloadXlsx("Lembur "+ymLabel(month),h,rows,`Lembur_${month}.xlsx`);
      showToast(`✓ Excel Lembur ${ymLabel(month)} diunduh`,C.s);
    } else {
      const filtered=(monitoringEntries||[]).filter((e:any)=>e.tanggal?.startsWith(month));
      if(!filtered.length){showToast("Tidak ada data monitoring "+ymLabel(month),C.w);return;}
      const h=["Tanggal","Jam","Suhu (°C)","Kelembaban (%)","Status","Petugas"];
      const rows=filtered.map((e:any)=>[e.tanggal||"",e.jam||"",e.suhu??""  ,e.kelembaban??"",monIsOK(e.suhu,e.kelembaban,monitoringCfg)?"SESUAI":"TIDAK SESUAI",e.petugas||""]);
      buildAndDownloadXlsx("Monitoring "+ymLabel(month),h,rows,`Monitoring_${month}.xlsx`);
      showToast(`✓ Excel Monitoring ${ymLabel(month)} diunduh`,C.s);
    }
  };

  /* ── Download dari Supabase berdasarkan bulan ── */
  const dlFromSupabase = async (month:string, type:"ops"|"lembur"|"monitoring") => {
    if(!month){showToast("Pilih bulan terlebih dahulu",C.w);return;}
    if(!supaCfg.url||!supaCfg.anonKey){showToast("Supabase belum dikonfigurasi",C.w);return;}
    setDlCloudBusy(true);
    try {
      const {createClient} = await import("@supabase/supabase-js");
      const supa = createClient(supaCfg.url, supaCfg.anonKey);
      const tbl = supaCfg.tableName||"hospital_backup";
      const {data, error} = await supa.from(tbl).select("*").order("created_at",{ascending:false}).limit(1);
      if(error||!data?.length) throw new Error(error?.message||"Data Supabase tidak ditemukan");
      const raw = data[0];
      if(type==="ops"){
        const allOps:any[] = raw.ops||raw.data?.ops||[];
        const rows = allOps.filter((o:any)=>o.date?.startsWith(month))
          .map((o:any,i:number)=>[i+1,o.date||"",o.time||"",o.patient||"",o.age||"",o.rm||"",o.opType||"",o.diagnosis||"",o.procedure||"",o.room||"",o.surgeon||"",o.anesthesiologist||"",o.assistantNurse||"",o.circulatingNurse||"",o.anesthesiaNurse||"",o.onloopNurse||"",o.rrKatim||"",o.status||""]);
        if(!rows.length){showToast("Tidak ada data operasi "+ymLabel(month)+" di Supabase",C.w);return;}
        const h=["No","Tanggal","Jam","Pasien","Usia","RM","Jenis","Diagnosa","Tindakan","Kamar","Operator","Anestesi","Asisten","Instrumen","P.Anestesi","Onloop","RR/Katim","Status"];
        buildAndDownloadXlsx("Operasi "+ymLabel(month),h,rows,`Operasi_Supa_${month}.xlsx`);
        showToast(`✓ ${rows.length} data operasi ${ymLabel(month)} dari Supabase`,C.s);
      } else if(type==="lembur"){
        const ld:any = raw.lemburData||raw.data?.lemburData||{};
        const lp:any[] = raw.lemburPegawai||raw.data?.lemburPegawai||[];
        const rows: (string|number)[][] = [];
        (lp.length?lp:(lemburPegawai||[])).forEach((p:any)=>{
          const k=`${p.id}_${month}`;
          const rec=ld[k];
          (rec?.entries||[]).forEach((e:any)=>rows.push([p.name,p.nik||"",ymLabel(month),e.tanggalAwal||"",e.tanggalAkhir||"",e.jamMasuk||"",e.jamKeluar||"",e.keperluanLembur||"",e.keterangan||"",e.ttd||""]));
        });
        if(!rows.length){showToast("Tidak ada data lembur "+ymLabel(month)+" di Supabase",C.w);return;}
        const h=["Nama","NIK","Bulan","Tgl Awal","Tgl Akhir","Masuk","Keluar","Keperluan","Keterangan","TTD"];
        buildAndDownloadXlsx("Lembur "+ymLabel(month),h,rows,`Lembur_Supa_${month}.xlsx`);
        showToast(`✓ Data lembur ${ymLabel(month)} dari Supabase diunduh`,C.s);
      } else {
        const me:any[] = raw.monitoringEntries||raw.data?.monitoringEntries||[];
        const cfg2:any = raw.monitoringCfg||raw.data?.monitoringCfg||monitoringCfg;
        const filtered=me.filter((e:any)=>e.tanggal?.startsWith(month));
        if(!filtered.length){showToast("Tidak ada data monitoring "+ymLabel(month)+" di Supabase",C.w);return;}
        const h=["Tanggal","Jam","Suhu (°C)","Kelembaban (%)","Status","Petugas"];
        const rows=filtered.map((e:any)=>[e.tanggal||"",e.jam||"",e.suhu??"",e.kelembaban??"",monIsOK(e.suhu,e.kelembaban,cfg2)?"SESUAI":"TIDAK SESUAI",e.petugas||""]);
        buildAndDownloadXlsx("Monitoring "+ymLabel(month),h,rows,`Monitoring_Supa_${month}.xlsx`);
        showToast(`✓ ${rows.length} data monitoring ${ymLabel(month)} dari Supabase`,C.s);
      }
    } catch(e:any){
      showToast("Gagal ambil dari Supabase: "+(e.message||e),C.d);
    } finally {
      setDlCloudBusy(false);
    }
  };

  /* ── Download dari Dropbox berdasarkan bulan ── */
  const dlFromDropbox = async (month:string, type:"ops"|"lembur"|"monitoring") => {
    if(!month){showToast("Pilih bulan terlebih dahulu",C.w);return;}
    if(!dbxCfg.token){showToast("Dropbox belum dikonfigurasi",C.w);return;}
    setDlCloudBusy(true);
    try {
      const path = dbxCfg.filePath||"/hospital_backup.json";
      const res = await fetch("https://content.dropboxapi.com/2/files/download",{
        method:"POST",
        headers:{"Authorization":"Bearer "+dbxCfg.token,"Dropbox-API-Arg":JSON.stringify({path})},
      });
      if(!res.ok) throw new Error("HTTP "+res.status);
      const raw = await res.json();
      if(type==="ops"){
        const allOps:any[] = raw.ops||raw.data?.ops||[];
        const rows = allOps.filter((o:any)=>o.date?.startsWith(month))
          .map((o:any,i:number)=>[i+1,o.date||"",o.time||"",o.patient||"",o.age||"",o.rm||"",o.opType||"",o.diagnosis||"",o.procedure||"",o.room||"",o.surgeon||"",o.anesthesiologist||"",o.assistantNurse||"",o.circulatingNurse||"",o.anesthesiaNurse||"",o.onloopNurse||"",o.rrKatim||"",o.status||""]);
        if(!rows.length){showToast("Tidak ada data operasi "+ymLabel(month)+" di Dropbox",C.w);return;}
        const h=["No","Tanggal","Jam","Pasien","Usia","RM","Jenis","Diagnosa","Tindakan","Kamar","Operator","Anestesi","Asisten","Instrumen","P.Anestesi","Onloop","RR/Katim","Status"];
        buildAndDownloadXlsx("Operasi "+ymLabel(month),h,rows,`Operasi_Dbx_${month}.xlsx`);
        showToast(`✓ ${rows.length} data operasi ${ymLabel(month)} dari Dropbox`,C.s);
      } else if(type==="lembur"){
        const ld:any = raw.lemburData||raw.data?.lemburData||{};
        const lp:any[] = raw.lemburPegawai||raw.data?.lemburPegawai||[];
        const rows:(string|number)[][]=[];
        (lp.length?lp:(lemburPegawai||[])).forEach((p:any)=>{
          const k=`${p.id}_${month}`;
          const rec=ld[k];
          (rec?.entries||[]).forEach((e:any)=>rows.push([p.name,p.nik||"",ymLabel(month),e.tanggalAwal||"",e.tanggalAkhir||"",e.jamMasuk||"",e.jamKeluar||"",e.keperluanLembur||"",e.keterangan||"",e.ttd||""]));
        });
        if(!rows.length){showToast("Tidak ada data lembur "+ymLabel(month)+" di Dropbox",C.w);return;}
        const h=["Nama","NIK","Bulan","Tgl Awal","Tgl Akhir","Masuk","Keluar","Keperluan","Keterangan","TTD"];
        buildAndDownloadXlsx("Lembur "+ymLabel(month),h,rows,`Lembur_Dbx_${month}.xlsx`);
        showToast(`✓ Data lembur ${ymLabel(month)} dari Dropbox diunduh`,C.s);
      } else {
        const me:any[] = raw.monitoringEntries||raw.data?.monitoringEntries||[];
        const cfg2:any = raw.monitoringCfg||raw.data?.monitoringCfg||monitoringCfg;
        const filtered=me.filter((e:any)=>e.tanggal?.startsWith(month));
        if(!filtered.length){showToast("Tidak ada data monitoring "+ymLabel(month)+" di Dropbox",C.w);return;}
        const h=["Tanggal","Jam","Suhu (°C)","Kelembaban (%)","Status","Petugas"];
        const rows=filtered.map((e:any)=>[e.tanggal||"",e.jam||"",e.suhu??"",e.kelembaban??"",monIsOK(e.suhu,e.kelembaban,cfg2)?"SESUAI":"TIDAK SESUAI",e.petugas||""]);
        buildAndDownloadXlsx("Monitoring "+ymLabel(month),h,rows,`Monitoring_Dbx_${month}.xlsx`);
        showToast(`✓ ${rows.length} data monitoring ${ymLabel(month)} dari Dropbox`,C.s);
      }
    } catch(e:any){
      showToast("Gagal ambil dari Dropbox: "+(e.message||e),C.d);
    } finally {
      setDlCloudBusy(false);
    }
  };

  return (
    <div>
      <div style={{display:"flex",gap:0,marginBottom:14,background:C.gBg,borderRadius:12,padding:4,flexWrap:"wrap"}}>
        {[{k:"notif",l:"Log"},{k:"jadwal",l:"Jadwal"},{k:"arsip",l:"Arsip"},{k:"unduh",l:"📥 Unduh"},{k:"import",l:"Import"},{k:"audit",l:"📋 Audit"},{k:"pengaturan",l:"⚙ Setelan"}].map(t=>(
          <button key={t.k} onClick={()=>setSub(t.k)} style={{flex:1,minWidth:56,padding:"8px 2px",borderRadius:10,border:"none",background:sub===t.k?C.white:"none",color:sub===t.k?C.p:C.g,fontSize:10,fontWeight:sub===t.k?700:500,cursor:"pointer",boxShadow:sub===t.k?"0 1px 6px rgba(0,0,0,.1)":"none"}}>{t.l}</button>
        ))}
      </div>

      {sub==="notif" && (
        <div>
          {notifs.length===0 && <Card><div style={{textAlign:"center",padding:24,color:C.tL,fontSize:13}}>Belum ada log notifikasi</div></Card>}
          {notifs.map((n:any)=>{
            const cfg=nCfg[n.type]||{c:C.g,bg:C.gBg};
            return (
              <Card key={n.id}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
                  <Bdg label={n.label} color={cfg.c} bg={cfg.bg}/>
                  <div style={{fontSize:11,color:C.tL}}>{n.sentAt}</div>
                </div>
                <div style={{fontSize:13,fontWeight:700,color:C.t}}>{n.patient}</div>
                <div style={{fontSize:12,color:C.tL,marginBottom:n.message?6:0}}>{n.procedure}</div>
                {n.message && (
                  <details>
                    <summary style={{fontSize:12,color:C.p,fontWeight:700,cursor:"pointer",outline:"none"}}>Lihat Isi Pesan</summary>
                    <div style={{marginTop:8,background:"#F9FAFB",borderRadius:8,padding:"10px 12px",fontSize:12,lineHeight:1.8,whiteSpace:"pre-wrap",border:"1px solid #DDE3EA",fontFamily:"inherit",maxHeight:200,overflowY:"auto"}}>{n.message}</div>
                  </details>
                )}
              </Card>
            );
          })}
        </div>
      )}

      {sub==="jadwal" && (
        <div>
          <Btn full outline color={C.p} onClick={()=>{
            if(!ops.length){showToast("Belum ada data jadwal",C.w);return;}
            const h=["No","Tanggal","Jam","Pasien","Usia","RM","Jenis","Diagnosa","Tindakan","Kamar","Operator","Anestesi","Asisten","PERAWAT Instrumen","P.Anestesi","Onloop","RR/Katim","Alergi","Gol.Darah","Status"];
            const rows=ops.map((op:any,i:number)=>[i+1,op.date,op.time,op.patient,op.age||"",op.rm||"",OT[op.opType as keyof typeof OT||"elektif"]&&OT[op.opType as keyof typeof OT||"elektif"].label||"Elektif",op.diagnosis,op.procedure,op.room,op.surgeon,op.anesthesiologist||"",op.assistantNurse||"",op.circulatingNurse||"",op.anesthesiaNurse||"",op.onloopNurse||"",op.rrKatim||"",op.allergy||"",op.bloodType,op.status]);
            downloadBlob([h,...rows].map((r:any[])=>r.map((c:any)=>'"'+String(c||"").replace(/"/g,'""')+'"').join(",")).join("\n"),"JadwalOK_PantiRini_"+todayDate()+".csv");
            showToast("✓ CSV berhasil diunduh",C.s);
          }} style={{marginBottom:12}}>⬇ Unduh CSV Jadwal</Btn>
          {ops.length===0 && <Card><div style={{textAlign:"center",padding:20,color:C.tL,fontSize:13}}>Belum ada data jadwal</div></Card>}
          {ops.map((op:any)=>{
            const sc=STS[op.status as keyof typeof STS]||STS.scheduled, ot=OT[op.opType as keyof typeof OT||"elektif"];
            return (
              <Card key={op.id} style={{padding:"12px 16px",marginBottom:6}}>
                <div style={{fontSize:13,fontWeight:700,color:C.t}}>{op.patient||"—"}</div>
                <div style={{fontSize:12,color:C.tL}}>{op.procedure} · {fD(op.date)} {op.time} WIB</div>
                <div style={{marginTop:6,display:"flex",gap:5,flexWrap:"wrap"}}>
                  <Bdg label={sc.l} color={sc.c} bg={sc.bg}/>
                  <Bdg label={ot&&ot.label||"Elektif"} color={ot&&ot.c||C.i} bg={ot&&ot.bg||C.iBg}/>
                  {op.allergy&&op.allergy!=="Tidak Ada"&&<Bdg label="Alergi" color={C.d} bg={C.dBg}/>}
                </div>
              </Card>
            );
          })}
        </div>
      )}

      {sub==="arsip" && (
        <div>
          {archive.length===0 && <Card><div style={{textAlign:"center",padding:24,color:C.tL,fontSize:13,lineHeight:1.6}}>Belum ada arsip.<br/>Data diarsipkan otomatis setiap pergantian hari.</div></Card>}
          {archive.map((a:any)=>(
            <details key={a.id} style={{marginBottom:10}}>
              <summary style={{background:C.white,borderRadius:12,padding:"14px 16px",cursor:"pointer",border:"1px solid #DDE3EA",listStyle:"none",fontWeight:700,fontSize:14,color:C.t}}>
                📅 {a.label} — {(a.ops||[]).length} operasi
                <div style={{fontSize:11,color:C.tL,fontWeight:400,marginTop:3}}>{a.archivedAt}</div>
              </summary>
              <div style={{paddingTop:8,paddingLeft:6}}>
                {(a.ops||[]).map((op:any)=>(
                  <Card key={op.id} style={{padding:"10px 14px",marginBottom:6}}>
                    <div style={{fontSize:13,fontWeight:700}}>{op.patient}</div>
                    <div style={{fontSize:12,color:C.tL}}>{op.procedure} · {op.time} WIB</div>
                  </Card>
                ))}
              </div>
            </details>
          ))}
        </div>
      )}

      {sub==="unduh" && (
        <div>
          {/* ── DOWNLOAD EXCEL PER BULAN dari Lokal / Supabase / Dropbox ── */}
          <Card style={{background:"#EDE7F6",border:"1px solid #B39DDB",marginBottom:14}}>
            <SH label="📥 Download Excel per Bulan" color="#4527A0"/>
            <div style={{fontSize:12,color:"#5E35B1",marginBottom:12,lineHeight:1.7}}>
              Pilih bulan &amp; jenis data, lalu download langsung sebagai file Excel (.xlsx).
              Bisa dari data <b>Lokal</b> (perangkat ini), <b>Supabase</b> (cloud), atau <b>Dropbox</b> (cloud).
            </div>
            <div style={{background:"#F3E5F5",borderRadius:8,padding:"8px 12px",marginBottom:12,fontSize:12,color:"#6A1B9A",lineHeight:1.6}}>
              ℹ️ <b>Tentang penyimpanan:</b> Data tersimpan di browser (localStorage) perangkat ini saja.
              Untuk akses dari HP/tablet lain, aktifkan <b>Backup Supabase</b> atau <b>Sinkronisasi Real-time</b> di tab ⚙ Setelan, lalu pilih "Download dari Supabase" di bawah.
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
              <LF label="Bulan &amp; Tahun">
                <input style={{...iS}} type="month" value={dlCloudMonth} onChange={e=>setDlCloudMonth(e.target.value)}/>
              </LF>
              <LF label="Jenis Data">
                <select style={{...iS,appearance:"none"} as React.CSSProperties} value={dlCloudType} onChange={e=>setDlCloudType(e.target.value as any)}>
                  <option value="ops">📋 Jadwal Operasi</option>
                  <option value="lembur">⏱ Data Lembur</option>
                  <option value="monitoring">🌡 Monitoring Suhu</option>
                </select>
              </LF>
            </div>
            <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
              <Btn full onClick={()=>dlFromLocal(dlCloudMonth,dlCloudType)} style={{flex:1,background:"#7B1FA2",color:"#fff",border:"none",minWidth:120}}>
                💾 Lokal
              </Btn>
              {supaCfg?.url&&supaCfg?.anonKey&&(
                <Btn full onClick={()=>dlFromSupabase(dlCloudMonth,dlCloudType)} style={{flex:1,background:dlCloudBusy?"#aaa":"#3ECF8E",color:"#fff",border:"none",minWidth:120,opacity:dlCloudBusy?.6:1,cursor:dlCloudBusy?"not-allowed":"pointer"}}>
                  {dlCloudBusy?"⏳ Mengunduh...":"☁️ Supabase"}
                </Btn>
              )}
              {dbxCfg?.token&&(
                <Btn full onClick={()=>dlFromDropbox(dlCloudMonth,dlCloudType)} style={{flex:1,background:dlCloudBusy?"#aaa":"#0061FF",color:"#fff",border:"none",minWidth:120,opacity:dlCloudBusy?.6:1,cursor:dlCloudBusy?"not-allowed":"pointer"}}>
                  {dlCloudBusy?"⏳ Mengunduh...":"📦 Dropbox"}
                </Btn>
              )}
            </div>
            {!supaCfg?.url&&!dbxCfg?.token&&(
              <div style={{marginTop:10,fontSize:11,color:"#7B1FA2",background:"#F3E5F5",borderRadius:8,padding:"6px 10px"}}>
                💡 Tombol Supabase &amp; Dropbox akan muncul setelah konfigurasi diisi di tab ⚙ Setelan.
              </div>
            )}
          </Card>

          {/* ── FITUR 3: Download Arsip Lama dengan Filter ── */}
          <Card style={{background:"#E8F5E9",border:"1px solid #A5D6A7",marginBottom:14}}>
            <SH label="📁 Download Data Arsip / Lama" color="#2E7D32"/>
            <div style={{fontSize:12,color:"#2E7D32",marginBottom:12,lineHeight:1.6}}>
              Unduh data historis berdasarkan tahun, bulan, dan jenis laporan.
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
              <LF label="Tahun">
                <select style={{...iS,appearance:"none"} as React.CSSProperties} value={arsipTahun} onChange={e=>setArsipTahun(e.target.value)}>
                  <option value="">Semua Tahun</option>
                  {allYears.map(y=><option key={y} value={y}>{y}</option>)}
                </select>
              </LF>
              <LF label="Bulan">
                <select style={{...iS,appearance:"none"} as React.CSSProperties} value={arsipBulan} onChange={e=>setArsipBulan(e.target.value)}>
                  <option value="">Semua Bulan</option>
                  {MONTHS.map(m=><option key={m} value={m}>{MONTH_LABELS[m]}</option>)}
                </select>
              </LF>
            </div>
            <LF label="Jenis Laporan">
              <select style={{...iS,appearance:"none"} as React.CSSProperties} value={arsipTipe} onChange={e=>setArsipTipe(e.target.value)}>
                <option value="laporan_operasi">Laporan Operasi</option>
                <option value="jadwal_operasi">Jadwal Operasi</option>
                <option value="data_lembur">Data Lembur</option>
                <option value="monitoring_suhu">Monitoring Suhu &amp; Kelembaban</option>
              </select>
            </LF>
            <div style={{background:"#F1F8E9",borderRadius:8,padding:"8px 12px",marginBottom:12,fontSize:12,color:"#388E3C"}}>
              {arsipTipe!=="data_lembur"
                ? <>{filteredOpsArsip.length} jadwal operasi ditemukan{arsipTahun?` (${arsipTahun}${arsipBulan?" "+MONTH_LABELS[arsipBulan]:""})`:" (semua data)"}</>
                : <>Data lembur{arsipTahun?` (${arsipTahun}${arsipBulan?" "+MONTH_LABELS[arsipBulan]:""})`:" (semua bulan)"}</>
              }
            </div>
            <div style={{display:"flex",gap:8}}>
              <Btn full onClick={downloadArsipExcel} style={{flex:1,background:"#217346",color:"#fff",border:"none"}}>
                📊 Download Excel
              </Btn>
              <Btn full onClick={downloadArsipWord} style={{flex:1,background:"#2B5797",color:"#fff",border:"none"}}>
                📄 Download Word
              </Btn>
            </div>
          </Card>

          {/* ── FITUR 2: Download File Tracking per Kategori ── */}
          <Card style={{background:"#E3F2FD",border:"1px solid #90CAF9",marginBottom:14}}>
            <SH label="📋 Download File Tracking Operasi" color="#1565C0"/>
            <div style={{fontSize:12,color:"#1565C0",marginBottom:12,lineHeight:1.6}}>
              Unduh semua data jadwal operasi sebagai file tracking (Excel atau Word).
              Total: <b>{ops.length} jadwal</b> operasi.
            </div>
            <div style={{display:"flex",gap:8}}>
              <Btn full onClick={()=>downloadTrackingOps("excel")} style={{flex:1,background:"#1565C0",color:"#fff",border:"none"}}>
                📊 Excel Tracking Operasi
              </Btn>
              <Btn full onClick={()=>downloadTrackingOps("word")} style={{flex:1,background:"#0D47A1",color:"#fff",border:"none"}}>
                📄 Word Tracking Operasi
              </Btn>
            </div>
          </Card>

          <Card style={{background:"#E0F2F1",border:"1px solid #80CBC4"}}>
            <SH label="📋 Download File Tracking Lembur" color="#00695C"/>
            <div style={{fontSize:12,color:"#00695C",marginBottom:12,lineHeight:1.6}}>
              Unduh semua data lembur dari semua pegawai sebagai file tracking (Excel atau Word).
              Total pegawai: <b>{(lemburPegawai||[]).length}</b>.
            </div>
            <div style={{display:"flex",gap:8}}>
              <Btn full onClick={()=>downloadTrackingLembur("excel")} style={{flex:1,background:"#00695C",color:"#fff",border:"none"}}>
                📊 Excel Tracking Lembur
              </Btn>
              <Btn full onClick={()=>downloadTrackingLembur("word")} style={{flex:1,background:"#004D40",color:"#fff",border:"none"}}>
                📄 Word Tracking Lembur
              </Btn>
            </div>
          </Card>
        </div>
      )}

      {sub==="import" && (
        <ImportExcel ops={ops} setOps={setOps} showToast={showToast}/>
      )}

      {sub==="audit" && (
        <div>
          <div style={{fontSize:13,fontWeight:700,color:C.t,marginBottom:10}}>📋 Riwayat Perubahan ({auditLog.length})</div>
          {auditLog.length===0 && <Card><div style={{textAlign:"center",padding:24,color:C.tL,fontSize:13}}>Belum ada riwayat. Perubahan pada jadwal akan tercatat di sini.</div></Card>}
          {[...auditLog].reverse().map((a:any)=>(
            <div key={a.id} style={{background:C.white,borderRadius:12,padding:"12px 16px",marginBottom:8,border:"1px solid #DDE3EA",borderLeft:`4px solid ${a.action==="Hapus"?C.d:a.action==="Edit"?C.i:C.s}`}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
                <span style={{fontSize:12,fontWeight:800,color:a.action==="Hapus"?C.d:a.action==="Edit"?C.i:C.s}}>{a.action==="Hapus"?"🗑 HAPUS":a.action==="Edit"?"✏️ EDIT":"✚ TAMBAH"}</span>
                <span style={{fontSize:11,color:C.tL}}>{a.time}</span>
              </div>
              <div style={{fontSize:13,fontWeight:700,color:C.t}}>{a.patient}</div>
              <div style={{fontSize:12,color:C.tL}}>{a.detail}</div>
            </div>
          ))}
        </div>
      )}

      {sub==="pengaturan" && (
        <div>
          {/* Mode Privasi */}
          <Card>
            <SH label="Mode Privasi" color="#4A148C"/>
            <Toggle value={privacyMode} onChange={(v:boolean)=>{setPrivacyMode(v);showToast(v?"🔒 Mode Privasi aktif":"Mode Privasi dimatikan",v?"#7B1FA2":C.g);}} color="#4A148C" label="🔒 Sembunyikan Data Pasien" sub="Nama & RM disamarkan di layar (pesan WA tetap asli)"/>
            {privacyMode && <div style={{background:"#F3E5F5",borderRadius:8,padding:"8px 12px",marginTop:10,fontSize:12,color:"#7B1FA2",lineHeight:1.6}}>Mode aktif. Aman untuk presentasi, audit & dokumentasi sistem.</div>}
          </Card>

          {/* Supabase Backup */}
          <Card>
            <SH label="☁ Backup ke Supabase" color="#3ECF8E"/>
            <div style={{background:"#ECFDF5",borderRadius:8,padding:"10px 12px",marginBottom:14,fontSize:12,color:"#065F46",lineHeight:1.7,border:"1px solid #A7F3D0"}}>
              <b>Cara setup:</b><br/>
              1. Buat project di <b>supabase.com</b><br/>
              2. Buat tabel: <code style={{background:"#D1FAE5",padding:"1px 5px",borderRadius:4}}>kamar_bedah_backup</code> dengan kolom <code style={{background:"#D1FAE5",padding:"1px 5px",borderRadius:4}}>created_at</code> (timestamptz) dan <code style={{background:"#D1FAE5",padding:"1px 5px",borderRadius:4}}>data</code> (text)<br/>
              3. Copy URL & Anon Key dari Project Settings → API
            </div>
            <LF label="Supabase Project URL">
              <input style={iS} placeholder="https://xxxx.supabase.co" value={supaCfg.url} onChange={(e:any)=>setSupaCfg((p:SupabaseConfig)=>({...p,url:e.target.value.trim()}))} autoComplete="off" spellCheck={false}/>
            </LF>
            <LF label="Supabase Anon Key">
              <input style={iS} type="password" placeholder="eyJhbGciOi..." value={supaCfg.anonKey} onChange={(e:any)=>setSupaCfg((p:SupabaseConfig)=>({...p,anonKey:e.target.value.trim()}))} autoComplete="off"/>
            </LF>
            <Toggle
              value={supaCfg.autoBackup}
              onChange={(v:boolean)=>setSupaCfg((p:SupabaseConfig)=>({...p,autoBackup:v}))}
              color="#3ECF8E"
              label="🔄 Auto-Backup JSON Otomatis"
              sub={`Backup otomatis setiap ${supaCfg.backupInterval} menit (semua data: ops, lembur, monitoring)`}
            />
            {supaCfg.autoBackup && (
              <div style={{marginTop:10}}>
                <LF label={`Interval Backup (menit) — saat ini: ${supaCfg.backupInterval} menit`}>
                  <input style={iS} type="number" min="5" max="1440" value={supaCfg.backupInterval} onChange={(e:any)=>setSupaCfg((p:SupabaseConfig)=>({...p,backupInterval:Math.max(5,parseInt(e.target.value)||60)}))}/>
                </LF>
              </div>
            )}
            <Toggle
              value={supaCfg.realtimeBackup}
              onChange={(v:boolean)=>setSupaCfg((p:SupabaseConfig)=>({...p,realtimeBackup:v}))}
              color="#059669"
              label="⚡ Backup Realtime ke Supabase"
              sub="Setiap ada perubahan data, backup dikirim ke Supabase dalam 5 detik"
            />
            <Toggle
              value={supaCfg.autoExcelBackup}
              onChange={(v:boolean)=>setSupaCfg((p:SupabaseConfig)=>({...p,autoExcelBackup:v}))}
              color="#0369A1"
              label="📊 Auto-Backup Excel 24 Jam ke Supabase"
              sub="Upload Excel Operasi, Lembur & Monitoring otomatis setiap 24 jam via Supabase Storage"
            />
            {supaCfg.lastBackup && (
              <div style={{background:"#ECFDF5",borderRadius:8,padding:"8px 12px",marginBottom:10,fontSize:12,color:"#065F46",border:"1px solid #A7F3D0"}}>
                🕐 Backup terakhir: {supaCfg.lastBackup}
              </div>
            )}
            {supaStatus && (
              <div style={{background:supaStatus.ok?"#ECFDF5":"#FEF2F2",borderRadius:8,padding:"8px 12px",marginBottom:10,fontSize:12,color:supaStatus.ok?"#065F46":C.d,border:`1px solid ${supaStatus.ok?"#A7F3D0":C.dL}`}}>
                {supaStatus.msg}
              </div>
            )}
            <Btn full color="#3ECF8E" onClick={onSupaBackup} disabled={supaBackingUp||!supaCfg.url||!supaCfg.anonKey} style={{marginTop:4}}>
              {supaBackingUp?"⟳ Memproses...":"☁ Backup Semua Data Sekarang ke Supabase"}
            </Btn>
            {/* ── Restore dari Supabase ── */}
            <div style={{marginTop:12,background:"#F0FDF4",borderRadius:10,padding:"12px 14px",border:"1px solid #BBF7D0"}}>
              <div style={{fontSize:12,fontWeight:700,color:"#065F46",marginBottom:6}}>⬆ Pulihkan dari Supabase (backup terakhir)</div>
              <div style={{fontSize:11,color:"#047857",marginBottom:10,lineHeight:1.6}}>
                Ambil data dari backup Supabase paling baru. Pilih kategori data yang ingin dipulihkan, atau pulihkan semua sekaligus.
              </div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:8}}>
                <Btn full outline color="#065F46" onClick={onSupaRestoreOps} disabled={supaBackingUp||!supaCfg.url||!supaCfg.anonKey} style={{flex:1,minWidth:130}}>
                  🗓 Pulihkan Operasi
                </Btn>
                <Btn full outline color="#065F46" onClick={onSupaRestoreLembur} disabled={supaBackingUp||!supaCfg.url||!supaCfg.anonKey} style={{flex:1,minWidth:130}}>
                  ⏰ Pulihkan Lembur
                </Btn>
                <Btn full outline color="#0369A1" onClick={onSupaRestoreMonitoring} disabled={supaBackingUp||!supaCfg.url||!supaCfg.anonKey} style={{flex:1,minWidth:130}}>
                  🌡 Pulihkan Monitoring
                </Btn>
              </div>
              <Btn full color="#065F46" onClick={onSupaRestoreAll} disabled={supaBackingUp||!supaCfg.url||!supaCfg.anonKey} style={{background:"#065F46",color:"#fff",border:"none",width:"100%"}}>
                ♻ Pulihkan SEMUA Data dari Supabase
              </Btn>
            </div>
          </Card>

          {/* Dropbox Backup */}
          <Card>
            <SH label="📦 Backup ke Dropbox" color="#0061FF"/>
            <div style={{background:"#EFF6FF",borderRadius:8,padding:"10px 12px",marginBottom:14,fontSize:12,color:"#1D4ED8",lineHeight:1.8,border:"1px solid #BFDBFE"}}>
              <b>Cara setup:</b><br/>
              1. Buka <b>dropbox.com/developers/apps</b><br/>
              2. Buat App → <b>Scoped access</b> → <b>Full Dropbox</b><br/>
              3. Di tab <b>Permissions</b>: aktifkan <code style={{background:"#DBEAFE",padding:"1px 4px",borderRadius:3}}>files.content.write</code> &amp; <code style={{background:"#DBEAFE",padding:"1px 4px",borderRadius:3}}>files.content.read</code><br/>
              4. Di tab <b>Settings</b> → Generated access token → copy ke sini
            </div>
            <LF label="Dropbox Access Token">
              <input style={iS} type="password" placeholder="sl...." value={dbxCfg.token} onChange={(e:any)=>setDbxCfg((p:DropboxConfig)=>({...p,token:e.target.value.trim()}))} autoComplete="off"/>
            </LF>
            <LF label="Path File Backup di Dropbox">
              <input style={iS} placeholder="/KamarBedahPantiRini/backup.json" value={dbxCfg.path} onChange={(e:any)=>setDbxCfg((p:DropboxConfig)=>({...p,path:e.target.value.trim()}))} spellCheck={false}/>
            </LF>
            <Toggle
              value={dbxCfg.autoBackup}
              onChange={(v:boolean)=>setDbxCfg((p:DropboxConfig)=>({...p,autoBackup:v}))}
              color="#0061FF"
              label="⏰ Auto-Backup Terjadwal"
              sub={`Backup otomatis setiap ${dbxCfg.backupInterval} menit`}
            />
            {dbxCfg.autoBackup && (
              <div style={{marginTop:8,marginBottom:4}}>
                <LF label={`Interval (menit) — saat ini: ${dbxCfg.backupInterval} menit`}>
                  <input style={iS} type="number" min="5" max="1440" value={dbxCfg.backupInterval} onChange={(e:any)=>setDbxCfg((p:DropboxConfig)=>({...p,backupInterval:Math.max(5,parseInt(e.target.value)||30)}))}/>
                </LF>
              </div>
            )}
            <Toggle
              value={dbxCfg.realtimeBackup}
              onChange={(v:boolean)=>setDbxCfg((p:DropboxConfig)=>({...p,realtimeBackup:v}))}
              color="#0061FF"
              label="⚡ Cadangkan Realtime (otomatis)"
              sub="Setiap ada perubahan data, backup dikirim ke Dropbox dalam 3 detik"
            />
            {dbxCfg.realtimeBackup && (
              <div style={{background:"#EFF6FF",borderRadius:8,padding:"8px 12px",marginBottom:8,fontSize:12,color:"#1D4ED8",border:"1px solid #BFDBFE",display:"flex",alignItems:"center",gap:6}}>
                <span style={{fontSize:16}}>⚡</span>
                <span>Realtime aktif — setiap perubahan jadwal &amp; lembur otomatis tersimpan ke Dropbox</span>
              </div>
            )}
            {/* ── Auto Excel every 24h ── */}
            <Toggle
              value={dbxCfg.autoExcelBackup}
              onChange={(v:boolean)=>setDbxCfg((p:DropboxConfig)=>({...p,autoExcelBackup:v}))}
              color="#3730A3"
              label="📊 Auto-Backup Excel setiap 24 Jam"
              sub="Upload Excel Operasi &amp; Lembur otomatis ke Dropbox setiap 24 jam"
            />
            {dbxCfg.autoExcelBackup && (
              <div style={{background:"#EEF2FF",borderRadius:8,padding:"10px 12px",marginBottom:8,fontSize:12,color:"#3730A3",border:"1px solid #C7D2FE"}}>
                <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:4}}>
                  <span style={{fontSize:15}}>📊</span>
                  <b>Auto Excel aktif</b>
                </div>
                <div style={{color:"#4F46E5",lineHeight:1.7}}>
                  Setiap 24 jam, file <b>Auto_Operasi_YYYY-MM-DD.xlsx</b>, <b>Auto_Lembur_YYYY-MM-DD.xlsx</b>, dan <b>Auto_Monitoring_YYYY-MM-DD.xlsx</b> di-upload ke Dropbox secara otomatis.
                  {dbxCfg.lastExcelBackup && (
                    <div style={{marginTop:4}}>🕐 Excel terakhir diupload: <b>{dbxCfg.lastExcelBackup}</b></div>
                  )}
                  {!dbxCfg.lastExcelBackup && (
                    <div style={{marginTop:4,color:"#6366F1"}}>⏳ Menunggu waktu upload pertama (setiap 24 jam sekali)</div>
                  )}
                </div>
              </div>
            )}
            {dbxCfg.lastBackup && (
              <div style={{background:"#EFF6FF",borderRadius:8,padding:"8px 12px",marginBottom:10,fontSize:12,color:"#1D4ED8",border:"1px solid #BFDBFE"}}>
                🕐 Backup JSON terakhir: <b>{dbxCfg.lastBackup}</b>
              </div>
            )}
            {dbxStatus && (
              <div style={{background:dbxStatus.ok?"#ECFDF5":"#FEF2F2",borderRadius:8,padding:"8px 12px",marginBottom:10,fontSize:12,color:dbxStatus.ok?"#065F46":C.d,border:`1px solid ${dbxStatus.ok?"#A7F3D0":C.dL}`}}>
                {dbxStatus.msg}
              </div>
            )}
            <Btn full color="#0061FF" onClick={onDbxBackup} disabled={dbxBacking||!dbxCfg.token} style={{background:"#0061FF",color:"#fff",border:"none",marginTop:4}}>
              {dbxBacking?"⟳ Menyimpan...":"📦 Backup Semua Data (JSON) ke Dropbox"}
            </Btn>
            {/* ── FITUR 1: Backup per kategori sebagai Excel ── */}
            <div style={{marginTop:10,background:"#EEF2FF",borderRadius:10,padding:"12px 14px"}}>
              <div style={{fontSize:12,fontWeight:700,color:"#3730A3",marginBottom:8}}>📊 Backup per Kategori (Excel) ke Dropbox</div>
              <div style={{fontSize:11,color:"#6366F1",marginBottom:10,lineHeight:1.6}}>
                Backup data Operasi atau Lembur secara terpisah sebagai file Excel (.xlsx) ke Dropbox.
                File tersimpan di folder yang sama dengan path backup utama.
              </div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                <Btn full outline color="#3730A3" onClick={onDbxBackupOpsXls} disabled={dbxBacking||!dbxCfg.token} style={{flex:1,minWidth:140}}>
                  🗓 Backup Operasi (Excel)
                </Btn>
                <Btn full outline color="#3730A3" onClick={onDbxBackupLemburXls} disabled={dbxBacking||!dbxCfg.token} style={{flex:1,minWidth:140}}>
                  ⏰ Backup Lembur (Excel)
                </Btn>
                <Btn full outline color="#0369A1" onClick={onDbxBackupMonitoringXls} disabled={dbxBacking||!dbxCfg.token} style={{flex:1,minWidth:140}}>
                  🌡 Backup Monitoring (Excel)
                </Btn>
              </div>
            </div>
            <div style={{marginTop:10}}>
              <div style={{fontSize:11,fontWeight:700,color:"#1D4ED8",marginBottom:6}}>⬇ Pulihkan dari Dropbox (pilih data yang ingin dipulihkan):</div>
              <div style={{display:"flex",gap:8}}>
                <Btn full outline color="#0061FF" onClick={onDbxRestoreOps} disabled={dbxBacking||!dbxCfg.token} style={{flex:1}}>
                  🗓 Jadwal Operasi
                </Btn>
                <Btn full outline color="#0061FF" onClick={onDbxRestoreLembur} disabled={dbxBacking||!dbxCfg.token} style={{flex:1}}>
                  ⏰ Data Lembur
                </Btn>
              </div>
            </div>
          </Card>

          {/* Backup JSON lokal */}
          <Card>
            <SH label="💾 Backup & Recovery Lokal" color={C.w}/>
            <div style={{fontSize:12,color:C.tL,marginBottom:12,lineHeight:1.6}}>Ekspor semua data lokal sebagai file JSON untuk backup darurat.</div>
            <Btn full outline color={C.w} onClick={()=>{
              const data={exportedAt:fNow(),ops,archive,notifs};
              downloadBlob(JSON.stringify(data,null,2),"BACKUP_OR_PantiRini_"+todayDate()+".json","application/json");
              showToast("✓ File JSON berhasil diunduh",C.s);
            }}>⬇ Ekspor Semua Data (JSON)</Btn>
          </Card>

          {/* Supabase Realtime */}
          <Card>
            <SH label="📡 Sinkronisasi Real-time" color="#1565C0"/>
            <div style={{background:"#E3F2FD",borderRadius:8,padding:"10px 12px",marginBottom:14,fontSize:12,color:"#1565C0",lineHeight:1.7,border:"1px solid #90CAF9"}}>
              Sinkronisasi data antar HP/komputer secara langsung menggunakan Supabase Realtime. Aktifkan setelah mengisi URL & Anon Key Supabase di atas.
            </div>
            <Toggle value={rtEnabled} onChange={(v:boolean)=>setRtEnabled(v)} color="#1565C0" label="📡 Aktifkan Sinkronisasi Real-time" sub="Semua perubahan jadwal akan disinkron ke perangkat lain"/>
            {rtStatus && (
              <div style={{marginTop:10,background:rtStatus==="online"?"#E8F5E9":"#FFF3E0",borderRadius:8,padding:"8px 12px",fontSize:12,color:rtStatus==="online"?"#2E7D32":"#E65100",fontWeight:700,border:`1px solid ${rtStatus==="online"?"#A5D6A7":"#FFCC80"}`}}>
                {rtStatus==="online"?"🟢 Terhubung — data disinkron secara real-time":rtStatus==="connecting"?"🟡 Menghubungkan...":"🔴 Tidak terhubung — periksa URL & Anon Key"}
              </div>
            )}
          </Card>

          <Card style={{background:C.sBg,border:"1px solid #2E7D3233"}}>
            <div style={{fontSize:12,color:C.s,lineHeight:2.1}}>
              ✓ <b>Zero dummy data</b> — sistem dimulai kosong<br/>
              ✓ <b>Spasi bebas</b> saat mengetik (sanitize hanya saat simpan)<br/>
              ✓ <b>PERAWAT Instrumen</b> mengganti Perawat Sirkuler<br/>
              ✓ <b>Supabase Backup</b> — manual & auto terjadwal<br/>
              ✓ <b>Pagination {PAGE_SIZE}/halaman</b> — tahan 100+ data<br/>
              ✓ <b>Auto-lock</b> setelah {LOCK_MS/60000} menit idle<br/>
              ✓ <b>Statistik lengkap</b> per bulan, kamar, dokter
            </div>
          </Card>
        </div>
      )}
    </div>
  );
}

/* ─── VIEW LEMBUR ────────────────────────────────────────────────────── */
const BULAN_ID = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
const nowYM = () => { const d=new Date(); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`; };
const ymLabel = (ym:string) => { const [y,m]=ym.split("-"); return `${BULAN_ID[parseInt(m)-1]} ${y}`; };

function ViewLembur({lemburPegawai, setLemburPegawai, lemburData, setLemburData, showToast, supaCfg}: any) {
  const [sub, setSub]         = useState<"catat"|"rekap"|"pegawai">("catat");
  const [selPeg, setSelPeg]   = useState<string>("");
  const [selYM,  setSelYM]    = useState(nowYM());
  const [editRow, setEditRow] = useState<string|null>(null);
  const [rowForm, setRowForm] = useState<any>({});
  const [kepRuang, setKepRuang]   = useState("");
  const [kepBidang, setKepBidang] = useState("");
  const [pegForm,  setPegForm]    = useState({name:"", nik:""});
  const [rekapYM,  setRekapYM]    = useState(nowYM());
  const fileRef = useRef<HTMLInputElement>(null);
  const [supaLemburBusy, setSupaLemburBusy] = useState(false);
  const [supaLemburYM, setSupaLemburYM] = useState(nowYM());

  const peg = lemburPegawai.find((p:any)=>p.id===selPeg);
  const key = selPeg && selYM ? `${selPeg}_${selYM}` : null;
  const rec: any = (key && lemburData[key]) || {entries:[], kepRuang:"", kepBidang:""};
  const entries: any[] = rec.entries || [];

  /* Sync kepRuang/kepBidang from saved record */
  useEffect(()=>{
    if(key && lemburData[key]){
      setKepRuang(lemburData[key].kepRuang||"");
      setKepBidang(lemburData[key].kepBidang||"");
    } else {
      setKepRuang(""); setKepBidang("");
    }
  },[key]);

  const saveEntry = () => {
    if(!rowForm.tanggalAwal||!rowForm.jamMasuk||!rowForm.jamKeluar){
      showToast("Tanggal awal lembur, jam masuk & keluar wajib diisi",C.d); return;
    }
    const upd = editRow
      ? entries.map((e:any)=>e.id===editRow?{...e,...rowForm}:e)
      : [...entries,{id:gId(),...rowForm,no:entries.length+1}];
    setLemburData((p:any)=>({...p,[key!]:{...rec,entries:upd}}));
    setEditRow(null); setRowForm({});
    showToast("✓ Baris disimpan",C.s);
  };

  const delEntry = (id:string) => {
    const upd = entries.filter((e:any)=>e.id!==id).map((e:any,i:number)=>({...e,no:i+1}));
    setLemburData((p:any)=>({...p,[key!]:{...rec,entries:upd}}));
    showToast("Baris dihapus");
  };

  const saveSigs = () => {
    if(!key) return;
    setLemburData((p:any)=>({...p,[key]:{...rec,kepRuang,kepBidang,savedAt:fNow()}}));
    showToast("✓ Tanda tangan tersimpan",C.s);
  };

  const handleImportJadwal = (file:File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target!.result as ArrayBuffer),{type:"array"});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data: any[][] = XLSX.utils.sheet_to_json(ws,{header:1});
        if(data.length<2){showToast("File kosong",C.d);return;}
        const header=(data[0] as string[]).map(h=>String(h||"").trim().toLowerCase());
        const jadwalIdx=header.findIndex(h=>h.includes("jadwal")||h.includes("dinas")||h.includes("tanggal")||h.includes("date"));
        const masukIdx =header.findIndex(h=>h.includes("masuk")||h.includes("in")||h.includes("datang"));
        const keluarIdx=header.findIndex(h=>h.includes("keluar")||h.includes("out")||h.includes("pulang"));
        const kepIdx   =header.findIndex(h=>h.includes("keperluan")||h.includes("lembur")||h.includes("pekerjaan"));
        const ketIdx   =header.findIndex(h=>h.includes("keterangan")||h.includes("ket")||h.includes("catatan"));
        let newEntries: any[] = [...entries];
        for(let i=1;i<data.length;i++){
          const row=data[i] as any[];
          if(!row||row.every((c:any)=>!c)) continue;
          const jadwal=jadwalIdx>=0?String(row[jadwalIdx]||"").trim():"";
          if(!jadwal) continue;
          /* detect month match */
          const dateMatch = jadwal.match(/(\d{4})-(\d{2})-(\d{2})/)||jadwal.match(/(\d{2})\/(\d{2})\/(\d{4})/);
          const existing = newEntries.find((e:any)=>(e.tanggalAwal||e.jadwalDinas)===jadwal);
          if(existing) continue;
          newEntries.push({
            id:gId(), no:newEntries.length+1,
            tanggalAwal:jadwal, tanggalAkhir:"",
            jamMasuk:masukIdx>=0?String(row[masukIdx]||"").trim():"",
            jamKeluar:keluarIdx>=0?String(row[keluarIdx]||"").trim():"",
            keperluanLembur:kepIdx>=0?String(row[kepIdx]||"").trim():"",
            keterangan:ketIdx>=0?String(row[ketIdx]||"").trim():"",
            ttd:"",
          });
        }
        setLemburData((p:any)=>({...p,[key!]:{...rec,entries:newEntries}}));
        showToast(`✓ ${newEntries.length-entries.length} baris diimport`,C.s);
      } catch(err){showToast("Gagal membaca file",C.d);}
    };
    reader.readAsArrayBuffer(file);
  };

  const downloadExcel = () => {
    if(!peg){showToast("Pilih pegawai terlebih dahulu",C.w);return;}
    if(!entries.length){showToast("Belum ada data lembur",C.w);return;}
    const ws = XLSX.utils.aoa_to_sheet([]);
    const merge = (r1:number,c1:number,r2:number,c2:number)=>({s:{r:r1,c:c1},e:{r:r2,c:c2}});
    const cell = (v:any,bold?:boolean,center?:boolean,border?:boolean,bg?:string,fontSize?:number):XLSX.CellObject=>{
      const s:any={font:{bold:!!bold,sz:fontSize||10},alignment:{horizontal:center?"center":"left",vertical:"center",wrapText:true}};
      if(border) s.border={top:{style:"thin"},bottom:{style:"thin"},left:{style:"thin"},right:{style:"thin"}};
      if(bg) s.fill={fgColor:{rgb:bg},patternType:"solid"};
      return {v,t:"s",s} as XLSX.CellObject;
    };
    const setC=(r:number,c:number,v:any,bold?:boolean,center?:boolean,border?:boolean,bg?:string,fs?:number)=>{
      const addr=XLSX.utils.encode_cell({r,c});
      ws[addr]=cell(v,bold,center,border,bg,fs);
    };
    /* Header */
    setC(0,0,"PENCATATAN LEMBUR",true,true,false,undefined,14);
    setC(1,0,HOSPITAL,false,true);
    setC(2,0,"",false,false);
    setC(3,0,"Nama Karyawan",true); setC(3,2,`: ${peg.name}`);
    setC(4,0,"NIK / NIP",true);    setC(4,2,`: ${peg.nik||"-"}`);
    setC(5,0,"Bulan",true);         setC(5,2,`: ${ymLabel(selYM)}`);
    setC(6,0,"",false,false);
    /* Table header */
    const TH=["No","Tgl Awal Lembur","Tgl Akhir Lembur","Jam Absen Masuk","Jam Absen Keluar","Keperluan Lembur","Keterangan","Tanda Tangan"];
    TH.forEach((h,i)=>setC(7,i,h,true,true,true,"004D40"));
    /* Table rows */
    entries.forEach((e:any,i:number)=>{
      const r=8+i;
      [e.no,e.tanggalAwal||e.jadwalDinas||"",e.tanggalAkhir||"",e.jamMasuk,e.jamKeluar,e.keperluanLembur,e.keterangan,e.ttd||""].forEach((v,c)=>setC(r,c,String(v||""),false,c===0||c===3||c===4,true,i%2===0?"F0FFF8":undefined));
    });
    /* Signatures */
    const sigRow=9+entries.length;
    setC(sigRow,0,"",false,false);
    setC(sigRow+1,0,`Mengetahui,`,true,true);
    setC(sigRow+1,4,`Menyetujui,`,true,true);
    setC(sigRow+2,0,"Kepala Ruang",false,true);
    setC(sigRow+2,4,"Kepala Bidang",false,true);
    setC(sigRow+6,0,kepRuang||"( _________________ )",false,true);
    setC(sigRow+6,4,kepBidang||"( _________________ )",false,true);
    /* Merges */
    ws["!merges"]=[
      merge(0,0,0,7), merge(1,0,1,7), merge(2,0,2,7),
      merge(3,0,3,1), merge(3,2,3,7),
      merge(4,0,4,1), merge(4,2,4,7),
      merge(5,0,5,1), merge(5,2,5,7),
      merge(sigRow+1,0,sigRow+1,3), merge(sigRow+1,4,sigRow+1,7),
      merge(sigRow+2,0,sigRow+2,3), merge(sigRow+2,4,sigRow+2,7),
      merge(sigRow+6,0,sigRow+6,3), merge(sigRow+6,4,sigRow+6,7),
    ];
    ws["!cols"]=[{wch:4},{wch:16},{wch:16},{wch:16},{wch:16},{wch:26},{wch:22},{wch:16}];
    ws["!ref"]=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:sigRow+7,c:7}});
    const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Lembur");
    XLSX.writeFile(wb,`Lembur_${peg.name.replace(/\s+/g,"_")}_${selYM}.xlsx`);
    showToast("✓ File Excel berhasil diunduh",C.s);
  };

  /* All months that have data for selected staff */
  const availableMonths = selPeg
    ? Object.keys(lemburData).filter(k=>k.startsWith(selPeg+"_")).map(k=>k.replace(selPeg+"_","")).sort().reverse()
    : [];

  /* Rekap: compute per-employee summary for rekapYM */
  const rekapRows = lemburPegawai.map((p:any)=>{
    const k=`${p.id}_${rekapYM}`;
    const rec2:any = lemburData[k]||{entries:[]};
    const ents:any[] = rec2.entries||[];
    const totalMins = ents.reduce((s:number,e:any)=>{
      if(!e.jamMasuk||!e.jamKeluar) return s;
      const [h1,m1]=e.jamMasuk.split(":").map(Number);
      const [h2,m2]=e.jamKeluar.split(":").map(Number);
      let m=(h2*60+m2)-(h1*60+m1); if(m<0) m+=24*60; return s+(m>0?m:0);
    },0);
    return {id:p.id, name:p.name, nik:p.nik||"—", jumlahHari:ents.length, totalMins, entries:ents, kepRuang:rec2.kepRuang||"", kepBidang:rec2.kepBidang||""};
  });

  const downloadRekapExcel = () => {
    if(!rekapRows.length){showToast("Belum ada data pegawai",C.w);return;}
    const ws = XLSX.utils.aoa_to_sheet([]);
    const merge=(r1:number,c1:number,r2:number,c2:number)=>({s:{r:r1,c:c1},e:{r:r2,c:c2}});
    const cell=(v:any,bold?:boolean,center?:boolean,border?:boolean,bg?:string,fs?:number):XLSX.CellObject=>{
      const s:any={font:{bold:!!bold,sz:fs||10,color:bg==="004D40"||bg==="003D33"?{rgb:"FFFFFF"}:undefined},alignment:{horizontal:center?"center":"left",vertical:"center",wrapText:true}};
      if(border) s.border={top:{style:"thin"},bottom:{style:"thin"},left:{style:"thin"},right:{style:"thin"}};
      if(bg) s.fill={fgColor:{rgb:bg},patternType:"solid"};
      return {v,t:"s",s} as XLSX.CellObject;
    };
    const setC=(r:number,c:number,v:any,bold?:boolean,center?:boolean,border?:boolean,bg?:string,fs?:number)=>{
      ws[XLSX.utils.encode_cell({r,c})]=cell(v,bold,center,border,bg,fs);
    };
    /* Title */
    setC(0,0,"REKAP LEMBUR BULANAN",true,true,false,undefined,14);
    setC(1,0,HOSPITAL,false,true);
    setC(2,0,`Bulan: ${ymLabel(rekapYM)}`,false,true);
    setC(3,0,"",false,false);
    /* Header row */
    const TH=["No","Nama Pegawai","NIK / NIP","Jumlah Hari Lembur","Total Jam Lembur","Keterangan"];
    TH.forEach((h,i)=>setC(4,i,h,true,true,true,"004D40",11));
    /* Data rows */
    const merges=[
      merge(0,0,0,5),merge(1,0,1,5),merge(2,0,2,5),merge(3,0,3,5),
    ];
    rekapRows.forEach((r:any,i:number)=>{
      const row=5+i;
      const bg=i%2===0?"F0FFF8":undefined;
      const jam=`${Math.floor(r.totalMins/60)} jam ${r.totalMins%60} menit`;
      setC(row,0,String(i+1),false,true,true,bg);
      setC(row,1,r.name,true,false,true,bg);
      setC(row,2,r.nik,false,false,true,bg);
      setC(row,3,String(r.jumlahHari),false,true,true,bg);
      setC(row,4,r.totalMins>0?jam:"—",false,true,true,bg);
      setC(row,5,r.kepRuang?`Kep. Ruang: ${r.kepRuang}`:"",false,false,true,bg);
    });
    /* Footer total */
    const footRow=5+rekapRows.length;
    const grandMins=rekapRows.reduce((s:number,r:any)=>s+r.totalMins,0);
    const grandHari=rekapRows.reduce((s:number,r:any)=>s+r.jumlahHari,0);
    setC(footRow,0,"TOTAL",true,true,true,"E0F2F1");
    setC(footRow,1,"",false,false,true,"E0F2F1");
    merges.push(merge(footRow,0,footRow,2));
    setC(footRow,2,"",false,false,true,"E0F2F1");
    setC(footRow,3,String(grandHari)+" hari",true,true,true,"E0F2F1");
    setC(footRow,4,`${Math.floor(grandMins/60)} jam ${grandMins%60} menit`,true,true,true,"E0F2F1");
    setC(footRow,5,"",false,false,true,"E0F2F1");
    ws["!merges"]=merges;
    ws["!cols"]=[{wch:4},{wch:28},{wch:18},{wch:20},{wch:22},{wch:28}];
    ws["!ref"]=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:footRow,c:5}});
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Rekap Lembur");
    /* Detail sheets per employee */
    rekapRows.forEach((r:any)=>{
      if(!r.entries.length) return;
      const ws2=XLSX.utils.aoa_to_sheet([]);
      const s2=(row:number,col:number,v:any,bold?:boolean,center?:boolean,border?:boolean,bg?:string)=>{
        const st:any={font:{bold:!!bold,sz:10},alignment:{horizontal:center?"center":"left",vertical:"center"}};
        if(border) st.border={top:{style:"thin"},bottom:{style:"thin"},left:{style:"thin"},right:{style:"thin"}};
        if(bg) st.fill={fgColor:{rgb:bg},patternType:"solid"};
        ws2[XLSX.utils.encode_cell({r:row,c:col})]={v,t:"s",s:st} as XLSX.CellObject;
      };
      s2(0,0,r.name,true,false); s2(0,2,r.nik,false,false);
      ["No","Tgl Awal Lembur","Tgl Akhir Lembur","Jam Masuk","Jam Keluar","Keperluan","Keterangan"].forEach((h,i)=>s2(1,i,h,true,true,true,"004D40"));
      r.entries.forEach((e:any,i:number)=>{
        const row=2+i; const bg=i%2===0?"F0FFF8":undefined;
        [String(e.no),e.tanggalAwal||e.jadwalDinas||"",e.tanggalAkhir||"",e.jamMasuk||"",e.jamKeluar||"",e.keperluanLembur||"",e.keterangan||""].forEach((v,c)=>s2(row,c,v,false,c===0||c===3||c===4,true,bg));
      });
      ws2["!ref"]=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:2+r.entries.length,c:6}});
      ws2["!cols"]=[{wch:4},{wch:16},{wch:16},{wch:14},{wch:14},{wch:24},{wch:20}];
      const safeName=r.name.replace(/[\\/?*[\]:]/g,"").substring(0,28);
      XLSX.utils.book_append_sheet(wb,ws2,safeName);
    });
    XLSX.writeFile(wb,`RekapLembur_${rekapYM}.xlsx`);
    showToast("✓ File rekap berhasil diunduh",C.s);
  };

  const downloadLemburFromSupa = async () => {
    if(!supaCfg?.url||!supaCfg?.anonKey||!supaCfg?.table){showToast("Supabase belum dikonfigurasi. Atur di tab Arsip → Setelan.",C.d);return;}
    setSupaLemburBusy(true);
    try {
      const {createClient} = await import("@supabase/supabase-js");
      const sb = createClient(supaCfg.url, supaCfg.anonKey);
      const {data,error} = await sb.from(supaCfg.table).select("data").order("created_at",{ascending:false}).limit(1);
      if(error||!data?.length) throw new Error(error?.message||"Data tidak ditemukan di Supabase");
      const raw = typeof data[0].data==="string"?JSON.parse(data[0].data):data[0].data;
      const ld:any = raw.lemburData||raw.data?.lemburData||{};
      const lp:any[] = (raw.lemburPegawai||raw.data?.lemburPegawai||[]);
      const src = lp.length?lp:lemburPegawai;
      const wb = XLSX.utils.book_new(); let found=false;
      src.forEach((p:any)=>{
        const k=`${p.id}_${supaLemburYM}`;
        if(ld[k]){
          found=true;
          const ents:any[]=(ld[k].entries||[]);
          const rows=ents.map((e:any,i:number)=>[i+1,e.tanggalAwal||"",e.tanggalAkhir||"",e.jamMasuk||"",e.jamKeluar||"",e.keperluanLembur||"",e.keterangan||""]);
          const ws=XLSX.utils.aoa_to_sheet([["No","Tgl Awal","Tgl Akhir","Jam Masuk","Jam Keluar","Keperluan Lembur","Keterangan"],...rows]);
          ws["!cols"]=[{wch:4},{wch:14},{wch:14},{wch:10},{wch:10},{wch:28},{wch:20}];
          XLSX.utils.book_append_sheet(wb,ws,p.name.slice(0,28));
        }
      });
      if(!found){showToast("Tidak ada data lembur "+ymLabel(supaLemburYM)+" di Supabase",C.w);setSupaLemburBusy(false);return;}
      XLSX.writeFile(wb,`Lembur_Cloud_${supaLemburYM}.xlsx`);
      showToast("✓ Data lembur "+ymLabel(supaLemburYM)+" dari Supabase diunduh",C.s);
    } catch(err:any){showToast("Gagal: "+(err?.message||"Error Supabase"),C.d);}
    setSupaLemburBusy(false);
  };

  return (
    <div>
      {/* Sub-tabs */}
      <div style={{display:"flex",gap:0,marginBottom:14,background:C.gBg,borderRadius:12,padding:4}}>
        {([{k:"catat",l:"📝 Catat Lembur"},{k:"rekap",l:"📊 Rekap Bulanan"},{k:"pegawai",l:"👤 Data Pegawai"}] as {k:"catat"|"rekap"|"pegawai";l:string}[]).map(t=>(
          <button key={t.k} onClick={()=>setSub(t.k as any)} style={{flex:1,padding:"9px 4px",borderRadius:10,border:"none",background:sub===t.k?C.white:"none",color:sub===t.k?C.p:C.g,fontSize:12,fontWeight:sub===t.k?700:500,cursor:"pointer",boxShadow:sub===t.k?"0 1px 6px rgba(0,0,0,.1)":"none"}}>{t.l}</button>
        ))}
      </div>

      {/* ── CATAT LEMBUR ── */}
      {sub==="catat" && (
        <div>
          {/* Pilih Pegawai + Bulan */}
          <Card>
            <SH label="🗓 Pilih Pegawai & Bulan"/>
            <LF label="Pegawai">
              <select style={{...iS,appearance:"none"} as React.CSSProperties} value={selPeg} onChange={e=>setSelPeg(e.target.value)}>
                <option value="">-- Pilih pegawai --</option>
                {lemburPegawai.map((p:any)=><option key={p.id} value={p.id}>{p.name} — {p.nik||"(NIK belum diisi)"}</option>)}
              </select>
            </LF>
            {lemburPegawai.length===0 && <div style={{fontSize:12,color:C.w,background:C.wBg,borderRadius:8,padding:"8px 12px",marginBottom:8}}>⚠ Belum ada data pegawai. Tambahkan di tab "👤 Data Pegawai" dulu.</div>}
            <div style={{display:"flex",gap:10,alignItems:"flex-end"}}>
              <div style={{flex:1}}>
                <LF label="Bulan">
                  <input style={iS} type="month" value={selYM} onChange={e=>setSelYM(e.target.value)}/>
                </LF>
              </div>
              {availableMonths.length>0 && (
                <div style={{flex:1}}>
                  <LF label="Bulan tersimpan">
                    <select style={{...iS,appearance:"none"} as React.CSSProperties} value={selYM} onChange={e=>setSelYM(e.target.value)}>
                      {availableMonths.map(ym=><option key={ym} value={ym}>{ymLabel(ym)} {lemburData[`${selPeg}_${ym}`]?.entries?.length?`(${lemburData[`${selPeg}_${ym}`].entries.length} baris)`:""}</option>)}
                    </select>
                  </LF>
                </div>
              )}
            </div>
          </Card>

          {peg && key && (
            <>
              {/* Header info */}
              <div style={{background:"linear-gradient(135deg,#003D33,#00695C)",borderRadius:14,padding:"16px 20px",marginBottom:14,color:"#fff"}}>
                <div style={{fontSize:14,fontWeight:800,letterSpacing:.3}}>📋 PENCATATAN LEMBUR</div>
                <div style={{fontSize:11,color:"rgba(255,255,255,.65)",marginTop:2}}>{HOSPITAL}</div>
                <div style={{marginTop:10,display:"grid",gridTemplateColumns:"1fr 1fr",gap:"4px 16px",fontSize:12}}>
                  <div><span style={{color:"rgba(255,255,255,.6)"}}>Nama: </span><b>{peg.name}</b></div>
                  <div><span style={{color:"rgba(255,255,255,.6)"}}>NIK/NIP: </span><b>{peg.nik||"—"}</b></div>
                  <div><span style={{color:"rgba(255,255,255,.6)"}}>Bulan: </span><b>{ymLabel(selYM)}</b></div>
                  <div><span style={{color:"rgba(255,255,255,.6)"}}>Total: </span><b>{entries.length} baris</b></div>
                </div>
              </div>

              {/* Actions */}
              <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap"}}>
                <Btn onClick={()=>{setEditRow(null);setRowForm({tanggalAwal:"",tanggalAkhir:"",jamMasuk:"",jamKeluar:"",keperluanLembur:"",keterangan:"",ttd:""});}} style={{flex:1}}>✚ Tambah Baris</Btn>
                <button onClick={()=>fileRef.current?.click()} style={{flex:1,background:C.iBg,border:`1px solid ${C.i}`,borderRadius:10,color:C.i,fontSize:12,fontWeight:700,padding:"10px 14px",cursor:"pointer",fontFamily:"inherit"}}>📥 Import Jadwal dari Excel</button>
                <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={e=>{const f=e.target.files?.[0];if(f)handleImportJadwal(f);e.target.value="";}}/>
                <Btn onClick={downloadExcel} style={{flex:1,background:"#1565C0",color:"#fff"}}>⬇ Download Excel</Btn>
              </div>

              {/* Add/Edit row form */}
              {(editRow!==null || rowForm.tanggalAwal!==undefined) && rowForm.tanggalAwal!==undefined && (
                <Card style={{marginBottom:14,background:"#F0FFF8",border:`1.5px solid ${C.p}33`}}>
                  <SH label={editRow?"✏️ Edit Baris":"✚ Tambah Baris Lembur"} color={C.p}/>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                    <LF label="Tanggal Awal Lembur" req>
                      <input style={iS} type="date" value={rowForm.tanggalAwal||""} onChange={e=>setRowForm((p:any)=>({...p,tanggalAwal:e.target.value}))}/>
                    </LF>
                    <LF label="Tanggal Akhir Lembur">
                      <input style={iS} type="date" value={rowForm.tanggalAkhir||""} onChange={e=>setRowForm((p:any)=>({...p,tanggalAkhir:e.target.value}))}/>
                    </LF>
                    <LF label="Jam Absen Masuk" req>
                      <input style={iS} type="time" value={rowForm.jamMasuk||""} onChange={e=>setRowForm((p:any)=>({...p,jamMasuk:e.target.value}))}/>
                    </LF>
                    <LF label="Jam Absen Keluar" req>
                      <input style={iS} type="time" value={rowForm.jamKeluar||""} onChange={e=>setRowForm((p:any)=>({...p,jamKeluar:e.target.value}))}/>
                    </LF>
                    <LF label="Keperluan Lembur">
                      <input style={iS} placeholder="cth: Menyiapkan alat operasi" value={rowForm.keperluanLembur||""} onChange={e=>setRowForm((p:any)=>({...p,keperluanLembur:e.target.value}))}/>
                    </LF>
                    <LF label="Keterangan">
                      <input style={iS} placeholder="Catatan tambahan..." value={rowForm.keterangan||""} onChange={e=>setRowForm((p:any)=>({...p,keterangan:e.target.value}))}/>
                    </LF>
                    <LF label="Tanda Tangan (nama/paraf)">
                      <input style={iS} placeholder="Nama atau paraf" value={rowForm.ttd||""} onChange={e=>setRowForm((p:any)=>({...p,ttd:e.target.value}))}/>
                    </LF>
                  </div>
                  <div style={{display:"flex",gap:10,marginTop:8}}>
                    <Btn full onClick={saveEntry} style={{flex:2}}>✓ Simpan</Btn>
                    <Btn full outline color={C.g} onClick={()=>{setEditRow(null);setRowForm({});}} style={{flex:1}}>Batal</Btn>
                  </div>
                </Card>
              )}

              {/* Table */}
              <Card style={{padding:0,overflow:"hidden"}}>
                <div style={{background:"#004D40",padding:"12px 16px"}}>
                  <div style={{fontSize:13,fontWeight:800,color:"#fff"}}>Tabel Lembur — {ymLabel(selYM)}</div>
                  <div style={{fontSize:11,color:"rgba(255,255,255,.6)",marginTop:2}}>{peg.name} · NIK: {peg.nik||"—"}</div>
                </div>
                {entries.length===0 ? (
                  <div style={{textAlign:"center",padding:"32px 16px",color:C.tL}}>
                    <div style={{fontSize:32,marginBottom:8}}>📋</div>
                    <div style={{fontSize:13,fontWeight:600}}>Belum ada data lembur bulan ini</div>
                    <div style={{fontSize:12,marginTop:4}}>Klik "Tambah Baris" atau "Import Jadwal dari Excel"</div>
                  </div>
                ) : (
                  <div style={{overflowX:"auto"}}>
                    <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:600}}>
                      <thead>
                        <tr style={{background:"#E0F2F1"}}>
                          {["No","Tgl Awal Lembur","Tgl Akhir Lembur","Masuk","Keluar","Durasi","Keperluan Lembur","Keterangan","TTD","Aksi"].map(h=>(
                            <th key={h} style={{padding:"8px 10px",textAlign:"left",color:"#004D40",fontWeight:700,borderBottom:"2px solid #004D40",whiteSpace:"nowrap"}}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {entries.map((e:any,i:number)=>{
                          const dur = e.jamMasuk&&e.jamKeluar ? (()=>{
                            const [h1,m1]=e.jamMasuk.split(":").map(Number);
                            const [h2,m2]=e.jamKeluar.split(":").map(Number);
                            let mins=(h2*60+m2)-(h1*60+m1);
                            if(mins<0) mins+=24*60;
                            if(mins===0) return "—";
                            return `${Math.floor(mins/60)}j ${mins%60}m`;
                          })() : "—";
                          return (
                            <tr key={e.id} style={{background:i%2===0?"#FAFBFD":"#F0FFF8",borderBottom:`1px solid ${C.b}`}}>
                              <td style={{padding:"8px 10px",textAlign:"center",fontWeight:700,color:C.p}}>{e.no}</td>
                              <td style={{padding:"8px 10px",fontWeight:600,color:C.t}}>{e.tanggalAwal||e.jadwalDinas||"—"}</td>
                              <td style={{padding:"8px 10px",fontWeight:600,color:C.t}}>{e.tanggalAkhir||"—"}</td>
                              <td style={{padding:"8px 10px",color:"#1565C0",fontWeight:600}}>{e.jamMasuk||"—"}</td>
                              <td style={{padding:"8px 10px",color:"#B71C1C",fontWeight:600}}>{e.jamKeluar||"—"}</td>
                              <td style={{padding:"8px 10px",color:C.s,fontWeight:700,whiteSpace:"nowrap"}}>{dur}</td>
                              <td style={{padding:"8px 10px",color:C.t}}>{e.keperluanLembur||"—"}</td>
                              <td style={{padding:"8px 10px",color:C.tL}}>{e.keterangan||"—"}</td>
                              <td style={{padding:"8px 10px",color:C.g,fontStyle:e.ttd?"normal":"italic"}}>{e.ttd||"·"}</td>
                              <td style={{padding:"8px 10px",whiteSpace:"nowrap"}}>
                                <button onClick={()=>{setEditRow(e.id);setRowForm({...e});}} style={{background:C.iBg,border:`1px solid ${C.i}`,borderRadius:6,color:C.i,padding:"3px 8px",cursor:"pointer",fontSize:11,marginRight:4}}>✏</button>
                                <button onClick={()=>delEntry(e.id)} style={{background:C.dBg,border:`1px solid ${C.dL}`,borderRadius:6,color:C.d,padding:"3px 8px",cursor:"pointer",fontSize:11}}>🗑</button>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                      {entries.length>0 && (()=>{
                        const totalMins=entries.reduce((sum:number,e:any)=>{
                          if(!e.jamMasuk||!e.jamKeluar) return sum;
                          const [h1,m1]=e.jamMasuk.split(":").map(Number);
                          const [h2,m2]=e.jamKeluar.split(":").map(Number);
                          let m=(h2*60+m2)-(h1*60+m1); if(m<0) m+=24*60; return sum+(m>0?m:0);
                        },0);
                        return <tfoot><tr style={{background:"#E0F2F1"}}><td colSpan={4} style={{padding:"8px 10px",fontWeight:700,color:"#004D40",textAlign:"right"}}>Total Lembur:</td><td colSpan={5} style={{padding:"8px 10px",fontWeight:800,color:C.s}}>{Math.floor(totalMins/60)} jam {totalMins%60} menit</td></tr></tfoot>;
                      })()}
                    </table>
                  </div>
                )}
              </Card>

              {/* Tanda Tangan */}
              <Card style={{marginTop:14}}>
                <SH label="✍ Tanda Tangan Pejabat"/>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
                  <LF label="Nama Kepala Ruang">
                    <input style={iS} placeholder="Nama Kepala Ruang..." value={kepRuang} onChange={e=>setKepRuang(e.target.value)}/>
                  </LF>
                  <LF label="Nama Kepala Bidang">
                    <input style={iS} placeholder="Nama Kepala Bidang..." value={kepBidang} onChange={e=>setKepBidang(e.target.value)}/>
                  </LF>
                </div>
                <Btn full onClick={saveSigs} style={{marginTop:4}}>💾 Simpan Tanda Tangan</Btn>
              </Card>

              {/* Download */}
              <Card style={{marginTop:14,background:"#E3F2FD",border:"1px solid #90CAF9"}}>
                <div style={{fontSize:13,fontWeight:700,color:"#1565C0",marginBottom:8}}>📥 Unduh Dokumen Lembur</div>
                <div style={{fontSize:12,color:"#1565C0",marginBottom:12,lineHeight:1.6}}>Download file Excel berformat resmi dengan header nama/NIK, tabel lembur, total durasi, dan kolom tanda tangan pejabat.</div>
                <Btn full onClick={downloadExcel} style={{background:"#1565C0",color:"#fff"}}>⬇ Download Excel — {peg.name} ({ymLabel(selYM)})</Btn>
              </Card>
            </>
          )}

          {!selPeg && lemburPegawai.length>0 && (
            <Card><div style={{textAlign:"center",padding:24,color:C.tL}}>
              <div style={{fontSize:32,marginBottom:8}}>👤</div>
              <div style={{fontSize:13,fontWeight:600}}>Pilih pegawai untuk melihat atau mencatat lembur</div>
            </div></Card>
          )}
        </div>
      )}

      {/* ── REKAP BULANAN ── */}
      {sub==="rekap" && (
        <div>
          <Card>
            <SH label="📊 Rekap Lembur Bulanan — Semua Pegawai"/>
            <div style={{display:"flex",gap:10,alignItems:"flex-end",flexWrap:"wrap"}}>
              <div style={{flex:1,minWidth:160}}>
                <LF label="Pilih Bulan">
                  <input style={iS} type="month" value={rekapYM} onChange={e=>setRekapYM(e.target.value)}/>
                </LF>
              </div>
              <Btn onClick={downloadRekapExcel} style={{background:"#1565C0",color:"#fff",whiteSpace:"nowrap"}}>⬇ Download Excel Rekap</Btn>
            </div>
          </Card>

          {/* ── Supabase cloud download ── */}
          <Card style={{background:"#EEF2FF",border:"1.5px solid #6366F133",marginBottom:14}}>
            <SH label="☁️ Download Lembur dari Supabase" color="#4F46E5"/>
            <div style={{fontSize:12,color:"#64748B",marginBottom:10}}>Unduh data lembur langsung dari backup cloud Supabase (tidak bergantung pada data perangkat ini). Data akan tersinkron dari semua device.</div>
            <div style={{display:"flex",gap:10,alignItems:"flex-end",flexWrap:"wrap"}}>
              <div style={{flex:1,minWidth:160}}>
                <LF label="Pilih Bulan"><input style={iS} type="month" value={supaLemburYM} onChange={e=>setSupaLemburYM(e.target.value)}/></LF>
              </div>
              <Btn onClick={downloadLemburFromSupa} style={{background:supaLemburBusy?"#6366F188":"#4F46E5",color:"#fff",whiteSpace:"nowrap"}} disabled={supaLemburBusy}>
                {supaLemburBusy?"⏳ Mengunduh...":"☁️ Download dari Supabase"}
              </Btn>
            </div>
            {!supaCfg?.url && <div style={{fontSize:11,color:"#EF4444",marginTop:8}}>⚠ Supabase belum dikonfigurasi — atur di tab Arsip → Setelan.</div>}
          </Card>

          {lemburPegawai.length===0 ? (
            <Card><div style={{textAlign:"center",padding:32,color:C.tL}}>
              <div style={{fontSize:32,marginBottom:8}}>👤</div>
              <div style={{fontSize:13,fontWeight:600}}>Belum ada data pegawai</div>
              <div style={{fontSize:12,marginTop:4}}>Tambahkan pegawai di tab "👤 Data Pegawai" terlebih dahulu</div>
            </div></Card>
          ) : (
            <>
              {/* Summary cards */}
              <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginBottom:14}}>
                {[
                  {label:"Total Pegawai",val:String(lemburPegawai.length)+" orang",icon:"👥",color:C.p},
                  {label:"Total Hari Lembur",val:String(rekapRows.reduce((s:number,r:any)=>s+r.jumlahHari,0))+" hari",icon:"📅",color:C.s},
                  {label:"Total Jam Lembur",val:(()=>{const m=rekapRows.reduce((s:number,r:any)=>s+r.totalMins,0);return `${Math.floor(m/60)}j ${m%60}m`;})(),icon:"⏱",color:"#1565C0"},
                ].map(c=>(
                  <div key={c.label} style={{background:C.white,borderRadius:12,padding:"12px 14px",boxShadow:"0 1px 6px rgba(0,0,0,.06)",textAlign:"center"}}>
                    <div style={{fontSize:22,marginBottom:4}}>{c.icon}</div>
                    <div style={{fontSize:18,fontWeight:800,color:c.color}}>{c.val}</div>
                    <div style={{fontSize:10,color:C.tL,marginTop:2}}>{c.label}</div>
                  </div>
                ))}
              </div>

              {/* Table */}
              <Card style={{padding:0,overflow:"hidden"}}>
                <div style={{background:"#004D40",padding:"12px 16px"}}>
                  <div style={{fontSize:13,fontWeight:800,color:"#fff"}}>Rekap Lembur — {ymLabel(rekapYM)}</div>
                  <div style={{fontSize:11,color:"rgba(255,255,255,.6)",marginTop:2}}>{HOSPITAL}</div>
                </div>
                <div style={{overflowX:"auto"}}>
                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:500}}>
                    <thead>
                      <tr style={{background:"#E0F2F1"}}>
                        {["No","Nama Pegawai","NIK / NIP","Hari Lembur","Total Jam","Detail"].map(h=>(
                          <th key={h} style={{padding:"9px 12px",textAlign:"left",color:"#004D40",fontWeight:700,borderBottom:"2px solid #004D40",whiteSpace:"nowrap"}}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {rekapRows.map((r:any,i:number)=>{
                        const jam=r.totalMins>0?`${Math.floor(r.totalMins/60)}j ${r.totalMins%60}m`:"—";
                        return (
                          <tr key={r.id} style={{background:i%2===0?"#FAFBFD":"#F0FFF8",borderBottom:`1px solid ${C.b}`}}>
                            <td style={{padding:"9px 12px",textAlign:"center",fontWeight:700,color:C.p}}>{i+1}</td>
                            <td style={{padding:"9px 12px",fontWeight:700,color:C.t}}>{r.name}</td>
                            <td style={{padding:"9px 12px",color:C.tL}}>{r.nik}</td>
                            <td style={{padding:"9px 12px",textAlign:"center"}}>
                              {r.jumlahHari>0
                                ? <span style={{background:C.sBg,color:C.s,borderRadius:20,padding:"2px 10px",fontWeight:700,fontSize:12}}>{r.jumlahHari} hari</span>
                                : <span style={{color:C.tL}}>—</span>}
                            </td>
                            <td style={{padding:"9px 12px",textAlign:"center",fontWeight:700,color:r.totalMins>0?"#1565C0":C.tL}}>{jam}</td>
                            <td style={{padding:"9px 12px"}}>
                              <button onClick={()=>{setSelPeg(r.id);setSelYM(rekapYM);setSub("catat" as any);}} style={{background:C.pBg,border:`1px solid ${C.p}33`,borderRadius:8,color:C.p,padding:"4px 10px",cursor:"pointer",fontSize:11,fontWeight:600,whiteSpace:"nowrap"}}>
                                {r.jumlahHari>0?"Lihat":"Input"} →
                              </button>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                    <tfoot>
                      <tr style={{background:"#E0F2F1"}}>
                        <td colSpan={3} style={{padding:"9px 12px",fontWeight:700,color:"#004D40",textAlign:"right"}}>TOTAL</td>
                        <td style={{padding:"9px 12px",textAlign:"center",fontWeight:800,color:C.s}}>
                          {rekapRows.reduce((s:number,r:any)=>s+r.jumlahHari,0)} hari
                        </td>
                        <td style={{padding:"9px 12px",textAlign:"center",fontWeight:800,color:"#1565C0"}}>
                          {(()=>{const m=rekapRows.reduce((s:number,r:any)=>s+r.totalMins,0);return `${Math.floor(m/60)}j ${m%60}m`;})()} 
                        </td>
                        <td/>
                      </tr>
                    </tfoot>
                  </table>
                </div>
              </Card>

              {/* Download card */}
              <Card style={{marginTop:14,background:"#E3F2FD",border:"1px solid #90CAF9"}}>
                <div style={{fontSize:13,fontWeight:700,color:"#1565C0",marginBottom:6}}>📥 Unduh Rekap Excel</div>
                <div style={{fontSize:12,color:"#1565C0",marginBottom:10,lineHeight:1.5}}>
                  File Excel berisi 1 sheet rekap ringkasan semua pegawai + sheet detail per pegawai (hanya yang ada datanya).
                </div>
                <Btn full onClick={downloadRekapExcel} style={{background:"#1565C0",color:"#fff"}}>⬇ Download Rekap Lembur — {ymLabel(rekapYM)}</Btn>
              </Card>
            </>
          )}
        </div>
      )}

      {/* ── DATA PEGAWAI ── */}
      {sub==="pegawai" && (
        <div>
          <Card>
            <SH label="✚ Tambah Pegawai"/>
            <div style={{display:"flex",gap:10}}>
              <div style={{flex:2}}>
                <LF label="Nama Lengkap">
                  <input style={iS} placeholder="Nama pegawai..." value={pegForm.name} onChange={e=>setPegForm(p=>({...p,name:e.target.value}))}/>
                </LF>
              </div>
              <div style={{flex:1}}>
                <LF label="NIK / NIP">
                  <input style={iS} placeholder="NIK/NIP..." value={pegForm.nik} onChange={e=>setPegForm(p=>({...p,nik:e.target.value}))}/>
                </LF>
              </div>
            </div>
            <Btn full onClick={()=>{
              if(!pegForm.name.trim()){showToast("Nama wajib diisi",C.d);return;}
              setLemburPegawai((p:any[])=>[...p,{id:gId(),name:pegForm.name.trim(),nik:pegForm.nik.trim()}]);
              setPegForm({name:"",nik:""});
              showToast("✓ Pegawai ditambahkan",C.s);
            }}>✚ Tambah Pegawai</Btn>
          </Card>

          {lemburPegawai.length===0 ? (
            <Card><div style={{textAlign:"center",padding:24,color:C.tL,fontSize:13}}>Belum ada data pegawai</div></Card>
          ) : (
            lemburPegawai.map((p:any)=>{
              const totalBulan=Object.keys(lemburData).filter(k=>k.startsWith(p.id+"_")).length;
              return (
                <Card key={p.id} style={{padding:"12px 16px",marginBottom:8}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                    <div>
                      <div style={{fontSize:14,fontWeight:700,color:C.t}}>{p.name}</div>
                      <div style={{fontSize:12,color:C.tL,marginTop:2}}>NIK/NIP: {p.nik||"—"} · {totalBulan} bulan tercatat</div>
                    </div>
                    <div style={{display:"flex",gap:6}}>
                      <button onClick={()=>{setSelPeg(p.id);setSub("catat");}} style={{background:C.pBg,border:`1px solid ${C.p}33`,borderRadius:8,color:C.p,padding:"5px 12px",cursor:"pointer",fontSize:12,fontWeight:600}}>Lihat Lembur</button>
                      <button onClick={()=>{
                        if(!window.confirm(`Hapus pegawai "${p.name}"?\nSemua data lembur pegawai ini juga akan dihapus.`)) return;
                        setLemburPegawai((prev:any[])=>prev.filter((x:any)=>x.id!==p.id));
                        setLemburData((prev:any)=>{const n={...prev};Object.keys(n).filter(k=>k.startsWith(p.id+"_")).forEach(k=>delete n[k]);return n;});
                        showToast("Pegawai dihapus");
                      }} style={{background:C.dBg,border:`1px solid ${C.dL}`,borderRadius:8,color:C.d,padding:"5px 10px",cursor:"pointer",fontSize:12}}>🗑</button>
                    </div>
                  </div>
                </Card>
              );
            })
          )}
        </div>
      )}
    </div>
  );
}

/* ─── MONITORING SUHU & KELEMBABAN ─────────────────────────────────── */
interface MonitoringEntry { id:string; tanggal:string; jam:string; suhu:number; kelembaban:number; petugas:string; createdAt:string; }
interface MonitoringCfg { suhuMin:number; suhuMax:number; rhMin:number; rhMax:number; lokasiRuang:string; kepalaKamarBedah:string; }
const defaultMonCfg: MonitoringCfg = { suhuMin:19, suhuMax:24, rhMin:45, rhMax:65, lokasiRuang:"Kamar Bedah", kepalaKamarBedah:"" };
const MON_JAMS = ["07:00","14:00","21:00"];
const MON_MONTHS_ID = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
const MC = { pri:"#0369A1", priL:"#0284C7", priBg:"#EFF6FF", ok:"#16A34A", okBg:"#DCFCE7", err:"#DC2626", errBg:"#FEE2E2" };

function monIsOK(suhu:number, rh:number, cfg:MonitoringCfg):boolean {
  return suhu>=cfg.suhuMin && suhu<=cfg.suhuMax && rh>=cfg.rhMin && rh<=cfg.rhMax;
}

function monXlsx(entries:MonitoringEntry[], cfg:MonitoringCfg, ym:string, type:"harian"|"rekap"|"akreditasi", showToast:(m:string,c?:string)=>void) {
  const [ys,ms]=ym.split("-"); const year=+ys; const month=+ms;
  const me=[...entries].filter(e=>e.tanggal.startsWith(ym)).sort((a,b)=>a.tanggal.localeCompare(b.tanggal)||a.jam.localeCompare(b.jam));
  if(!me.length){showToast("Tidak ada data untuk periode ini","#E65100");return;}
  const suhuV=me.map(e=>e.suhu).filter(v=>v>0);
  const rhV=me.map(e=>e.kelembaban).filter(v=>v>0);
  const avgS=suhuV.length?+(suhuV.reduce((a,b)=>a+b,0)/suhuV.length).toFixed(2):0;
  const avgR=rhV.length?+(rhV.reduce((a,b)=>a+b,0)/rhV.length).toFixed(2):0;
  const tidakS=me.filter(e=>!monIsOK(e.suhu,e.kelembaban,cfg)).length;
  const mName=MON_MONTHS_ID[month-1];
  const wb=XLSX.utils.book_new();
  let lastDate="";
  const rows1=me.map(e=>{const ok=monIsOK(e.suhu,e.kelembaban,cfg);const r=[e.tanggal===lastDate?"":e.tanggal,e.jam,e.suhu,e.kelembaban,cfg.suhuMin+"-"+cfg.suhuMax,cfg.rhMin+"-"+cfg.rhMax,ok?"SESUAI":"TIDAK SESUAI",e.petugas];lastDate=e.tanggal;return r;});
  const ws1=XLSX.utils.aoa_to_sheet([["Tanggal","Jam","Suhu (°C)","Kelembaban (%)","Standar Suhu","Standar RH","Status","Petugas"],...rows1]);
  ws1["!cols"]=[{wch:14},{wch:8},{wch:10},{wch:14},{wch:13},{wch:12},{wch:16},{wch:16}];
  XLSX.utils.book_append_sheet(wb,ws1,"Monitoring Harian");
  const ws2=XLSX.utils.aoa_to_sheet([["REKAP MONITORING SUHU & KELEMBABAN",""],[""+mName+" "+year,""],["",""],["Parameter","Hasil"],["Rata-rata Suhu (°C)",avgS],["Rata-rata Kelembaban (%)",avgR],["Jumlah Pengukuran",me.length],["Tidak Sesuai Standar",tidakS],["Kepatuhan (%)",me.length?+((me.length-tidakS)/me.length*100).toFixed(1):0]]);
  ws2["!cols"]=[{wch:28},{wch:14}]; ws2["!merges"]=[{s:{r:0,c:0},e:{r:0,c:1}},{s:{r:1,c:0},e:{r:1,c:1}}];
  XLSX.utils.book_append_sheet(wb,ws2,"Rekap Bulanan");
  if(type==="akreditasi"){
    const ws3=XLSX.utils.aoa_to_sheet([
      ["LAPORAN MONITORING SUHU DAN KELEMBABAN KAMAR BEDAH","","",""],
      ["Bulan: "+mName+" "+year,"","",""],
      [HOSPITAL,"","",""],["","","",""],
      ["Parameter","Nilai","",""],
      ["Rata-rata Suhu (°C)",avgS,"",""],["Rata-rata Kelembaban (%)",avgR,"",""],
      ["Jumlah Pengukuran",me.length,"",""],["Tidak Sesuai Standar",tidakS,"",""],
      ["Kepatuhan (%)",me.length?+((me.length-tidakS)/me.length*100).toFixed(1):0,"",""],
      ["","","",""],
      ["Petugas Monitoring","","Kepala Kamar Bedah",""],
      ["","","",""],["","","",""],["","","",""],
      ["(..........................)","","(..........................)",""],
    ]);
    ws3["!cols"]=[{wch:28},{wch:16},{wch:28},{wch:16}];
    ws3["!merges"]=[{s:{r:0,c:0},e:{r:0,c:3}},{s:{r:1,c:0},e:{r:1,c:3}},{s:{r:2,c:0},e:{r:2,c:3}}];
    XLSX.utils.book_append_sheet(wb,ws3,"Laporan Akreditasi");
  }
  XLSX.writeFile(wb,"Monitoring_"+(type==="harian"?"Harian":type==="rekap"?"Rekap":"Akreditasi")+"_"+ym+".xlsx");
  showToast("✓ Monitoring "+(type==="harian"?"Harian":type==="rekap"?"Rekap":"Akreditasi")+" "+mName+" "+year+" diunduh",MC.ok);
}

function ViewMonitoring({monitoringEntries,setMonitoringEntries,monitoringCfg,setMonitoringCfg,showToast,supaCfg}:any) {
  const [subTab,setSubTab]=useState("harian");
  const [selMonth,setSelMonth]=useState(todayDate().slice(0,7));
  const [supaMonBusy,setSupaMonBusy]=useState(false);
  const [supaMonYM,setSupaMonYM]=useState(todayDate().slice(0,7));
  const [formDate,setFormDate]=useState(todayDate());
  type SL={suhu:string;kelembaban:string;petugas:string};
  const ES:SL={suhu:"",kelembaban:"",petugas:""};
  const mkSlots=():(Record<string,SL>)=>({"07:00":{...ES},"14:00":{...ES},"21:00":{...ES}});
  const [slots,setSlots]=useState<Record<string,SL>>(mkSlots);
  const [cfgForm,setCfgForm]=useState<MonitoringCfg>({...monitoringCfg});
  useEffect(()=>{
    const ns=mkSlots();
    MON_JAMS.forEach(jam=>{const ex=monitoringEntries.find((e:MonitoringEntry)=>e.tanggal===formDate&&e.jam===jam);if(ex) ns[jam]={suhu:String(ex.suhu),kelembaban:String(ex.kelembaban),petugas:ex.petugas};});
    setSlots(ns);
  },[formDate]);
  const handleSave=()=>{
    const nw:MonitoringEntry[]=[];
    MON_JAMS.forEach(jam=>{const s=slots[jam];if(s.suhu&&s.kelembaban) nw.push({id:gId(),tanggal:formDate,jam,suhu:parseFloat(s.suhu),kelembaban:parseFloat(s.kelembaban),petugas:s.petugas,createdAt:fNow()});});
    if(!nw.length){showToast("Isi minimal satu slot suhu/kelembaban","#E65100");return;}
    const filtered=monitoringEntries.filter((e:MonitoringEntry)=>!(e.tanggal===formDate&&MON_JAMS.includes(e.jam)));
    setMonitoringEntries([...filtered,...nw]);
    showToast("✓ "+nw.length+" data monitoring "+formDate+" disimpan",MC.ok);
  };

  const downloadMonFromSupa = async (type:"harian"|"rekap"|"akreditasi") => {
    if(!supaCfg?.url||!supaCfg?.anonKey||!supaCfg?.table){showToast("Supabase belum dikonfigurasi. Atur di tab Arsip → Setelan.",MC.err);return;}
    setSupaMonBusy(true);
    try {
      const {createClient} = await import("@supabase/supabase-js");
      const sb = createClient(supaCfg.url, supaCfg.anonKey);
      const {data,error} = await sb.from(supaCfg.table).select("data").order("created_at",{ascending:false}).limit(1);
      if(error||!data?.length) throw new Error(error?.message||"Data tidak ditemukan di Supabase");
      const raw = typeof data[0].data==="string"?JSON.parse(data[0].data):data[0].data;
      const me:MonitoringEntry[] = (raw.monitoringEntries||raw.data?.monitoringEntries||[]).filter((e:MonitoringEntry)=>e.tanggal?.startsWith(supaMonYM));
      const cfg2:MonitoringCfg = raw.monitoringCfg||raw.data?.monitoringCfg||monitoringCfg;
      if(!me.length){showToast("Tidak ada data monitoring "+supaMonYM+" di Supabase",MC.err);setSupaMonBusy(false);return;}
      monXlsx(me, cfg2, supaMonYM, type, showToast);
      showToast("✓ Monitoring dari Supabase ("+supaMonYM+") diunduh",MC.ok);
    } catch(err:any){showToast("Gagal: "+(err?.message||"Error Supabase"),MC.err);}
    setSupaMonBusy(false);
  };
  const me=monitoringEntries.filter((e:MonitoringEntry)=>e.tanggal.startsWith(selMonth)).sort((a:MonitoringEntry,b:MonitoringEntry)=>a.tanggal.localeCompare(b.tanggal)||a.jam.localeCompare(b.jam));
  const suhuV=me.map((e:MonitoringEntry)=>e.suhu).filter((v:number)=>v>0);
  const rhV=me.map((e:MonitoringEntry)=>e.kelembaban).filter((v:number)=>v>0);
  const avgS=suhuV.length?+(suhuV.reduce((a:number,b:number)=>a+b,0)/suhuV.length).toFixed(1):0;
  const avgR=rhV.length?+(rhV.reduce((a:number,b:number)=>a+b,0)/rhV.length).toFixed(1):0;
  const tidakS=me.filter((e:MonitoringEntry)=>!monIsOK(e.suhu,e.kelembaban,monitoringCfg)).length;
  const cmpPct=me.length?+((me.length-tidakS)/me.length*100).toFixed(1):0;
  const chartData=me.map((e:MonitoringEntry)=>({name:e.tanggal.slice(8)+"/"+(e.jam==="07:00"?"07":e.jam==="14:00"?"14":"21"),suhu:e.suhu,kelembaban:e.kelembaban}));
  const iS:React.CSSProperties={padding:"8px 12px",border:"1px solid #D1D5DB",borderRadius:8,fontSize:14,width:"100%",background:"#fff",fontFamily:"inherit",outline:"none"};
  const mLabel=MON_MONTHS_ID[+selMonth.slice(5)-1]+" "+selMonth.slice(0,4);
  return (
    <div>
      <div style={{marginBottom:14}}>
        <div style={{fontWeight:800,fontSize:18,color:MC.pri,marginBottom:2}}>🌡 Monitoring Suhu &amp; Kelembaban</div>
        <div style={{fontSize:12,color:"#64748B"}}>{monitoringCfg.lokasiRuang} · Standar Suhu {monitoringCfg.suhuMin}-{monitoringCfg.suhuMax}°C · RH {monitoringCfg.rhMin}-{monitoringCfg.rhMax}%</div>
      </div>
      <div style={{display:"flex",gap:6,marginBottom:18,flexWrap:"wrap"}}>
        {[{k:"harian",l:"📋 Harian"},{k:"grafik",l:"📊 Grafik"},{k:"unduh",l:"📥 Unduh"},{k:"standar",l:"⚙ Standar"}].map(t=>(
          <button key={t.k} onClick={()=>setSubTab(t.k)} style={{padding:"7px 18px",borderRadius:20,border:"none",background:subTab===t.k?MC.pri:"#E2E8F0",color:subTab===t.k?"#fff":"#475569",fontWeight:600,fontSize:13,cursor:"pointer",fontFamily:"inherit",transition:"all .15s"}}>{t.l}</button>
        ))}
      </div>
      {(subTab==="harian"||subTab==="grafik"||subTab==="unduh")&&(
        <div style={{display:"flex",gap:10,alignItems:"flex-end",marginBottom:18,flexWrap:"wrap"}}>
          <div><div style={{fontSize:11,fontWeight:700,color:"#64748B",marginBottom:4}}>Bulan</div><input type="month" value={selMonth} onChange={e=>setSelMonth(e.target.value)} style={{...iS,width:"auto"}}/></div>
          {subTab==="harian"&&<div><div style={{fontSize:11,fontWeight:700,color:"#64748B",marginBottom:4}}>Tanggal Input</div><input type="date" value={formDate} onChange={e=>setFormDate(e.target.value)} style={{...iS,width:"auto"}}/></div>}
        </div>
      )}
      {/* ── HARIAN ── */}
      {subTab==="harian"&&(
        <div>
          <div style={{background:"linear-gradient(135deg,#EFF6FF,#F0FDF4)",border:"1px solid #BAE6FD",borderRadius:14,padding:18,marginBottom:18}}>
            <div style={{fontWeight:700,color:MC.pri,fontSize:14,marginBottom:12}}>📝 Input Pemantauan — {fD(formDate)}</div>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                <thead><tr>{["Jam","Suhu (°C)","Kelembaban (%)","Petugas","Status"].map(h=><th key={h} style={{padding:"8px 12px",background:"#DBEAFE",color:MC.pri,fontWeight:700,textAlign:"left",borderBottom:"2px solid #BAE6FD",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
                <tbody>
                  {MON_JAMS.map(jam=>{
                    const s=slots[jam];const sv=parseFloat(s.suhu);const rv=parseFloat(s.kelembaban);
                    const hasV=!!(s.suhu&&s.kelembaban&&!isNaN(sv)&&!isNaN(rv));
                    const ok:boolean|null=hasV?monIsOK(sv,rv,monitoringCfg):null;
                    return (
                      <tr key={jam} style={{background:"#fff",borderBottom:"1px solid #E2E8F0"}}>
                        <td style={{padding:"8px 12px",fontWeight:700,color:MC.pri,whiteSpace:"nowrap"}}>{jam}</td>
                        <td style={{padding:"6px 8px"}}><input type="number" step="0.1" min="10" max="40" placeholder="mis. 21.5" value={s.suhu} onChange={e=>setSlots(p=>({...p,[jam]:{...p[jam],suhu:e.target.value}}))} style={{...iS,width:100}}/></td>
                        <td style={{padding:"6px 8px"}}><input type="number" step="1" min="0" max="100" placeholder="mis. 55" value={s.kelembaban} onChange={e=>setSlots(p=>({...p,[jam]:{...p[jam],kelembaban:e.target.value}}))} style={{...iS,width:85}}/></td>
                        <td style={{padding:"6px 8px"}}><input type="text" placeholder="Nama petugas" value={s.petugas} onChange={e=>setSlots(p=>({...p,[jam]:{...p[jam],petugas:e.target.value}}))} style={{...iS,minWidth:130}}/></td>
                        <td style={{padding:"8px 12px",whiteSpace:"nowrap"}}>
                          {ok===null?<span style={{color:"#94A3B8",fontSize:12}}>—</span>:ok?<span style={{background:MC.okBg,color:MC.ok,fontWeight:700,fontSize:11,padding:"3px 10px",borderRadius:20}}>✓ SESUAI</span>:<span style={{background:MC.errBg,color:MC.err,fontWeight:700,fontSize:11,padding:"3px 10px",borderRadius:20}}>✕ TIDAK SESUAI</span>}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            <div style={{marginTop:14}}><Btn color={MC.pri} onClick={handleSave} style={{background:MC.pri,color:"#fff",border:"none"}}>💾 Simpan Data {formDate}</Btn></div>
          </div>
          <div style={{background:"#fff",border:"1px solid #E2E8F0",borderRadius:12,overflow:"hidden"}}>
            <div style={{background:"linear-gradient(90deg,"+MC.pri+","+MC.priL+")",padding:"10px 16px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <span style={{color:"#fff",fontWeight:700,fontSize:14}}>📋 Data {mLabel}</span>
              <span style={{color:"rgba(255,255,255,.75)",fontSize:12}}>{me.length} entri</span>
            </div>
            {me.length===0?(
              <div style={{padding:32,textAlign:"center",color:"#94A3B8"}}><div style={{fontSize:32,marginBottom:8}}>🌡</div><div style={{fontWeight:600}}>Belum ada data monitoring bulan ini</div><div style={{fontSize:12,marginTop:4}}>Gunakan form di atas untuk menambahkan data</div></div>
            ):(
              <div style={{overflowX:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                  <thead style={{background:"#F8FAFC"}}><tr>{["Tanggal","Jam","Suhu","RH","Std Suhu","Std RH","Status","Petugas",""].map(h=><th key={h} style={{padding:"8px 12px",textAlign:"left",fontWeight:700,color:"#475569",fontSize:11,borderBottom:"1px solid #E2E8F0",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
                  <tbody>
                    {me.map((e:MonitoringEntry,i:number)=>{
                      const ok=monIsOK(e.suhu,e.kelembaban,monitoringCfg); const prevDate=i>0?me[i-1].tanggal:"";
                      return (
                        <tr key={e.id} style={{borderBottom:"1px solid #F1F5F9",background:i%2===0?"#fff":"#FAFBFC"}}>
                          <td style={{padding:"7px 12px",fontWeight:e.tanggal!==prevDate?700:400,color:"#374151"}}>{e.tanggal!==prevDate?e.tanggal:""}</td>
                          <td style={{padding:"7px 12px",fontWeight:700,color:MC.pri}}>{e.jam}</td>
                          <td style={{padding:"7px 12px",fontWeight:700,color:ok?MC.ok:MC.err}}>{e.suhu}°C</td>
                          <td style={{padding:"7px 12px",fontWeight:700,color:ok?MC.ok:MC.err}}>{e.kelembaban}%</td>
                          <td style={{padding:"7px 12px",color:"#64748B",fontSize:12}}>{monitoringCfg.suhuMin}-{monitoringCfg.suhuMax}</td>
                          <td style={{padding:"7px 12px",color:"#64748B",fontSize:12}}>{monitoringCfg.rhMin}-{monitoringCfg.rhMax}</td>
                          <td style={{padding:"7px 12px"}}><span style={{background:ok?MC.okBg:MC.errBg,color:ok?MC.ok:MC.err,fontSize:11,fontWeight:700,padding:"2px 8px",borderRadius:12}}>{ok?"SESUAI":"TIDAK SESUAI"}</span></td>
                          <td style={{padding:"7px 12px",color:"#374151"}}>{e.petugas}</td>
                          <td style={{padding:"7px 12px"}}><button onClick={()=>{setMonitoringEntries((p:MonitoringEntry[])=>p.filter(x=>x.id!==e.id));showToast("Data dihapus");}} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444",fontSize:16}}>✕</button></td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </div>
      )}
      {/* ── GRAFIK ── */}
      {subTab==="grafik"&&(
        <div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(120px,1fr))",gap:10,marginBottom:20}}>
            {([{l:"Rata-rata Suhu",v:avgS+"°C",ico:"🌡",bg:"#DBEAFE",c:"#1D4ED8"},{l:"Rata-rata RH",v:avgR+"%",ico:"💧",bg:"#DCFCE7",c:"#16A34A"},{l:"Total Ukur",v:String(me.length),ico:"📊",bg:"#FEF3C7",c:"#B45309"},{l:"Tidak Sesuai",v:String(tidakS),ico:"⚠️",bg:"#FEE2E2",c:"#DC2626"},{l:"Kepatuhan",v:cmpPct+"%",ico:"✅",bg:"#F0FDF4",c:"#16A34A"}] as {l:string;v:string;ico:string;bg:string;c:string}[]).map(s=>(
              <div key={s.l} style={{background:s.bg,borderRadius:12,padding:"12px 10px",textAlign:"center",border:"1px solid "+s.c+"20"}}>
                <div style={{fontSize:20,marginBottom:3}}>{s.ico}</div>
                <div style={{fontSize:17,fontWeight:800,color:s.c}}>{s.v}</div>
                <div style={{fontSize:10,color:"#64748B",marginTop:2,lineHeight:1.3}}>{s.l}</div>
              </div>
            ))}
          </div>
          {chartData.length===0?(
            <div style={{textAlign:"center",padding:48,color:"#94A3B8"}}><div style={{fontSize:40,marginBottom:10}}>📊</div><div style={{fontWeight:600}}>Belum ada data untuk bulan ini</div></div>
          ):(
            <>
              <div style={{background:"#fff",border:"1px solid #E2E8F0",borderRadius:12,padding:"16px 16px 8px",marginBottom:14}}>
                <div style={{fontWeight:700,color:MC.pri,marginBottom:10,fontSize:13}}>🌡 Grafik Suhu — {mLabel}</div>
                <ResponsiveContainer width="100%" height={200}>
                  <LineChart data={chartData} margin={{top:4,right:20,left:0,bottom:4}}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#F1F5F9"/>
                    <XAxis dataKey="name" tick={{fontSize:9}} interval={Math.max(0,Math.floor(chartData.length/10)-1)}/>
                    <YAxis domain={[15,30]} tick={{fontSize:11}} unit="°C" width={44}/>
                    <Tooltip formatter={(v:any)=>[v+"°C","Suhu"]}/>
                    <ReferenceLine y={monitoringCfg.suhuMin} stroke="#EF4444" strokeDasharray="4 3" label={{value:"Min "+monitoringCfg.suhuMin,position:"insideTopRight",fontSize:9,fill:"#EF4444"}}/>
                    <ReferenceLine y={monitoringCfg.suhuMax} stroke="#EF4444" strokeDasharray="4 3" label={{value:"Maks "+monitoringCfg.suhuMax,position:"insideBottomRight",fontSize:9,fill:"#EF4444"}}/>
                    <Line type="monotone" dataKey="suhu" stroke={MC.pri} strokeWidth={2.5} dot={{r:2,fill:MC.pri}} activeDot={{r:5}}/>
                  </LineChart>
                </ResponsiveContainer>
              </div>
              <div style={{background:"#fff",border:"1px solid #E2E8F0",borderRadius:12,padding:"16px 16px 8px"}}>
                <div style={{fontWeight:700,color:MC.ok,marginBottom:10,fontSize:13}}>💧 Grafik Kelembaban — {mLabel}</div>
                <ResponsiveContainer width="100%" height={200}>
                  <LineChart data={chartData} margin={{top:4,right:20,left:0,bottom:4}}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#F1F5F9"/>
                    <XAxis dataKey="name" tick={{fontSize:9}} interval={Math.max(0,Math.floor(chartData.length/10)-1)}/>
                    <YAxis domain={[30,90]} tick={{fontSize:11}} unit="%" width={44}/>
                    <Tooltip formatter={(v:any)=>[v+"%","Kelembaban"]}/>
                    <ReferenceLine y={monitoringCfg.rhMin} stroke="#EF4444" strokeDasharray="4 3" label={{value:"Min "+monitoringCfg.rhMin,position:"insideTopRight",fontSize:9,fill:"#EF4444"}}/>
                    <ReferenceLine y={monitoringCfg.rhMax} stroke="#EF4444" strokeDasharray="4 3" label={{value:"Maks "+monitoringCfg.rhMax,position:"insideBottomRight",fontSize:9,fill:"#EF4444"}}/>
                    <Line type="monotone" dataKey="kelembaban" stroke={MC.ok} strokeWidth={2.5} dot={{r:2,fill:MC.ok}} activeDot={{r:5}}/>
                  </LineChart>
                </ResponsiveContainer>
              </div>
            </>
          )}
        </div>
      )}
      {/* ── UNDUH ── */}
      {subTab==="unduh"&&(
        <div>
          <div style={{display:"grid",gap:12,marginBottom:16}}>
            {([{type:"harian" as const,ico:"📋",title:"Laporan Monitoring Harian",desc:"Data harian + rekap statistik bulanan",c:"#0369A1",bg:"#EFF6FF"},{type:"rekap" as const,ico:"📊",title:"Rekap Bulanan",desc:"Statistik ringkas: rata-rata suhu, RH, kepatuhan",c:"#7C3AED",bg:"#F5F3FF"},{type:"akreditasi" as const,ico:"🏅",title:"Laporan Akreditasi",desc:"Laporan lengkap dengan format akreditasi & kolom tanda tangan",c:"#B45309",bg:"#FFFBEB"}]).map(r=>(
              <div key={r.type} style={{background:r.bg,border:"1px solid "+r.c+"25",borderRadius:12,padding:"14px 16px",display:"flex",justifyContent:"space-between",alignItems:"center",gap:12,flexWrap:"wrap"}}>
                <div style={{flex:1}}><div style={{fontWeight:700,color:r.c,fontSize:14,marginBottom:3}}>{r.ico} {r.title}</div><div style={{fontSize:12,color:"#64748B"}}>{r.desc}</div></div>
                <Btn color={r.c} onClick={()=>monXlsx(monitoringEntries,monitoringCfg,selMonth,r.type,showToast)} style={{background:r.c,color:"#fff",border:"none",whiteSpace:"nowrap",flexShrink:0}}>⬇ Excel</Btn>
              </div>
            ))}
          </div>
          <div style={{background:"#F8FAFC",borderRadius:10,padding:"12px 14px",fontSize:12,color:"#64748B",border:"1px solid #E2E8F0",lineHeight:1.7,marginBottom:14}}>
            <b>📌 Penyimpanan:</b> Data monitoring disertakan dalam backup otomatis JSON bersama data operasi dan lembur. Aktifkan di tab Arsip → Setelan → Dropbox.
          </div>
          {/* ── Supabase cloud download ── */}
          <div style={{background:"#EEF2FF",border:"1.5px solid #6366F133",borderRadius:14,padding:"16px 16px 14px"}}>
            <div style={{fontWeight:700,color:"#4F46E5",fontSize:14,marginBottom:6}}>☁️ Download Monitoring dari Supabase</div>
            <div style={{fontSize:12,color:"#64748B",marginBottom:12}}>Unduh data monitoring langsung dari cloud Supabase, tanpa bergantung data perangkat ini. Data tersinkron dari semua device.</div>
            <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap",marginBottom:12}}>
              <div style={{flex:1,minWidth:160}}>
                <div style={{fontSize:12,fontWeight:700,color:"#64748B",marginBottom:4}}>Pilih Bulan</div>
                <input style={iS} type="month" value={supaMonYM} onChange={e=>setSupaMonYM(e.target.value)} disabled={supaMonBusy}/>
              </div>
            </div>
            <div style={{display:"grid",gap:8}}>
              {([{type:"harian" as const,ico:"📋",title:"Harian",c:"#0369A1"},{type:"rekap" as const,ico:"📊",title:"Rekap Bulanan",c:"#7C3AED"},{type:"akreditasi" as const,ico:"🏅",title:"Akreditasi",c:"#B45309"}]).map(r=>(
                <div key={r.type} style={{display:"flex",justifyContent:"space-between",alignItems:"center",background:"rgba(255,255,255,.7)",borderRadius:10,padding:"10px 14px",gap:10}}>
                  <div style={{fontWeight:600,color:r.c,fontSize:13}}>{r.ico} {r.title}</div>
                  <button onClick={()=>downloadMonFromSupa(r.type)} disabled={supaMonBusy} style={{background:supaMonBusy?"#9CA3AF":"#4F46E5",color:"#fff",border:"none",borderRadius:8,padding:"6px 14px",cursor:supaMonBusy?"not-allowed":"pointer",fontSize:12,fontWeight:700,fontFamily:"inherit",whiteSpace:"nowrap"}}>
                    {supaMonBusy?"⏳ ...":"☁️ Download Excel"}
                  </button>
                </div>
              ))}
            </div>
            {!supaCfg?.url && <div style={{fontSize:11,color:"#EF4444",marginTop:10}}>⚠ Supabase belum dikonfigurasi — atur di tab Arsip → Setelan.</div>}
          </div>
        </div>
      )}
      {/* ── STANDAR ── */}
      {subTab==="standar"&&(
        <div style={{maxWidth:480}}>
          <div style={{background:"#fff",border:"1px solid #E2E8F0",borderRadius:12,padding:18}}>
            <div style={{fontWeight:700,color:MC.pri,fontSize:14,marginBottom:16}}>⚙ Konfigurasi Standar Monitoring</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:14}}>
              {[{k:"suhuMin",l:"Suhu Min (°C)"},{k:"suhuMax",l:"Suhu Maks (°C)"},{k:"rhMin",l:"RH Min (%)"},{k:"rhMax",l:"RH Maks (%)"}].map(f=>(
                <div key={f.k}><div style={{fontSize:12,fontWeight:700,color:"#64748B",marginBottom:4}}>{f.l}</div><input type="number" step="0.5" value={(cfgForm as any)[f.k]} onChange={e=>setCfgForm((p:MonitoringCfg)=>({...p,[f.k]:parseFloat(e.target.value)||0}))} style={iS}/></div>
              ))}
            </div>
            <div style={{marginBottom:12}}><div style={{fontSize:12,fontWeight:700,color:"#64748B",marginBottom:4}}>Lokasi / Nama Ruang</div><input type="text" value={cfgForm.lokasiRuang} onChange={e=>setCfgForm((p:MonitoringCfg)=>({...p,lokasiRuang:e.target.value}))} style={iS} placeholder="Kamar Bedah"/></div>
            <div style={{marginBottom:16}}><div style={{fontSize:12,fontWeight:700,color:"#64748B",marginBottom:4}}>Kepala Kamar Bedah</div><input type="text" value={cfgForm.kepalaKamarBedah} onChange={e=>setCfgForm((p:MonitoringCfg)=>({...p,kepalaKamarBedah:e.target.value}))} style={iS} placeholder="Nama kepala kamar bedah"/></div>
            <div style={{background:MC.okBg,borderRadius:8,padding:"8px 12px",marginBottom:16,fontSize:12,color:MC.ok,border:"1px solid #BBF7D0"}}>Standar aktif: Suhu {monitoringCfg.suhuMin}-{monitoringCfg.suhuMax}°C · RH {monitoringCfg.rhMin}-{monitoringCfg.rhMax}%</div>
            <Btn color={MC.pri} onClick={()=>{setMonitoringCfg(cfgForm);showToast("✓ Pengaturan standar disimpan",MC.ok);}} style={{background:MC.pri,color:"#fff",border:"none"}}>💾 Simpan Pengaturan</Btn>
          </div>
        </div>
      )}
    </div>
  );
}

/* ─── MAIN APP ───────────────────────────────────────────────────────── */
export default function App() {
  const [pinOK,    setPinOK]    = useState(false);
  const [savedPin, setSavedPin] = useState(CONFIG.DEFAULT_PIN);
  const [tab,      setTab]      = useState("jadwal");
  const [ops,      setOps]      = useState<any[]>(()=>{try{return JSON.parse(localStorage.getItem("kb_ops")||"[]");}catch{return [];}});
  const [staff,    setStaff]    = useState<any[]>([]);
  const [roster,   setRoster]   = useState<any[]>([]);
  const [notifs,   setNotifs]   = useState<any[]>([]);
  const [archive,  setArchive]  = useState<any[]>([]);
  const [toast,    setToast]    = useState<any>(null);
  const [opForm,   setOpForm]   = useState({...EOP, date:todayDate()});
  const [editingOp,setEditingOp]= useState<any>(null);
  const [opErrors, setOpErrors] = useState<any>({});
  const [dupWarn,  setDupWarn]  = useState(false);
  const [lSet,     setLSet]     = useState({greeting:gWord(),recipients:"",keterangan:"",pembawaHP:"",katimPhone:"",grupPhone:""});
  const [lSby,     setLSby]     = useState({anest:[""],cyto:[""],nurses:[""],pembawaHP:""});
  const [reqOpId,  setReqOpId]  = useState<string|null>(null);
  const [reqText,  setReqText]  = useState("");
  const [privacyMode,setPM]     = useState(false);
  const [supaCfg,  setSupaCfg]  = useState<SupabaseConfig>(()=>{try{return JSON.parse(localStorage.getItem("kb_supaCfg")||"null")||defaultSupaCfg;}catch{return defaultSupaCfg;}});
  const [supaStatus,setSupaStatus] = useState<{ok:boolean;msg:string}|null>(null);
  const [supaBackingUp,setSupaBU] = useState(false);
  const [dbxCfg,   setDbxCfg]   = useState<DropboxConfig>(()=>{try{return JSON.parse(localStorage.getItem("kb_dbxCfg")||"null")||defaultDbxCfg;}catch{return defaultDbxCfg;}});
  const [dbxStatus,setDbxStatus] = useState<{ok:boolean;msg:string}|null>(null);
  const [dbxBacking,setDbxBacking] = useState(false);
  const dbxAutoRef      = useRef<ReturnType<typeof setInterval>|null>(null);
  const dbxRtRef        = useRef<ReturnType<typeof setTimeout>|null>(null);
  const dbxExcelAutoRef = useRef<ReturnType<typeof setInterval>|null>(null);
  const dualRtRef       = useRef<ReturnType<typeof setTimeout>|null>(null);
  const supaStartupRef  = useRef(false);
  const [templates,      setTemplates]      = useState<any[]>([]);
  const [auditLog,       setAuditLog]       = useState<any[]>([]);
  const [rtEnabled,      setRtEnabled]      = useState(false);
  const [rtStatus,       setRtStatus]       = useState<"offline"|"connecting"|"online">("offline");
  const [lemburPegawai,  setLemburPegawai]  = useState<any[]>(()=>{try{return JSON.parse(localStorage.getItem("kb_lemburPegawai")||"[]");}catch{return [];}});
  const [lemburData,     setLemburData]     = useState<Record<string,any>>(()=>{try{return JSON.parse(localStorage.getItem("kb_lemburData")||"{}");}catch{return {};}});
  const [monitoringEntries, setMonitoringEntries] = useState<MonitoringEntry[]>(()=>{try{return JSON.parse(localStorage.getItem("kb_monitoring")||"[]");}catch{return [];}});
  const [monitoringCfg,     setMonitoringCfg]     = useState<MonitoringCfg>(()=>{try{return JSON.parse(localStorage.getItem("kb_monCfg")||"null")||defaultMonCfg;}catch{return defaultMonCfg;}});
  const [currentTime, setCurrentTime] = useState(new Date());
  const [storageWarn, setStorageWarn] = useState<{pct:number;kb:number}|null>(null);
  const lastAct   = useRef(Date.now());
  const autoBackupRef = useRef<ReturnType<typeof setInterval>|null>(null);
  const rtChannelRef  = useRef<RealtimeChannel|null>(null);
  const rtIgnoreRef   = useRef(false);

  const logAudit = useCallback((action: string, op: any) => {
    setAuditLog(p=>[...p,{id:gId(), action, patient:op.patient||"—", detail:`${op.procedure||""} · ${op.date||""} ${op.time||""}`, time:fNow()}]);
  },[]);

  /* Live clock — update every second */
  useEffect(()=>{
    const iv = setInterval(()=>setCurrentTime(new Date()), 1000);
    return ()=>clearInterval(iv);
  },[]);

  /* localStorage storage warning */
  useEffect(()=>{
    const check = () => {
      try {
        let total = 0;
        for(const k of Object.keys(localStorage)) total += (localStorage.getItem(k)||"").length + k.length;
        const kb = Math.round(total/1024);
        const pct = Math.min(100, Math.round(total/(5*1024*1024)*100));
        setStorageWarn(pct>=70?{pct,kb}:null);
      } catch {}
    };
    check();
    const iv = setInterval(check, 30000);
    return ()=>clearInterval(iv);
  },[ops,lemburData,monitoringEntries]);

  /* Persist config + data to localStorage (cache/fallback for offline use) */
  useEffect(()=>{ try{localStorage.setItem("kb_ops",JSON.stringify(ops));}catch{} },[ops]);
  useEffect(()=>{ try{localStorage.setItem("kb_monitoring",JSON.stringify(monitoringEntries));}catch{} },[monitoringEntries]);
  useEffect(()=>{ try{localStorage.setItem("kb_monCfg",JSON.stringify(monitoringCfg));}catch{} },[monitoringCfg]);
  useEffect(()=>{ try{localStorage.setItem("kb_lemburPegawai",JSON.stringify(lemburPegawai));}catch{} },[lemburPegawai]);
  useEffect(()=>{ try{localStorage.setItem("kb_lemburData",JSON.stringify(lemburData));}catch{} },[lemburData]);
  useEffect(()=>{ try{localStorage.setItem("kb_dbxCfg",JSON.stringify(dbxCfg));}catch{} },[dbxCfg]);
  useEffect(()=>{ try{const s={url:supaCfg.url,anonKey:supaCfg.anonKey,table:supaCfg.table,autoBackup:supaCfg.autoBackup,backupInterval:supaCfg.backupInterval,realtimeBackup:supaCfg.realtimeBackup};localStorage.setItem("kb_supaCfg",JSON.stringify(s));}catch{} },[supaCfg.url,supaCfg.anonKey,supaCfg.table,supaCfg.autoBackup,supaCfg.backupInterval,supaCfg.realtimeBackup]);

  /* Load primary data from Supabase on startup (multi-device: cloud is source of truth) */
  useEffect(()=>{
    if(supaStartupRef.current) return;
    supaStartupRef.current=true;
    if(!supaCfg.url||!supaCfg.anonKey) return;
    (async()=>{
      try{
        const res=await supabaseRestore(supaCfg);
        if(res.ok&&res.data){
          if(Array.isArray(res.data.ops)&&res.data.ops.length) setOps(res.data.ops);
          if(Array.isArray(res.data.lemburPegawai)&&res.data.lemburPegawai.length) setLemburPegawai(res.data.lemburPegawai);
          if(res.data.lemburData&&Object.keys(res.data.lemburData).length) setLemburData(res.data.lemburData);
          if(Array.isArray(res.data.monitoringEntries)&&res.data.monitoringEntries.length) setMonitoringEntries(res.data.monitoringEntries);
          if(res.data.monitoringCfg) setMonitoringCfg(res.data.monitoringCfg);
          if(Array.isArray(res.data.archive)) setArchive(res.data.archive);
        }
      }catch{}
    })();
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[]);

  /* Dropbox auto-backup interval */
  useEffect(()=>{
    if(dbxAutoRef.current) clearInterval(dbxAutoRef.current);
    if(!dbxCfg.autoBackup||!dbxCfg.token) return;
    const ms=dbxCfg.backupInterval*60*1000;
    dbxAutoRef.current=setInterval(async()=>{
      const data={ops,archive,notifs,lemburPegawai,lemburData,monitoringEntries,monitoringCfg};
      const res=await dropboxUpload(dbxCfg,data);
      if(res.ok) setDbxCfg(p=>({...p,lastBackup:fNow()}));
    },ms);
    return()=>{ if(dbxAutoRef.current) clearInterval(dbxAutoRef.current); };
  },[dbxCfg.autoBackup,dbxCfg.token,dbxCfg.backupInterval]);

  /* ── COMBINED dual realtime backup: Supabase + Dropbox fired SIMULTANEOUSLY ── */
  useEffect(()=>{
    const hasSupa = supaCfg.realtimeBackup && supaCfg.url && supaCfg.anonKey;
    const hasDbx  = dbxCfg.realtimeBackup && dbxCfg.token;
    if(!hasSupa && !hasDbx) return;
    if(dualRtRef.current) clearTimeout(dualRtRef.current);
    dualRtRef.current = setTimeout(async()=>{
      const data = {exportedAt:fNow(), ops, archive, notifs, lemburPegawai, lemburData, monitoringEntries, monitoringCfg};
      const tasks: Promise<any>[] = [];
      if(hasSupa) tasks.push(supabaseBackup(supaCfg, data).then(r=>{ if(r.ok) setSupaCfg(p=>({...p,lastBackup:fNow()})); }));
      if(hasDbx)  tasks.push(dropboxUpload(dbxCfg, data).then(r=>{ if(r.ok) setDbxCfg(p=>({...p,lastBackup:fNow()})); }));
      await Promise.all(tasks);
    }, 5000);
    return()=>{ if(dualRtRef.current) clearTimeout(dualRtRef.current); };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[ops, lemburData, lemburPegawai, monitoringEntries]);

  /* ── Auto Excel backup every 24 hours ── */
  useEffect(()=>{
    if(dbxExcelAutoRef.current) clearInterval(dbxExcelAutoRef.current);
    if(!dbxCfg.autoExcelBackup||!dbxCfg.token) return;
    const run24h = async (currentOps: any[], currentLemburPegawai: any[], currentLemburData: Record<string,any>, currentCfg: DropboxConfig) => {
      const now = Date.now();
      const last = currentCfg.lastExcelBackupTs||0;
      if(now - last < 24*60*60*1000) return; /* 24h not yet elapsed */
      const folder = currentCfg.path.replace(/[^/]+$/, "");
      const stamp = new Date().toISOString().slice(0,10);
      /* — Operasi Excel — */
      const wbOps = XLSX.utils.book_new();
      const h = ["No","Tanggal","Jam","Pasien","Usia","RM","Jenis","Diagnosa","Tindakan","Kamar","Operator","Anestesi","Asisten","Instrumen","P.Anestesi","Onloop","RR/Katim","Status","Dibuat"];
      const rows = [...currentOps].sort((a:any,b:any)=>a.date.localeCompare(b.date)||a.time.localeCompare(b.time)).map((o:any,i:number)=>[i+1,o.date,o.time,o.patient||"",o.age||"",o.rm||"",o.opType||"",o.diagnosis||"",o.procedure||"",o.room||"",o.surgeon||"",o.anesthesiologist||"",o.assistantNurse||"",o.circulatingNurse||"",o.anesthesiaNurse||"",o.onloopNurse||"",o.rrKatim||"",o.status||"",o.createdAt||""]);
      const wsOps = XLSX.utils.aoa_to_sheet([h,...rows]);
      XLSX.utils.book_append_sheet(wbOps, wsOps, "Jadwal Operasi");
      const resOps = await dropboxUploadExcel(currentCfg.token, `${folder}Auto_Operasi_${stamp}.xlsx`, wbOps);
      /* — Lembur Excel — */
      const wbLmb = XLSX.utils.book_new();
      currentLemburPegawai.forEach((p:any)=>{
        const entries: any[][]=[];
        Object.keys(currentLemburData).filter(k=>k.startsWith(p.id+"_")).sort().forEach(k=>{
          const rec=currentLemburData[k];
          (rec.entries||[]).forEach((e:any)=>entries.push([p.name,p.nik||"",k.replace(p.id+"_",""),e.tanggalAwal||"",e.tanggalAkhir||"",e.jamMasuk||"",e.jamKeluar||"",e.keperluanLembur||"",e.keterangan||"",e.ttd||""]));
        });
        if(entries.length){const ws=XLSX.utils.aoa_to_sheet([["Nama","NIK","Bulan","Tgl Awal","Tgl Akhir","Jam Masuk","Jam Keluar","Keperluan","Keterangan","TTD"],...entries]);XLSX.utils.book_append_sheet(wbLmb,ws,p.name.slice(0,28));}
      });
      if(!wbLmb.SheetNames.length) XLSX.utils.book_append_sheet(wbLmb,XLSX.utils.aoa_to_sheet([["Belum ada data lembur"]]),"Data Lembur");
      const resLmb = await dropboxUploadExcel(currentCfg.token, `${folder}Auto_Lembur_${stamp}.xlsx`, wbLmb);
      /* — Monitoring Excel — */
      const wbMon = XLSX.utils.book_new();
      const monRows = monitoringEntries.map((e:MonitoringEntry)=>[e.tanggal,e.jam,e.suhu,e.kelembaban,monIsOK(e.suhu,e.kelembaban,monitoringCfg)?"SESUAI":"TIDAK SESUAI",e.petugas,monitoringCfg.lokasiRuang||""]);
      const wsMonMain = XLSX.utils.aoa_to_sheet([["Tanggal","Jam","Suhu (°C)","Kelembaban (%)","Status","Petugas","Lokasi"],...monRows]);
      wsMonMain["!cols"]=[{wch:12},{wch:8},{wch:12},{wch:14},{wch:16},{wch:22},{wch:20}];
      XLSX.utils.book_append_sheet(wbMon, wsMonMain, "Monitoring Suhu");
      const wsMonCfg = XLSX.utils.aoa_to_sheet([["Parameter","Nilai"],["Lokasi",monitoringCfg.lokasiRuang],["Suhu Min",monitoringCfg.suhuMin],["Suhu Max",monitoringCfg.suhuMax],["RH Min",monitoringCfg.rhMin],["RH Max",monitoringCfg.rhMax],["Kepala KB",monitoringCfg.kepalaKamarBedah]]);
      XLSX.utils.book_append_sheet(wbMon, wsMonCfg, "Konfigurasi");
      const resMon = await dropboxUploadExcel(currentCfg.token, `${folder}Auto_Monitoring_${stamp}.xlsx`, wbMon);
      if(resOps.ok&&resLmb.ok&&resMon.ok){
        setDbxCfg((p:DropboxConfig)=>({...p,lastExcelBackup:fNow(),lastExcelBackupTs:Date.now()}));
      }
    };
    /* Check immediately on enable, then every hour */
    run24h(ops, lemburPegawai, lemburData, dbxCfg);
    dbxExcelAutoRef.current = setInterval(()=>{
      setDbxCfg(p=>{ run24h(ops, lemburPegawai, lemburData, p); return p; });
    }, 60*60*1000);
    return()=>{ if(dbxExcelAutoRef.current) clearInterval(dbxExcelAutoRef.current); };
  },[dbxCfg.autoExcelBackup, dbxCfg.token]);

  /* Inject responsive CSS */
  useEffect(()=>{
    const id="kb-css"; if(document.getElementById(id))return;
    const s=document.createElement("style"); s.id=id;
    s.textContent=`
      *{box-sizing:border-box;}
      html,body{margin:0;padding:0;}

      /* ── Mobile default (<640px) ── */
      body{background:${C.bg};font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;}
      .kbShell{width:100%;min-height:100vh;}
      .kb-header{padding:12px 16px 0;}
      .kb-header-inner{margin-bottom:10px;}
      .kb-tabs{display:flex;overflow-x:auto;-webkit-overflow-scrolling:touch;gap:0;padding-bottom:0;}
      .kb-tab-btn{padding:9px 12px;flex-shrink:0;white-space:nowrap;font-size:13px;font-weight:600;}
      .kb-content{padding:14px 16px 32px;}
      .kb-sidebar{display:none;}
      .kb-main{display:block;}
      .kb-hospital{font-size:10px;}
      .kb-title{font-size:15px;}

      /* ── Tablet (640px–1023px) ── */
      @media(min-width:640px){
        body{background:#E0EEE9;}
        .kbShell{max-width:720px;margin:0 auto;box-shadow:0 4px 24px rgba(0,0,0,.10);}
        input,select,textarea{font-size:15px!important;}
        .kb-tab-btn{padding:10px 16px;font-size:14px;font-weight:700;}
        .kb-content{padding:18px 22px 36px;}
        .kb-title{font-size:16px;}
        .kb-hospital{font-size:11px;}
      }

      /* ── Desktop (≥1024px) ── */
      @media(min-width:1024px){
        body{
          background:linear-gradient(135deg,#002B24 0%,#004D40 50%,#00695C 100%)!important;
          min-height:100vh;
          display:flex;align-items:flex-start;justify-content:center;
          padding:28px 20px;
        }
        #root{width:100%;max-width:1200px;display:flex;justify-content:center;}
        .kbShell{
          width:100%;max-width:1160px;
          border-radius:20px;
          overflow:hidden;
          box-shadow:0 24px 80px rgba(0,0,0,.40);
          min-height:calc(100vh - 56px);
          display:flex;flex-direction:column;
          margin:0 auto;
        }
        .kb-header{padding:18px 28px 0;}
        .kb-header-inner{margin-bottom:0;}
        .kb-tabs{display:none;}
        .kb-sidebar{
          display:flex;flex-direction:column;
          width:200px;min-width:200px;
          background:linear-gradient(180deg,#003D33,#00695C);
          padding:16px 10px;gap:4px;position:relative;
          border-right:1px solid rgba(255,255,255,.10);
          min-height:calc(100vh - 56px - 60px);
        }
        .kb-sidebar-btn{
          display:flex;align-items:center;gap:10px;
          padding:11px 14px;border-radius:10px;
          font-size:14px;font-weight:600;
          color:rgba(255,255,255,.75);
          background:none;border:none;cursor:pointer;
          text-align:left;width:100%;
          font-family:inherit;transition:all .15s;
        }
        .kb-sidebar-btn:hover{background:rgba(255,255,255,.10);color:#fff;}
        .kb-sidebar-btn.active{background:rgba(255,255,255,.18);color:#fff;font-weight:800;}
        .kb-main{display:flex;flex:1;}
        .kb-content{flex:1;padding:24px 32px 40px;overflow-y:auto;}
        .kb-title{font-size:18px;}
        .kb-hospital{font-size:12px;}
        input,select,textarea{font-size:15px!important;}
      }

      /* ── Large desktop (≥1280px) ── */
      @media(min-width:1280px){
        #root{max-width:1300px;}
        .kbShell{max-width:1260px;}
        .kb-sidebar{width:220px;min-width:220px;}
        .kb-sidebar-btn{font-size:14px;padding:12px 16px;}
        .kb-content{padding:28px 40px 48px;}
      }
    `;
    document.head.appendChild(s);
    return()=>{const el=document.getElementById(id);if(el)el.remove();};
  },[]);

  /* Auto-lock after LOCK_MS idle */
  useEffect(()=>{
    if(!pinOK)return;
    const rst=()=>{lastAct.current=Date.now();};
    const evts=["click","touchstart","keydown","mousemove","scroll"];
    evts.forEach(e=>window.addEventListener(e,rst,{passive:true}));
    const iv=setInterval(()=>{if(Date.now()-lastAct.current>LOCK_MS)setPinOK(false);},10000);
    return()=>{evts.forEach(e=>window.removeEventListener(e,rst));clearInterval(iv);};
  },[pinOK]);

  /* Supabase Realtime sync */
  useEffect(()=>{
    if(rtChannelRef.current){
      rtChannelRef.current.unsubscribe();
      rtChannelRef.current=null;
      setRtStatus("offline");
    }
    if(!rtEnabled||!supaCfg.url||!supaCfg.anonKey) return;
    setRtStatus("connecting");
    try {
      const client = createClient(supaCfg.url, supaCfg.anonKey);
      const ch = client.channel("kamar-bedah-sync")
        .on("broadcast",{event:"ops-update"},(payload:any)=>{
          if(rtIgnoreRef.current) return;
          if(payload.payload?.ops) setOps(payload.payload.ops);
        })
        .subscribe((status:string)=>{
          setRtStatus(status==="SUBSCRIBED"?"online":"offline");
        });
      rtChannelRef.current = ch;
    } catch(e){ setRtStatus("offline"); }
    return()=>{
      if(rtChannelRef.current){rtChannelRef.current.unsubscribe();rtChannelRef.current=null;}
      setRtStatus("offline");
    };
  },[rtEnabled,supaCfg.url,supaCfg.anonKey]);

  /* Broadcast ops changes when realtime is on */
  const broadcastOps = useCallback((newOps: any[])=>{
    if(!rtChannelRef.current||rtStatus!=="online") return;
    rtIgnoreRef.current=true;
    rtChannelRef.current.send({type:"broadcast",event:"ops-update",payload:{ops:newOps}}).then(()=>{ setTimeout(()=>{rtIgnoreRef.current=false;},300); });
  },[rtStatus]);

  /* Auto-backup to Supabase */
  useEffect(()=>{
    if(autoBackupRef.current) clearInterval(autoBackupRef.current);
    if(!supaCfg.autoBackup || !supaCfg.url || !supaCfg.anonKey) return;
    const ms = supaCfg.backupInterval * 60 * 1000;
    autoBackupRef.current = setInterval(async ()=>{
      const data = {exportedAt:fNow(), ops, archive, notifs, lemburPegawai, lemburData, monitoringEntries, monitoringCfg};
      const res = await supabaseBackup(supaCfg, data);
      if(res.ok) setSupaCfg(p=>({...p,lastBackup:fNow()}));
    }, ms);
    return ()=>{ if(autoBackupRef.current) clearInterval(autoBackupRef.current); };
  }, [supaCfg.autoBackup, supaCfg.url, supaCfg.anonKey, supaCfg.backupInterval, ops, archive, notifs, lemburPegawai, lemburData, monitoringEntries]);

  /* Supabase realtime backup is now handled by the combined dual-backup effect above */

  const showToast = useCallback((msg: string,color?: string)=>{
    setToast({msg,color:color||C.s}); setTimeout(()=>setToast(null),3200);
  },[]);

  const handlePinVerify = (newPin?: string) => {
    if(newPin){setSavedPin(newPin);showToast("✓ PIN berhasil diubah",C.s);}
    lastAct.current=Date.now(); setPinOK(true);
  };

  const handleSupaBackup = async () => {
    setSupaBU(true); setSupaStatus(null);
    const data = {exportedAt:fNow(), ops, archive, notifs, lemburPegawai, lemburData, monitoringEntries, monitoringCfg};
    const res = await supabaseBackup(supaCfg, data);
    setSupaStatus(res);
    if(res.ok) { setSupaCfg(p=>({...p,lastBackup:fNow()})); showToast("✓ Backup Supabase berhasil","#3ECF8E"); }
    else { showToast(res.msg,C.d); }
    setSupaBU(false);
  };

  const handleSupaRestoreOps = async () => {
    if(!window.confirm("Pulihkan Jadwal Operasi dari Supabase?\nData jadwal operasi lokal akan ditimpa dengan data dari Supabase.")) return;
    setSupaBU(true); setSupaStatus(null);
    const res = await supabaseRestore(supaCfg);
    setSupaStatus({ok:res.ok, msg:res.msg});
    if(res.ok && res.data) {
      if(res.data.ops) { setOps(res.data.ops); showToast("✓ Jadwal operasi dipulihkan dari Supabase","#3ECF8E"); }
      else showToast("⚠ Data operasi tidak ditemukan di backup","#E65100");
    } else { showToast(res.msg, C.d); }
    setSupaBU(false);
  };

  const handleSupaRestoreLembur = async () => {
    if(!window.confirm("Pulihkan Data Lembur dari Supabase?\nData lembur lokal akan ditimpa dengan data dari Supabase.")) return;
    setSupaBU(true); setSupaStatus(null);
    const res = await supabaseRestore(supaCfg);
    setSupaStatus({ok:res.ok, msg:res.msg});
    if(res.ok && res.data) {
      if(res.data.lemburPegawai) setLemburPegawai(res.data.lemburPegawai);
      if(res.data.lemburData) setLemburData(res.data.lemburData);
      showToast("✓ Data lembur dipulihkan dari Supabase","#3ECF8E");
    } else { showToast(res.msg, C.d); }
    setSupaBU(false);
  };

  const handleSupaRestoreMonitoring = async () => {
    if(!window.confirm("Pulihkan Data Monitoring dari Supabase?\nData monitoring lokal akan ditimpa dengan data dari Supabase.")) return;
    setSupaBU(true); setSupaStatus(null);
    const res = await supabaseRestore(supaCfg);
    setSupaStatus({ok:res.ok, msg:res.msg});
    if(res.ok && res.data) {
      if(res.data.monitoringEntries) setMonitoringEntries(res.data.monitoringEntries);
      if(res.data.monitoringCfg) setMonitoringCfg(res.data.monitoringCfg);
      showToast("✓ Data monitoring dipulihkan dari Supabase","#3ECF8E");
    } else { showToast(res.msg, C.d); }
    setSupaBU(false);
  };

  const handleSupaRestoreAll = async () => {
    if(!window.confirm("Pulihkan SEMUA data dari Supabase?\nOperasi, Lembur, dan Monitoring lokal akan ditimpa dengan data dari Supabase.")) return;
    setSupaBU(true); setSupaStatus(null);
    const res = await supabaseRestore(supaCfg);
    setSupaStatus({ok:res.ok, msg:res.msg});
    if(res.ok && res.data) {
      if(res.data.ops) setOps(res.data.ops);
      if(res.data.archive) setArchive(res.data.archive);
      if(res.data.notifs) setNotifs(res.data.notifs);
      if(res.data.lemburPegawai) setLemburPegawai(res.data.lemburPegawai);
      if(res.data.lemburData) setLemburData(res.data.lemburData);
      if(res.data.monitoringEntries) setMonitoringEntries(res.data.monitoringEntries);
      if(res.data.monitoringCfg) setMonitoringCfg(res.data.monitoringCfg);
      showToast("✓ Semua data berhasil dipulihkan dari Supabase","#3ECF8E");
    } else { showToast(res.msg, C.d); }
    setSupaBU(false);
  };

  const handleDropboxBackupMonitoringXls = async () => {
    if(!dbxCfg.token){showToast("Access Token Dropbox belum diisi",C.d);return;}
    setDbxBacking(true); setDbxStatus(null);
    const wb = XLSX.utils.book_new();
    const monRows = monitoringEntries.map((e:MonitoringEntry)=>[e.tanggal,e.jam,e.suhu,e.kelembaban,monIsOK(e.suhu,e.kelembaban,monitoringCfg)?"SESUAI":"TIDAK SESUAI",e.petugas,monitoringCfg.lokasiRuang||""]);
    const ws = XLSX.utils.aoa_to_sheet([["Tanggal","Jam","Suhu (°C)","Kelembaban (%)","Status","Petugas","Lokasi"],...monRows]);
    ws["!cols"]=[{wch:12},{wch:8},{wch:12},{wch:14},{wch:16},{wch:22},{wch:20}];
    XLSX.utils.book_append_sheet(wb, ws, "Monitoring Suhu");
    if(!wb.SheetNames.length){
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["Belum ada data monitoring"]]), "Monitoring");
    }
    const folder = dbxCfg.path.replace(/[^/]+$/, "");
    const path = `${folder}Monitoring_${todayDate()}.xlsx`;
    const res = await dropboxUploadExcel(dbxCfg.token, path, wb);
    setDbxStatus(res);
    if(res.ok){ setDbxCfg((p:DropboxConfig)=>({...p,lastBackup:fNow()})); showToast("✓ Excel Monitoring berhasil diupload ke Dropbox","#0369A1"); }
    else showToast(res.msg, C.d);
    setDbxBacking(false);
  };

  const handleDropboxBackup = async () => {
    setDbxBacking(true); setDbxStatus(null);
    const data = {ops, archive, notifs, lemburPegawai, lemburData, monitoringEntries, monitoringCfg};
    const res = await dropboxUpload(dbxCfg, data);
    setDbxStatus(res);
    if(res.ok) { setDbxCfg(p=>({...p,lastBackup:fNow()})); showToast("✓ Backup Dropbox berhasil","#0061FF"); }
    else { showToast(res.msg, C.d); }
    setDbxBacking(false);
  };

  const handleDropboxBackupOpsXls = async () => {
    if(!dbxCfg.token){showToast("Access Token Dropbox belum diisi",C.d);return;}
    setDbxBacking(true); setDbxStatus(null);
    const wb = XLSX.utils.book_new();
    const h = ["No","Tanggal","Jam","Pasien","Usia","RM","Jenis","Diagnosa","Tindakan","Kamar","Operator","Anestesi","Asisten","Instrumen","P.Anestesi","Onloop","RR/Katim","Status","Dibuat"];
    const rows = [...ops].sort((a:any,b:any)=>a.date.localeCompare(b.date)||a.time.localeCompare(b.time)).map((o:any,i:number)=>[i+1,o.date,o.time,o.patient||"",o.age||"",o.rm||"",o.opType||"",o.diagnosis||"",o.procedure||"",o.room||"",o.surgeon||"",o.anesthesiologist||"",o.assistantNurse||"",o.circulatingNurse||"",o.anesthesiaNurse||"",o.onloopNurse||"",o.rrKatim||"",o.status||"",o.createdAt||""]);
    const ws = XLSX.utils.aoa_to_sheet([h,...rows]);
    ws["!cols"]=[{wch:4},{wch:12},{wch:7},{wch:22},{wch:5},{wch:10},{wch:9},{wch:26},{wch:26},{wch:14},{wch:22},{wch:22},{wch:20},{wch:20},{wch:16},{wch:16},{wch:14},{wch:12},{wch:20}];
    XLSX.utils.book_append_sheet(wb, ws, "Jadwal Operasi");
    const folder = dbxCfg.path.replace(/[^/]+$/, "");
    const path = `${folder}Operasi_${todayDate()}.xlsx`;
    const res = await dropboxUploadExcel(dbxCfg.token, path, wb);
    setDbxStatus(res);
    if(res.ok){ setDbxCfg((p:DropboxConfig)=>({...p,lastBackup:fNow()})); showToast("✓ Excel Operasi berhasil diupload ke Dropbox","#3730A3"); }
    else showToast(res.msg, C.d);
    setDbxBacking(false);
  };

  const handleDropboxBackupLemburXls = async () => {
    if(!dbxCfg.token){showToast("Access Token Dropbox belum diisi",C.d);return;}
    setDbxBacking(true); setDbxStatus(null);
    const wb = XLSX.utils.book_new();
    /* Per-employee sheets */
    lemburPegawai.forEach((p:any)=>{
      const pegEntries: any[][] = [];
      Object.keys(lemburData).filter(k=>k.startsWith(p.id+"_")).sort().forEach(k=>{
        const rec=lemburData[k];
        (rec.entries||[]).forEach((e:any)=>pegEntries.push([p.name,p.nik||"",k.replace(p.id+"_",""),e.tanggalAwal||"",e.tanggalAkhir||"",e.jamMasuk||"",e.jamKeluar||"",e.keperluanLembur||"",e.keterangan||"",e.ttd||""]));
      });
      if(pegEntries.length){
        const ws=XLSX.utils.aoa_to_sheet([["Nama","NIK","Bulan","Tgl Awal","Tgl Akhir","Jam Masuk","Jam Keluar","Keperluan","Keterangan","TTD"],...pegEntries]);
        ws["!cols"]=[{wch:22},{wch:14},{wch:10},{wch:14},{wch:14},{wch:10},{wch:10},{wch:28},{wch:20},{wch:14}];
        XLSX.utils.book_append_sheet(wb, ws, p.name.slice(0,28));
      }
    });
    if(!wb.SheetNames.length){
      /* Empty — add placeholder */
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["Belum ada data lembur"]]), "Data Lembur");
    }
    const folder = dbxCfg.path.replace(/[^/]+$/, "");
    const path = `${folder}Lembur_${todayDate()}.xlsx`;
    const res = await dropboxUploadExcel(dbxCfg.token, path, wb);
    setDbxStatus(res);
    if(res.ok){ setDbxCfg((p:DropboxConfig)=>({...p,lastBackup:fNow()})); showToast("✓ Excel Lembur berhasil diupload ke Dropbox","#3730A3"); }
    else showToast(res.msg, C.d);
    setDbxBacking(false);
  };

  const handleDropboxRestoreOps = async () => {
    if(!window.confirm("Pulihkan Jadwal Operasi dari Dropbox?\nData jadwal operasi lokal akan ditimpa dengan data dari Dropbox.")) return;
    setDbxBacking(true); setDbxStatus(null);
    const res = await dropboxDownload(dbxCfg);
    setDbxStatus({ok:res.ok, msg:res.msg});
    if(res.ok && res.data) {
      if(res.data.ops) setOps(res.data.ops);
      setDbxCfg(p=>({...p,lastBackup:fNow()}));
      showToast("✓ Jadwal operasi berhasil dipulihkan dari Dropbox","#0061FF");
    } else { showToast(res.msg, C.d); }
    setDbxBacking(false);
  };

  const handleDropboxRestoreLembur = async () => {
    if(!window.confirm("Pulihkan Data Lembur dari Dropbox?\nData lembur lokal akan ditimpa dengan data dari Dropbox.")) return;
    setDbxBacking(true); setDbxStatus(null);
    const res = await dropboxDownload(dbxCfg);
    setDbxStatus({ok:res.ok, msg:res.msg});
    if(res.ok && res.data) {
      if(res.data.lemburPegawai) setLemburPegawai(res.data.lemburPegawai);
      if(res.data.lemburData) setLemburData(res.data.lemburData);
      setDbxCfg(p=>({...p,lastBackup:fNow()}));
      showToast("✓ Data lembur berhasil dipulihkan dari Dropbox","#0061FF");
    } else { showToast(res.msg, C.d); }
    setDbxBacking(false);
  };

  const saveOp = () => {
    const e: any={};
    if(!opForm.patient||!opForm.patient.trim()) e.patient="Nama pasien wajib";
    if(!opForm.diagnosis||!opForm.diagnosis.trim()) e.diagnosis="Diagnosa wajib";
    if(!opForm.procedure||!opForm.procedure.trim()) e.procedure="Nama tindakan wajib";
    if(!opForm.date) e.date="Tanggal wajib";
    if(!opForm.time) e.time="Jam wajib";
    if(!opForm.surgeon) e.surgeon="Dokter operator wajib";
    setOpErrors(e);
    if(Object.keys(e).length){showToast("Lengkapi field yang wajib ✱",C.d);return;}
    const clean={...opForm,patient:sanitize(opForm.patient||""),diagnosis:sanitize(opForm.diagnosis||""),procedure:sanitize(opForm.procedure||""),ruangAsal:sanitize(opForm.ruangAsal||""),allergy:sanitize(opForm.allergy||"Tidak Ada"),specialNeeds:sanitize(opForm.specialNeeds||""),assistantNurse:sanitize(opForm.assistantNurse||""),rrKatim:sanitize(opForm.rrKatim||"")};
    const dup=ops.some(o=>o.id!==editingOp?.id&&o.patient&&o.patient.toLowerCase().trim()===clean.patient.toLowerCase()&&o.date===clean.date&&o.time===clean.time);
    if(dup){setDupWarn(true);showToast("⚠ Duplikasi: pasien, tanggal & jam sudah ada",C.d);return;}
    setDupWarn(false);
    if(editingOp){
      const updated=ops.map(o=>o.id===editingOp.id?{...editingOp,...clean}:o);
      setOps(updated); logAudit("Edit",clean); broadcastOps(updated);
      showToast("✓ Jadwal berhasil diperbarui",C.s);
    } else {
      const newOp={...clean,id:gId(),status:"scheduled",reminders:[],requests:[],createdAt:fNow()};
      const updated=[...ops,newOp];
      setOps(updated); logAudit("Tambah",clean); broadcastOps(updated);
      showToast("✓ Jadwal berhasil disimpan",C.s);
    }
    setOpForm({...EOP,date:todayDate()}); setEditingOp(null); setOpErrors({}); setDupWarn(false);
    setTab("jadwal");
  };
  const resetOp    = () => {setOpForm({...EOP,date:todayDate()});setEditingOp(null);setOpErrors({});setDupWarn(false);};
  const startEditOp= (op:any)  => {setOpForm({...op});setEditingOp(op);setOpErrors({});setDupWarn(false);setTab("daftar");};
  const deleteOp   = (id:string)  => {
    const op=ops.find(o=>o.id===id);
    const updated=ops.filter(o=>o.id!==id);
    setOps(updated); if(op) logAudit("Hapus",op); broadcastOps(updated);
    showToast("Jadwal dihapus");
  };
  const sendReminder=(op:any,type:string)=>{
    if((op.reminders||[]).includes(type))return;
    setNotifs(p=>[{id:gId(),type,label:(type==="H-1"?"H-1":"1 Jam")+" → "+op.surgeon,patient:op.patient,procedure:op.procedure,message:"",sentAt:fNow()},...p]);
    setOps(p=>p.map(o=>o.id===op.id?{...o,reminders:[...(o.reminders||[]),type]}:o));
    showToast("Pengingat dicatat di Log");
  };
  const getPhone = (name: string) => { const s=staff.find(x=>x.name===name); return s&&s.phone||null; };
  const addReq   = (opId: string) => {
    if(!reqText.trim())return;
    setOps(p=>p.map(o=>o.id===opId?{...o,requests:[...(o.requests||[]),{id:gId(),text:sanitize(reqText.trim()),time:fNow()}]}:o));
    setReqText(""); setReqOpId(null); showToast("Permintaan disimpan");
  };
  const delReq = (opId: string,rId: string) => setOps(p=>p.map(o=>o.id===opId?{...o,requests:(o.requests||[]).filter((r:any)=>r.id!==rId)}:o));

  if(!pinOK) return <PinScreen onVerify={handlePinVerify} savedPin={savedPin}/>;

  const TABS = [
    {k:"jadwal",l:"📋 Jadwal"},
    {k:"daftar",l:"📝 Daftar"},
    {k:"laporan",l:"📄 Laporan"},
    {k:"wa",l:"💬 Kirim WA"},
    {k:"statistik",l:"📊 Statistik"},
    {k:"staf",l:"👥 Staf"},
    {k:"lembur",l:"⏰ Lembur"},
    {k:"monitoring",l:"🌡 Monitoring"},
    {k:"arsip",l:"🗂 Arsip"},
  ];

  const content = (
    <ErrorBoundary>
      {tab==="jadwal"     && <ViewJadwal ops={ops} setOps={setOps} startEditOp={startEditOp} deleteOp={deleteOp} sendReminder={sendReminder} reqOpId={reqOpId} setReqOpId={setReqOpId} reqText={reqText} setReqText={setReqText} addReq={addReq} delReq={delReq} getPhone={getPhone} setNotifs={setNotifs} showToast={showToast} privacyMode={privacyMode}/>}
      {tab==="daftar"     && <ViewDaftar opForm={opForm} setOpForm={setOpForm} editingOp={editingOp} resetOp={resetOp} saveOp={saveOp} opErrors={opErrors} setOpErrors={setOpErrors} staff={staff} setTab={setTab} dupWarning={dupWarn} templates={templates} setTemplates={setTemplates} showToast={showToast}/>}
      {tab==="laporan"    && <ViewLaporan ops={ops} lSet={lSet} setLSet={setLSet} lSby={lSby} setLSby={setLSby} staff={staff} roster={roster} showToast={showToast}/>}
      {tab==="wa"         && <ViewKirimWA ops={ops} staff={staff} setNotifs={setNotifs} showToast={showToast}/>}
      {tab==="statistik"  && <ViewStatistik ops={ops} archive={archive}/>}
      {tab==="staf"       && <ViewStaf staff={staff} setStaff={setStaff} roster={roster} setRoster={setRoster} showToast={showToast}/>}
      {tab==="lembur"     && <ViewLembur lemburPegawai={lemburPegawai} setLemburPegawai={setLemburPegawai} lemburData={lemburData} setLemburData={setLemburData} showToast={showToast} supaCfg={supaCfg}/>}
      {tab==="monitoring" && <ViewMonitoring monitoringEntries={monitoringEntries} setMonitoringEntries={setMonitoringEntries} monitoringCfg={monitoringCfg} setMonitoringCfg={setMonitoringCfg} showToast={showToast} supaCfg={supaCfg}/>}
      {tab==="arsip"      && <ViewArsip ops={ops} setOps={setOps} notifs={notifs} archive={archive} showToast={showToast} privacyMode={privacyMode} setPrivacyMode={setPM} supaCfg={supaCfg} setSupaCfg={setSupaCfg} onSupaBackup={handleSupaBackup} onSupaRestoreOps={handleSupaRestoreOps} onSupaRestoreLembur={handleSupaRestoreLembur} onSupaRestoreMonitoring={handleSupaRestoreMonitoring} onSupaRestoreAll={handleSupaRestoreAll} supaStatus={supaStatus} supaBackingUp={supaBackingUp} auditLog={auditLog} rtStatus={rtStatus} rtEnabled={rtEnabled} setRtEnabled={setRtEnabled} dbxCfg={dbxCfg} setDbxCfg={setDbxCfg} onDbxBackup={handleDropboxBackup} onDbxRestoreOps={handleDropboxRestoreOps} onDbxRestoreLembur={handleDropboxRestoreLembur} dbxStatus={dbxStatus} dbxBacking={dbxBacking} onDbxBackupOpsXls={handleDropboxBackupOpsXls} onDbxBackupLemburXls={handleDropboxBackupLemburXls} onDbxBackupMonitoringXls={handleDropboxBackupMonitoringXls} lemburData={lemburData} lemburPegawai={lemburPegawai} monitoringEntries={monitoringEntries} monitoringCfg={monitoringCfg}/>}
    </ErrorBoundary>
  );

  return (
    <div className="kbShell" style={{color:C.t,fontFamily:"-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif"}}>

      {/* ── Sticky header (mobile + tablet: full header; desktop: top bar only) ── */}
      <div className="kb-header" style={{background:`linear-gradient(135deg,#003D33,${C.p})`,position:"sticky",top:0,zIndex:100,boxShadow:"0 2px 12px rgba(0,0,0,.18)"}}>
        <div className="kb-header-inner" style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <img src="/logo.jpeg" alt="Logo KB" style={{width:44,height:44,borderRadius:"50%",objectFit:"cover",border:"2px solid rgba(255,255,255,.4)",animation:"lgFloat 3s ease-in-out infinite",flexShrink:0}}/>
            <div>
              <div className="kb-title" style={{fontWeight:800,color:"#fff",letterSpacing:.3}}>Kamar Bedah</div>
              <div className="kb-hospital" style={{color:"rgba(255,255,255,.65)",marginTop:2}}>{HOSPITAL}</div>
              <div style={{color:"rgba(255,255,255,.55)",fontSize:10,marginTop:2,letterSpacing:.2}}>{currentTime.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"})} · {currentTime.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"})}</div>
            </div>
          </div>
          <div style={{display:"flex",gap:8,alignItems:"center"}}>
            {rtStatus==="online" && <span style={{background:"#1B5E20",color:"#A5D6A7",fontSize:10,fontWeight:700,padding:"3px 9px",borderRadius:20,border:"1px solid #2E7D32"}}>📡 Sync</span>}
            {ops.filter((o:any)=>o.status==="ongoing").length>0 && (
              <span style={{background:"#F44336",color:"#fff",fontSize:11,fontWeight:700,padding:"3px 9px",borderRadius:20}}>
                ● {ops.filter((o:any)=>o.status==="ongoing").length} berlangsung
              </span>
            )}
            <button onClick={()=>setPinOK(false)} style={{background:"rgba(255,255,255,.15)",border:"none",borderRadius:8,padding:"6px 12px",color:"#fff",fontSize:11,cursor:"pointer",fontFamily:"inherit",fontWeight:700}}>🔒 Kunci</button>
          </div>
        </div>
        {/* Mobile/tablet tabs — hidden on desktop via CSS */}
        <div className="kb-tabs">
          {TABS.map(t=>(
            <button key={t.k} onClick={()=>setTab(t.k)} className="kb-tab-btn"
              style={{border:"none",background:tab===t.k?"rgba(255,255,255,.18)":"none",color:tab===t.k?"#fff":"rgba(255,255,255,.6)",fontWeight:tab===t.k?700:400,cursor:"pointer",borderBottom:tab===t.k?"3px solid #fff":"3px solid transparent",borderRadius:"8px 8px 0 0",transition:"all .15s",fontFamily:"inherit"}}>
              {t.l}
            </button>
          ))}
        </div>
      </div>

      {/* ── Storage warning banner ── */}
      {storageWarn && (
        <div style={{background:storageWarn.pct>=90?"#B91C1C":"#B45309",color:"#fff",padding:"8px 16px",fontSize:12,fontWeight:700,textAlign:"center",display:"flex",alignItems:"center",justifyContent:"center",gap:10,flexWrap:"wrap",zIndex:99}}>
          <span>⚠️ Penyimpanan browser {storageWarn.pct}% penuh ({storageWarn.kb} KB / 5 MB).{storageWarn.pct>=90?" Segera backup ke Supabase atau Dropbox!":" Disarankan backup ke cloud agar data tidak hilang."}</span>
          <button onClick={()=>setTab("arsip")} style={{background:"rgba(255,255,255,.25)",border:"1px solid rgba(255,255,255,.5)",borderRadius:6,color:"#fff",padding:"3px 12px",cursor:"pointer",fontSize:11,fontWeight:700,fontFamily:"inherit"}}>→ Buka Arsip & Backup</button>
          <button onClick={()=>setStorageWarn(null)} style={{background:"rgba(0,0,0,.2)",border:"none",borderRadius:6,color:"rgba(255,255,255,.8)",padding:"3px 8px",cursor:"pointer",fontSize:11,fontFamily:"inherit"}}>✕</button>
        </div>
      )}

      {/* ── Body: sidebar (desktop only) + content ── */}
      <div className="kb-main" style={{background:C.bg,flex:1}}>

        {/* Desktop sidebar — hidden on mobile via CSS */}
        <aside className="kb-sidebar">
          <div style={{fontSize:11,fontWeight:700,color:"rgba(255,255,255,.4)",letterSpacing:1,textTransform:"uppercase",padding:"4px 14px 8px",marginBottom:4}}>Menu</div>
          {TABS.filter(t=>t.k!=="arsip").map(t=>(
            <button key={t.k} onClick={()=>setTab(t.k)}
              className={`kb-sidebar-btn${tab===t.k?" active":""}`}>
              <span style={{fontSize:16,flexShrink:0}}>{t.l.split(" ")[0]}</span>
              <span>{t.l.split(" ").slice(1).join(" ")}</span>
              {t.k==="jadwal" && ops.filter((o:any)=>o.status==="ongoing").length>0 && (
                <span style={{marginLeft:"auto",background:"#F44336",color:"#fff",fontSize:9,fontWeight:700,padding:"2px 6px",borderRadius:10}}>{ops.filter((o:any)=>o.status==="ongoing").length}</span>
              )}
            </button>
          ))}
          <div style={{flex:1}}/>
          {TABS.filter(t=>t.k==="arsip").map(t=>(
            <button key={t.k} onClick={()=>setTab(t.k)}
              className={`kb-sidebar-btn${tab===t.k?" active":""}`}>
              <span style={{fontSize:16,flexShrink:0}}>{t.l.split(" ")[0]}</span>
              <span>{t.l.split(" ").slice(1).join(" ")}</span>
              {notifs.length>0 && <span style={{marginLeft:"auto",background:"rgba(255,255,255,.25)",color:"#fff",fontSize:9,fontWeight:700,padding:"2px 6px",borderRadius:10}}>{notifs.length}</span>}
            </button>
          ))}
          {/* Info footer in sidebar */}
          <div style={{padding:"12px 14px 4px",borderTop:"1px solid rgba(255,255,255,.1)"}}>
            <div style={{fontSize:10,color:"rgba(255,255,255,.35)",lineHeight:1.8}}>
              {ops.length} jadwal · {staff.length} staf<br/>
              {privacyMode && <span style={{color:"#CE93D8"}}>🔒 Mode Privasi</span>}
            </div>
          </div>
        </aside>

        {/* Content area */}
        <div className="kb-content">
          {content}
        </div>
      </div>

      {/* Toast */}
      {toast && (
        <div style={{position:"fixed",bottom:24,left:"50%",transform:"translateX(-50%)",background:toast.color,color:"#fff",padding:"11px 22px",borderRadius:14,fontSize:13,fontWeight:700,zIndex:9999,boxShadow:"0 4px 20px rgba(0,0,0,.2)",whiteSpace:"nowrap",maxWidth:"90vw",textAlign:"center"}}>
          {toast.msg}
        </div>
      )}
    </div>
  );
}
