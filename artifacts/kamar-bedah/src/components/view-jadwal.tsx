/* src/components/view-jadwal.tsx
   Tab "Jadwal" — daftar operasi aktif dengan filter status, pencarian,
   alert CYTO/emergency, riwayat (jadwal lewat), dan aksi cepat (mulai/
   selesai/batal/hapus/permintaan operator). Dipindah dari App.tsx
   (langkah 6 refactor — pisah View/layar besar, batch 7), TIDAK ADA
   perubahan logika, termasuk seluruh FIX AUDIT & FIX BUG di dalamnya. */

import { useState } from "react";
import type { ViewJadwalProps } from "../types";
import { C, OT, STS, iS, PAGE_SIZE, HOSPITAL } from "../constants";
import { todayDate, tmrwDate, fDMY, fNow, sanitize, maskName, maskRM, normalizePhone62, gId, msgCyto } from "../utils";
import { AmbilLogMedis } from "../services/supabase";
import { Card, Btn, WaBt, Bdg, SkeletonList } from "./primitives";
import { ActionIcon } from "./icons";
import { Marquee } from "./effects";

export function ViewJadwal({ops,setOps,startEditOp,deleteOp,sendReminder,reqOpId,setReqOpId,reqText,setReqText,addReq,delReq,getPhone,setNotifs,showToast,privacyMode,role,upsertOneToSupa,initialLoading}: ViewJadwalProps) {
  const [cId,setCId]     = useState<string|null>(null);
  const [cR,setCR]       = useState("");
  const [cytoM,setCytoM] = useState<any>(null);
  const [delC,setDelC]   = useState<string|null>(null);
  const [page,setPage]   = useState(0);
  const [fSt,setFSt]     = useState("all");
  const [q,setQ]         = useState("");
  const [showRiwayat,setShowRiwayat] = useState(false);

  const todayStr = todayDate();
  const activeOpsBase = ops.filter((o:any)=>o.date>=todayStr);
  const pastOps = ops.filter((o:any)=>o.date<todayStr).sort((a:any,b:any)=>b.date.localeCompare(a.date)||b.time.localeCompare(a.time));

  const sorted = [...activeOpsBase]
    .filter((op:any) => {
      if(fSt!=="all"&&op.status!==fSt) return false;
      if(q.trim()){ const lq=q.toLowerCase(); return (op.patient||"").toLowerCase().includes(lq)||(op.procedure||"").toLowerCase().includes(lq)||(op.surgeon||"").toLowerCase().includes(lq); }
      return true;
    })
    .sort((a:any,b:any)=>{
      const at=a.date===todayStr, bt=b.date===todayStr;
      if(at&&a.opType==="cyto"&&!(bt&&b.opType==="cyto")) return -1;
      if(bt&&b.opType==="cyto"&&!(at&&a.opType==="cyto")) return 1;
      const aDone = a.status==="done"||a.status==="batal";
      const bDone = b.status==="done"||b.status==="batal";
      if(!aDone && bDone) return -1;
      if(aDone && !bDone) return 1;
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
    window.open(`https://wa.me/${normalizePhone62(cytoM.ph)}?text=${encodeURIComponent(cytoM.msg)}`,"_blank");
    setNotifs((p:any)=>[{id:gId(),type:"cyto",label:"⚡ WA Cyto → "+cytoM.op.surgeon,patient:cytoM.op.patient,procedure:cytoM.op.procedure,message:cytoM.msg,sentAt:cytoM.ts},...p].slice(0,200));
    showToast("✓ WhatsApp dibuka — tekan Kirim di WA","#25D366");
    setCytoM(null);
  };

  return (
    <div>
      {privacyMode && <div style={{background:"#4A148C",borderRadius:10,padding:"8px 14px",marginBottom:8,fontSize:12,fontWeight:700,color:"#fff"}}>🔒 Mode Privasi — Nama & RM disamarkan</div>}
      {cytoActive.length>0 && (
        <div style={{background:"#FFCDD2",borderRadius:10,padding:"10px 14px",marginBottom:12,border:"2px solid #F44336"}}>
          <div style={{fontSize:13,fontWeight:800,color:"#B71C1C",marginBottom:4}}>⚠️ OPERASI CYTO / EMERGENCY AKTIF</div>
          {cytoActive.length>1 && (
            <div style={{marginBottom:6,color:"#B71C1C",fontSize:11,fontWeight:700}}>
              <Marquee speed={cytoActive.length*4}>
                {cytoActive.map((o:any)=>`🔴 ${maskName(o.patient,privacyMode)} (${o.time} WIB)`).join("　★　")}
              </Marquee>
            </div>
          )}
          {cytoActive.map((o:any)=><div key={o.id} style={{fontSize:12,color:"#B71C1C"}}>• {maskName(o.patient,privacyMode)} — {o.procedure} ({o.time} WIB)</div>)}
        </div>
      )}
      <div style={{background:"#fff",borderRadius:16,padding:"14px 16px",marginBottom:14,
        boxShadow:"6px 6px 16px rgba(148,163,184,.16), -4px -4px 12px rgba(255,255,255,.7)",
        border:`1px solid ${C.b}`}}>
        <div style={{display:"flex",gap:8,marginBottom:10}}>
          <div style={{position:"relative",flex:1}}>
            <span style={{position:"absolute",left:12,top:"50%",transform:"translateY(-50%)",color:C.tL,display:"flex"}}>
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="11" cy="11" r="7"/><path d="m21 21-4.35-4.35"/></svg>
            </span>
            <input style={{...iS,width:"100%",padding:"11px 14px 11px 36px",fontSize:14,borderRadius:12}} placeholder="Cari pasien, tindakan, dokter..." value={q} onChange={e=>hQ(e.target.value)}/>
          </div>
          {q&&<button onClick={()=>hQ("")} style={{background:C.dBg,border:`1px solid ${C.dL}`,borderRadius:12,color:C.d,padding:"0 16px",cursor:"pointer",fontSize:16,fontWeight:700}}>✕</button>}
        </div>
        <div style={{display:"flex",gap:6,flexWrap:"wrap",alignItems:"center"}}>
          {[{v:"all",l:"Semua",icon:null},{v:"scheduled",l:"Terjadwal",icon:STS.scheduled.icon},{v:"ongoing",l:"Berlangsung",icon:STS.ongoing.icon},{v:"done",l:"Selesai",icon:STS.done.icon},{v:"batal",l:"Batal",icon:STS.batal.icon}].map(f=>(
            <button key={f.v} onClick={()=>hF(f.v)} style={{padding:"7px 14px",borderRadius:20,border:`1.5px solid ${fSt===f.v?C.p:C.b}`,
              background:fSt===f.v?`linear-gradient(135deg,${C.p},${C.p}dd)`:"#FAFBFD",color:fSt===f.v?"#fff":C.g,fontSize:12,
              fontWeight:fSt===f.v?700:500,cursor:"pointer",fontFamily:"inherit",display:"inline-flex",alignItems:"center",gap:5,
              boxShadow:fSt===f.v?`0 2px 8px ${C.p}44`:"none",transition:"all .15s"}}>
              {f.icon && <span style={{fontSize:9}}>{f.icon}</span>}{f.l}
            </button>
          ))}
          <span style={{marginLeft:"auto",fontSize:12,color:C.tL,fontWeight:600}}>{sorted.length} jadwal</span>
          <button onClick={()=>{
            const printDate = prompt("Cetak jadwal tanggal (format YYYY-MM-DD, mis. 2026-06-29 — kosongkan untuk semua):");
            const filtered = printDate ? ops.filter((o:any)=>o.date===printDate) : sorted;
            if(!filtered.length){showToast("Tidak ada jadwal untuk dicetak",C.w);return;}
            const rows = filtered.map((op:any,i:number)=>`<tr><td>${i+1}</td><td>${fDMY(op.date)} ${op.time}</td><td><b>${op.patient}</b></td><td>${op.procedure}</td><td>${op.room}</td><td>${op.surgeon}</td><td>${op.anesthesiologist||"—"}</td><td style="text-align:center"><span style="background:${STS[op.status as keyof typeof STS]?.c||C.g};color:#fff;padding:2px 8px;border-radius:10px;font-size:10px">${STS[op.status as keyof typeof STS]?.l||op.status}</span></td></tr>`).join("");
            const w=window.open("","_blank","width=900,height=700")!;
            w.document.write(`<!DOCTYPE html><html><head><title>Jadwal Operasi ${printDate||"Semua"} — ${HOSPITAL}</title><style>body{font-family:Arial,sans-serif;font-size:12px;margin:20px}h2{color:#16685F;margin-bottom:4px}p{color:#555;margin-bottom:12px}table{width:100%;border-collapse:collapse}th{background:#16685F;color:#fff;padding:7px 10px;text-align:left;font-size:11px}td{padding:6px 10px;border-bottom:1px solid #e0e0e0;vertical-align:top}tr:nth-child(even)td{background:#f5f5f5}@media print{button{display:none}}</style></head><body><h2>🏥 Jadwal Operasi — ${HOSPITAL}</h2><p>${printDate?`Tanggal: ${printDate}`:"Semua jadwal"} · Dicetak: ${fNow()}</p><button onclick="window.print()" style="margin-bottom:12px;background:#16685F;color:#fff;border:none;padding:8px 18px;border-radius:6px;cursor:pointer;font-size:13px">🖨 Cetak / Simpan PDF</button><table><thead><tr><th>#</th><th>Tanggal/Jam</th><th>Pasien</th><th>Tindakan</th><th>Kamar</th><th>Dr. Bedah</th><th>Dr. Anestesi</th><th>Status</th></tr></thead><tbody>${rows}</tbody></table></body></html>`);
            w.document.close();
          }} style={{padding:"5px 12px",borderRadius:20,border:`1px solid ${C.p}`,background:C.pBg,color:C.p,fontSize:12,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>
            🖨 Cetak
          </button>
        </div>
      </div>
      {initialLoading && sorted.length===0 ? (
        <SkeletonList count={4}/>
      ) : sorted.length===0 && (
        <Card><div style={{textAlign:"center",padding:28,color:C.tL}}>
          <div style={{fontSize:36,marginBottom:8}}>📋</div>
          <div style={{fontSize:14,fontWeight:600}}>{activeOpsBase.length===0?"Tidak ada jadwal operasi yang akan datang.":"Tidak ada hasil pencarian"}</div>
          {activeOpsBase.length===0&&<div style={{fontSize:12,marginTop:4}}>Klik tab 📝 Daftar untuk mendaftarkan</div>}
        </div></Card>
      )}
      {paged.map((op:any, _idx:number)=>{
        const sc=STS[op.status as keyof typeof STS]||STS.scheduled, ot=OT[op.opType as keyof typeof OT||"elektif"];
        const hasAlg=op.allergy&&op.allergy!=="Tidak Ada", isBatal=op.status==="batal";
        const isDone   = op.status==="done";
        const isToday  = op.date===todayStr;
        const isTomorrow = op.date === tmrwDate();
        const isActive = isToday && !isDone && !isBatal;

        if(isDone && isToday) return (
          <Card key={op.id} style={{opacity:.75,transition:"opacity .3s",marginBottom:6,animation:`kbTabFadeIn .35s ease-out ${Math.min(_idx*35,350)}ms both`}}>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <span style={{fontSize:11,color:C.s,fontWeight:700,flexShrink:0}}>✅</span>
              <span style={{fontSize:12,color:"#475569",flex:1,textDecoration:"line-through",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{maskName(op.patient,privacyMode)} — {op.time} — {op.procedure}</span>
              <Bdg label="Selesai" color={C.s} bg={C.sBg}/>
              <Btn sm outline color={C.p} onClick={()=>startEditOp(op)}><span style={{display:"inline-flex",alignItems:"center",gap:5}}><ActionIcon k="edit"/> Edit</span></Btn>
            </div>
          </Card>
        );

        const isCytoGlowActive = op.opType==="cyto" && !isBatal;
        const accentColor = isCytoGlowActive?"#B71C1C":isBatal?"#94A3B8":op.status==="ongoing"?"#00897B":op.status==="done"?"#2E7D32":"#4338CA";
        return (
          <Card key={op.id}
            hi={isCytoGlowActive?"#F44336":hasAlg&&!isBatal?C.dL:undefined}
            className={isCytoGlowActive?"kb-cyto-glow":undefined}
            style={{
              opacity:isBatal?.72:1,
              ...(isCytoGlowActive ? {} : {
                background:isActive?"#fef2f2":(isTomorrow&&!isBatal&&!isDone)?"#F3E8FF":undefined,
                border:isActive?"1.5px solid #fca5a5":(isTomorrow&&!isBatal&&!isDone)?"1.5px solid #D8B4FE":undefined,
              }),
              transition:"all .3s",
              position:"relative",
              paddingLeft:24,
              overflow:"hidden",
              animation:`kbTabFadeIn .35s ease-out ${Math.min(_idx*35,350)}ms both`,
            }}
          >
            <div style={{position:"absolute",top:0,left:0,bottom:0,width:5,background:accentColor,
              ...(op.status==="ongoing"&&!isBatal ? {animation:"kbPulse 1.6s ease-in-out infinite"} : {})}}/>
            {op.opType==="cyto"&&!isBatal && <div style={{background:"#FFCDD2",borderRadius:8,padding:"6px 12px",marginBottom:8,fontSize:12,fontWeight:700,color:"#B71C1C"}}>⚠️ CYTO / EMERGENCY — SEGERA DITANGANI</div>}
            {hasAlg&&op.opType!=="cyto"&&!isBatal && <div style={{background:C.dBg,borderRadius:8,padding:"5px 12px",marginBottom:8,fontSize:12,fontWeight:700,color:C.d}}>⚠ ALERGI: {op.allergy}</div>}
            {isTomorrow&&!isBatal&&!isDone && <div style={{background:"#EDE9FE",borderRadius:8,padding:"5px 12px",marginBottom:8,fontSize:12,fontWeight:700,color:"#6B21A8",display:"flex",alignItems:"center",gap:6}}>📅 OPERASI BESOK (H+1) — Siapkan persiapan sebelum shift selesai</div>}
            {isBatal && <div style={{background:C.dBg,borderRadius:8,padding:"5px 12px",marginBottom:8,fontSize:12,fontWeight:700,color:C.d}}>✕ {op.cancelReason?"BATAL: "+op.cancelReason:"BATAL/DITUNDA"}</div>}

            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8,gap:8}}>
              <div style={{flex:1,minWidth:0}}>
                <div style={{display:"flex",alignItems:"center",gap:7,marginBottom:2}}>
                  {isActive && (
                    <span style={{width:10,height:10,borderRadius:"50%",background:"#ef4444",flexShrink:0,display:"inline-block",animation:"redPulse 1.2s ease-in-out infinite"}}/>
                  )}
                  <div style={{fontSize:15,fontWeight:800,color:(isTomorrow&&!isBatal&&!isDone)?"#6B21A8":C.t,textDecoration:isBatal?"line-through":"none",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{maskName(op.patient,privacyMode)||"—"}</div>
                </div>
                <div style={{fontSize:12,color:C.tL}}>{op.age?op.age+"th  ":""}{op.rm?"RM "+maskRM(op.rm,privacyMode)+"  ":""}{op.bloodType}</div>
              </div>
              <div style={{display:"flex",flexDirection:"column",gap:4,alignItems:"flex-end",flexShrink:0}}>
                <Bdg label={ot?.label||"Elektif"} color={ot?.c||C.i} bg={ot?.bg||C.iBg}/>
                <Bdg label={sc.l} color={sc.c} bg={sc.bg} icon={sc.icon}/>
              </div>
              {!isBatal && (
                <div style={{display:"flex",gap:6,flexShrink:0,alignItems:"center"}}>
                  {op.status==="scheduled" && (
                    <button onClick={()=>{
                      /* Pengaman: satu kamar tidak bisa menjalankan dua operasi sekaligus.
                         Kalau masih ada operasi lain BERLANGSUNG di kamar yang sama (biasanya
                         karena lupa menekan Selesai), tawarkan untuk menutupnya dulu. */
                      if(op.room){
                        const prev=ops.filter((o:any)=>o.id!==op.id && o.status==="ongoing" && o.room===op.room);
                        if(prev.length>0){
                          const nama=prev.slice(0,3).map((o:any)=>maskName(o.patient,privacyMode)).join(", ")+(prev.length>3?" dan "+(prev.length-3)+" lainnya":"");
                          if(!window.confirm(`Di ${op.room} masih tercatat BERLANGSUNG: ${nama}. Tandai operasi itu SELESAI lalu mulai operasi ini? (OK = ya, Batal = tidak ada perubahan)`)) return;
                          const nowIso=new Date().toISOString();
                          prev.forEach((o:any)=>{
                            const closed={...o,status:"done",updated_at:nowIso};
                            setOps((p:any)=>p.map((x:any)=>x.id===o.id?closed:x));
                            AmbilLogMedis("JADWAL_OPERASI","Sistem","TUTUP_OTOMATIS_SEBELUM_MULAI","Operasi ditandai selesai karena operasi berikutnya dimulai di kamar yang sama",{id:o.id,room:o.room,next:op.id});
                            upsertOneToSupa("kb_operasi",closed).then((res:any)=>{ if(!res?.ok){ setOps((p:any)=>p.map((x:any)=>x.id===o.id?o:x)); showToast(`⚠ Gagal menutup operasi sebelumnya: ${res?.error||"error"}`,C.d); } });
                          });
                        }
                      }
                      const upd={...op,status:"ongoing",updated_at:new Date().toISOString()};
                      setOps((p:any)=>p.map((o:any)=>o.id===op.id?upd:o));
                      AmbilLogMedis("JADWAL_OPERASI","Admin","MULAI_OPERASI",
                        `Operasi dimulai: ${op.patient||"-"} — ${op.procedure||"-"}`,{id:op.id,patient:op.patient,room:op.room});
                      upsertOneToSupa("kb_operasi",upd).then((res:any)=>{ if(!res?.ok){ setOps((p:any)=>p.map((o:any)=>o.id===op.id?op:o)); showToast(`⚠ Gagal: ${res?.error||"error"}`,C.d); } });
                    }} style={{background:"#22c55e",color:"#fff",border:"none",borderRadius:10,padding:"10px 18px",fontSize:14,fontWeight:800,cursor:"pointer",fontFamily:"inherit",boxShadow:"0 2px 8px rgba(34,197,94,.4)",display:"inline-flex",alignItems:"center",gap:6}}>
                      <ActionIcon k="play" size={13}/> Mulai
                    </button>
                  )}
                  {op.status==="ongoing" && (
                    <button onClick={()=>{
                      const upd={...op,status:"done",updated_at:new Date().toISOString()};
                      setOps((p:any)=>p.map((o:any)=>o.id===op.id?upd:o));
                      AmbilLogMedis("JADWAL_OPERASI","Admin","SELESAI_OPERASI",
                        `Operasi selesai: ${op.patient||"-"} — ${op.procedure||"-"}`,{id:op.id,patient:op.patient,room:op.room});
                      upsertOneToSupa("kb_operasi",upd).then((res:any)=>{ if(!res?.ok){ setOps((p:any)=>p.map((o:any)=>o.id===op.id?op:o)); showToast(`⚠ Gagal: ${res?.error||"error"}`,C.d); } });
                    }} style={{background:"#ef4444",color:"#fff",border:"none",borderRadius:10,padding:"10px 18px",fontSize:14,fontWeight:800,cursor:"pointer",fontFamily:"inherit",boxShadow:"0 2px 8px rgba(239,68,68,.4)",display:"inline-flex",alignItems:"center",gap:6}}>
                      <ActionIcon k="check" size={13}/> Selesai
                    </button>
                  )}
                </div>
              )}
            </div>

            <div style={{background:"#F8FBF0",border:"1px solid #D4E6A0",borderRadius:10,padding:"10px 12px",marginBottom:8}}>
              <div style={{fontSize:14,fontWeight:700,color:"#33691E"}}>{op.procedure}</div>
              <div style={{fontSize:12,color:"#558B2F",marginTop:2}}>{op.diagnosis}</div>
              <div style={{fontSize:12,color:"#558B2F",marginTop:3}}>📅 {fDMY(op.date)} · {op.time} WIB · {op.room}</div>
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
                <Btn sm outline color={C.p} onClick={()=>startEditOp(op)}><span style={{display:"inline-flex",alignItems:"center",gap:5}}><ActionIcon k="edit"/> Edit</span></Btn>
                <Btn sm outline color={C.w} onClick={()=>{setReqOpId(op.id);setReqText("");}}><span style={{display:"inline-flex",alignItems:"center",gap:5}}><ActionIcon k="plus-circle"/> Permintaan</span></Btn>
                <Btn sm outline color={C.d} onClick={()=>{setCId(op.id);setCR("");}}><span style={{display:"inline-flex",alignItems:"center",gap:5}}><ActionIcon k="x-circle"/> Batal/Tunda</span></Btn>
              </div>
            )}
            <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
              {isBatal && <Btn sm outline color={C.g} onClick={()=>{
                const upd={...op,status:"scheduled",cancelReason:"",updated_at:new Date().toISOString()};
                setOps((p:any)=>p.map((o:any)=>o.id===op.id?upd:o));
                AmbilLogMedis("JADWAL_OPERASI","Admin","AKTIFKAN_OPERASI",
                  `Operasi diaktifkan kembali: ${op.patient||"-"}`,{id:op.id,patient:op.patient});
                upsertOneToSupa("kb_operasi",upd).then((res:any)=>{ if(!res?.ok){ setOps((p:any)=>p.map((o:any)=>o.id===op.id?op:o)); showToast(`⚠ Gagal: ${res?.error||"error"}`,C.d); } });
              }}><span style={{display:"inline-flex",alignItems:"center",gap:5}}><ActionIcon k="refresh"/> Aktifkan</span></Btn>}
              <Btn sm outline color={C.d} onClick={()=>setDelC(op.id)}><span style={{display:"inline-flex",alignItems:"center",gap:5}}><ActionIcon k="trash"/> Hapus</span></Btn>
            </div>
            {cId===op.id && (
              <div style={{marginTop:12,paddingTop:12,borderTop:`1px solid ${C.b}`}}>
                <div style={{fontSize:13,fontWeight:700,color:C.d,marginBottom:8}}>Alasan Pembatalan / Penundaan</div>
                <input style={{...iS,marginBottom:8}} placeholder="Cth: Pasien belum puasa, kondisi tidak stabil" value={cR} onChange={e=>setCR(e.target.value)}/>
                <div style={{display:"flex",gap:8}}><Btn full color={C.d} onClick={()=>{
                  const upd={...op,status:"batal",cancelReason:sanitize(cR),updated_at:new Date().toISOString()};
                  setOps((p:any)=>p.map((o:any)=>o.id===op.id?upd:o));
                  setCId(null);setCR("");showToast("Operasi ditandai batal",C.w);
                  AmbilLogMedis("JADWAL_OPERASI","Admin","BATAL_OPERASI",
                    `Operasi dibatalkan: ${op.patient||"-"} — alasan: ${cR||"-"}`,
                    {id:op.id,patient:op.patient,cancelReason:cR});
                  upsertOneToSupa("kb_operasi",upd).then((res:any)=>{ if(!res?.ok){ setOps((p:any)=>p.map((o:any)=>o.id===op.id?op:o)); showToast(`⚠ Gagal: ${res?.error||"error"}`,C.d); } });
                }}>Konfirmasi Batal</Btn><Btn full outline color={C.g} onClick={()=>setCId(null)}>Tutup</Btn></div>
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
      {pastOps.length>0 && (
        <details style={{marginTop:16}} onToggle={(e:any)=>setShowRiwayat(e.target.open)}>
          <summary style={{background:"#F0F4F8",borderRadius:12,padding:"12px 16px",cursor:"pointer",border:`1px solid ${C.b}`,listStyle:"none",fontWeight:700,fontSize:13,color:C.g}}>
            📜 Riwayat — {pastOps.length} jadwal sudah berlalu
          </summary>
          {showRiwayat && (
            <div style={{paddingTop:10}}>
              {pastOps.map((op:any)=>{
                const sc=STS[op.status as keyof typeof STS]||STS.scheduled;
                return (
                  <Card key={op.id} style={{padding:"10px 14px",marginBottom:6,opacity:.85}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
                      <div>
                        <div style={{fontSize:13,fontWeight:700,color:C.t}}>{maskName(op.patient,privacyMode)}</div>
                        <div style={{fontSize:12,color:C.tL,marginTop:2}}>{op.procedure} · {op.room}</div>
                        <div style={{fontSize:11,color:C.tL,marginTop:2}}>{fDMY(op.date)} · {op.time} WIB · {op.surgeon}</div>
                      </div>
                      <span style={{background:sc.c,color:"#fff",fontSize:10,padding:"3px 9px",borderRadius:10,fontWeight:700,whiteSpace:"nowrap"}}>{sc.l}</span>
                    </div>
                  </Card>
                );
              })}
            </div>
          )}
        </details>
      )}
      {delC && (
        <div className="kb-modal-backdrop" style={{position:"fixed",inset:0,background:"rgba(15,23,42,.45)",backdropFilter:"blur(6px)",WebkitBackdropFilter:"blur(6px)",zIndex:400,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
          <div style={{background:C.white,borderRadius:18,padding:"24px 20px",maxWidth:320,width:"100%",boxShadow:"0 12px 40px rgba(0,0,0,.25)"}}>
            <div style={{textAlign:"center",marginBottom:14}}><div style={{fontSize:36}}>🗑</div><div style={{fontSize:15,fontWeight:700,color:C.d,marginTop:6}}>Hapus Jadwal?</div><div style={{fontSize:12,color:C.tL,marginTop:4}}>Tindakan ini tidak dapat dibatalkan.</div></div>
            <div style={{display:"flex",gap:8}}><Btn full outline color={C.g} onClick={()=>setDelC(null)} style={{flex:1}}>Batal</Btn><Btn full color={C.d} onClick={()=>{deleteOp(delC);setDelC(null);}} style={{flex:1}}>Hapus</Btn></div>
          </div>
        </div>
      )}
      {cytoM && (
        <div className="kb-modal-backdrop" style={{position:"fixed",inset:0,background:"rgba(15,23,42,.5)",backdropFilter:"blur(6px)",WebkitBackdropFilter:"blur(6px)",zIndex:500,display:"flex",alignItems:"flex-end",justifyContent:"center"}}>
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
