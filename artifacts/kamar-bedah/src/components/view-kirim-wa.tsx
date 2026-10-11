/* src/components/view-kirim-wa.tsx
   Tab "Kirim WA" — kirim pengingat H-1 (besok) ke dokter bedah & anestesi
   lewat WhatsApp, dengan draft pesan yang wajib di-preview dulu sebelum
   dibuka ke WhatsApp (tidak ada auto-send). Dipindah dari App.tsx (langkah
   6 refactor — pisah View/layar besar, batch 3), TIDAK ADA perubahan
   logika, termasuk seluruh FIX AUDIT di dalamnya. */

import { useState } from "react";
import type { ViewKirimWAProps } from "../types";
import { C, HOSPITAL } from "../constants";
import { tmrwDate, fDMY, fNow, normalizePhone62, gId, msgSurgeon, msgAnest } from "../utils";
import { Card, WaBt } from "./primitives";

export function ViewKirimWA({ops,staff,setNotifs,showToast}: ViewKirimWAProps) {
  const [prev,setPrev] = useState<Record<string,boolean>>({});
  const getPhone = (n: string) => staff.find((x:any)=>x.name===n)?.phone||null;
  const tmrwOps  = ops.filter((o:any)=>o.date===tmrwDate()&&o.status!=="batal");
  const sendWA   = (ph: string | null,msg: string,name: string,lbl: string) => {
    /* FIX AUDIT #46 (lanjutan): sama seperti confirmCyto & WA Katim/Grup —
       pakai normalizePhone62() supaya nomor berawalan "0" (format umum
       yang diisi manual di form Staf) tetap membentuk link wa.me yang
       valid, bukan hanya nomor yang kebetulan sudah diketik dalam format
       62xxx sejak awal. */
    if(ph) window.open(`https://wa.me/${normalizePhone62(ph)}?text=${encodeURIComponent(msg)}`,"_blank");
    /* FIX AUDIT #24: sama seperti di ViewJadwal — cap 200 entri terbaru. */
    setNotifs((p:any)=>[{id:gId(),type:"wa_direct",label:lbl,patient:name,procedure:"Jadwal H-1",message:msg,sentAt:fNow()},...p].slice(0,200));
    showToast("✓ WhatsApp dibuka — tekan Kirim di WA","#25D366");
  };
  const surgeons = [...new Set(tmrwOps.map((o:any)=>o.surgeon))].filter(Boolean);
  const anests   = [...new Set(tmrwOps.map((o:any)=>o.anesthesiologist))].filter(Boolean);
  return (
    <div>
      <div style={{background:`linear-gradient(90deg,${C.pBg},${C.waBg})`,borderRadius:14,padding:"14px 16px",marginBottom:14,border:`1px solid ${C.pL}33`}}>
        <div style={{fontSize:14,fontWeight:700,color:C.p}}>Kamar Bedah {HOSPITAL}</div>
        <div style={{fontSize:12,color:C.tL,marginTop:2}}>Pengiriman H-1 — Besok: {fDMY(tmrwDate())} · {tmrwOps.length} operasi</div>
      </div>
      <div style={{background:C.iBg,borderRadius:10,padding:"10px 14px",marginBottom:14,fontSize:12,color:C.i,lineHeight:1.6}}>ℹ️ Pesan bersifat DRAFT. Pratinjau wajib ditampilkan. Tekan Kirim di WhatsApp untuk konfirmasi manual.</div>
      {tmrwOps.length===0 && <Card><div style={{textAlign:"center",padding:20,color:C.tL,fontSize:13}}>Belum ada operasi terjadwal untuk besok ({fDMY(tmrwDate())})</div></Card>}
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
