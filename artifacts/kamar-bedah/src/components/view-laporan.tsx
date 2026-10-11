/* src/components/view-laporan.tsx
   Tab "Laporan" — laporan Kepala Jaga (salam, penerima, keterangan, acara
   hari ini, rencana besok — dengan penanganan khusus hari Sabtu yang
   merender Minggu & Senin terpisah, siaga anestesi/cyto/perawat, dan
   preview wajib sebelum kirim WA ke Katim/Grup). Dipindah dari App.tsx
   (langkah 6 refactor — pisah View/layar besar, batch 7), TIDAK ADA
   perubahan logika, termasuk seluruh komentar teknis di dalamnya. */

import { useState, useEffect } from "react";
import type { ViewLaporanProps } from "../types";
import { C, OT, STS, iS, DEFAULT_RECIPIENT } from "../constants";
import { gWord, toLocalISODate, tmrwDate, todayDate, fDMY, maskName, normalizePhone62, buildLaporan } from "../utils";
import { Row, Card, SH, LF, Bdg, Btn, WaBt } from "./primitives";

export function ViewLaporan({ops,staff,roster,showToast,role,privacyMode}: ViewLaporanProps) {
  const [lSet, setLSet] = useState({greeting:gWord(),recipients:"",keterangan:"",pembawaHP:"",katimPhone:"",grupPhone:""});
  const [lSby, setLSby] = useState({anest:["",""],cyto:["","",""],nurses:["","","",""],pembawaHP:""});
  const [preview,setPreview] = useState("");
  const [foc,setFoc]         = useState<string|null>(null);
  const todayOps = ops.filter((o:any)=>o.date===todayDate());
  /* ── Poin ④: Manajemen tanggal anti-overwrite (khusus Hari Sabtu) ──
     Saat hari ini Sabtu, Minggu (H+1) dan Senin (H+2) WAJIB diinstansiasi
     sebagai dua objek Date terpisah di memori — TIDAK boleh mutasi setDate
     beruntun pada satu referensi objek yang sama (mencegah data leakage
     antar hari). Setiap objek dikonversi ke string YYYY-MM-DD memakai
     waktu LOKAL (toLocalISODate — bukan .toISOString() yang berbasis UTC),
     supaya tanggal tidak bergeser mundur. */
  const _todayObj   = new Date();
  const _isSaturday = _todayObj.getDay() === 6;
  /* Objek tanggal Minggu (H+1) — instansiasi independen, bukan turunan _todayObj yang dimutasi */
  const _mingguObj  = new Date(_todayObj.getFullYear(), _todayObj.getMonth(), _todayObj.getDate() + 1);
  const _mingguStr  = toLocalISODate(_mingguObj);
  /* Objek tanggal Senin (H+2) — instansiasi independen kedua, terpisah dari objek Minggu */
  const _seninObj   = new Date(_todayObj.getFullYear(), _todayObj.getMonth(), _todayObj.getDate() + 2);
  const _seninStr   = toLocalISODate(_seninObj);
  /* _besokStr/_lusaStr dipertahankan untuk kompatibilitas dengan tmrwOps (hari non-Sabtu) */
  const _besokStr   = tmrwDate();
  const _lusaStr    = _seninStr;
  const tmrwOps = ops
    .filter((o:any)=>{
      if(o.status==="batal") return false;
      return _isSaturday
        ? (o.date===_besokStr || o.date===_lusaStr)
        : o.date===_besokStr;
    })
    .sort((a:any,b:any)=> a.date!==b.date ? a.date.localeCompare(b.date) : a.time.localeCompare(b.time));
  /* ── Render terpisah khusus Sabtu: filter ketat per-hari, tidak digabung ── */
  const mingguOps = _isSaturday
    ? ops.filter((o:any)=>o.status!=="batal" && o.date===_mingguStr).sort((a:any,b:any)=>a.time.localeCompare(b.time))
    : [];
  const seninOps = _isSaturday
    ? ops.filter((o:any)=>o.status!=="batal" && o.date===_seninStr).sort((a:any,b:any)=>a.time.localeCompare(b.time))
    : [];
  const rTmrw    = roster.find((r:any)=>r.date===todayDate());
  const tActive  = todayOps.filter((o:any)=>o.status!=="batal");
  const tBatal   = todayOps.filter((o:any)=>o.status==="batal");

  useEffect(()=>{
    if(rTmrw) setLSby((p:any)=>({
      ...p,
      anest: (()=>{
        const a = rTmrw.anestJagaList?.length ? rTmrw.anestJagaList : (rTmrw.anestJaga ? rTmrw.anestJaga.split(/[,;]/).map((x:string)=>x.trim()).filter(Boolean) : []);
        const cur = p.anest.filter((x:any)=>x);
        const base = a.length>0 ? [...a] : [...cur];
        while(base.length<2) base.push("");
        return base.slice(0,2).concat(base.slice(2));
      })(),
      cyto: (()=>{
        const c = rTmrw.anestCytoList?.length ? rTmrw.anestCytoList : (rTmrw.anestCyto ? rTmrw.anestCyto.split(/[,;]/).map((x:string)=>x.trim()).filter(Boolean) : []);
        const cur = p.cyto.filter((x:any)=>x);
        const base = c.length>0 ? [...c] : [...cur];
        while(base.length<3) base.push("");
        return base.slice(0,3).concat(base.slice(3));
      })(),
      nurses: (()=>{
        const n = rTmrw.nurses?.filter(Boolean).length ? rTmrw.nurses.filter(Boolean) : (p.nurses.filter((x:any)=>x).length ? p.nurses.filter((x:any)=>x) : []);
        while(n.length<4) n.push("");
        return n;
      })(),
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
      {rTmrw ? <div style={{background:C.sBg,borderRadius:10,padding:"8px 14px",marginBottom:12,fontSize:12,color:C.s,border:`1px solid ${C.s}33`}}>✅ Jadwal siaga hari ini ditemukan — kolom siaga terisi otomatis. Bisa diedit.</div>
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
        {tActive.length===0
          ? <div style={{fontSize:13,color:C.tL,textAlign:"center",padding:"6px 0",fontStyle:"italic"}}>Hari Ini Tidak Ada Acara Operasi</div>
          : [...tActive].sort((a:any,b:any)=>a.time.localeCompare(b.time)).map((op:any,i:number)=>(
            <div key={op.id} style={{display:"flex",gap:10,marginBottom:8,paddingBottom:8,borderBottom:i<tActive.length-1?`1px solid ${C.b}`:"none",fontStyle:"italic"}}>
              <div style={{background:op.opType==="cyto"?"#FFCDD2":C.sBg,borderRadius:8,padding:"5px 9px",fontSize:12,fontWeight:700,color:op.opType==="cyto"?"#B71C1C":C.s,flexShrink:0,minWidth:52,textAlign:"center"}}>{op.time}</div>
              <div style={{flex:1}}>
                <div style={{fontSize:13,fontWeight:700,color:C.t}}>{maskName(op.patient,privacyMode)}{op.age?`, ${op.age}th`:""}</div>
                <div style={{fontSize:11,color:C.tL}}>{op.surgeon} / {op.anesthesiologist||"—"}</div>
              </div>
              <Bdg label={STS[op.status as keyof typeof STS]?.l||"Terjadwal"} color={STS[op.status as keyof typeof STS]?.c||C.i} bg={STS[op.status as keyof typeof STS]?.bg||C.iBg} icon={STS[op.status as keyof typeof STS]?.icon}/>
            </div>
          ))
        }
        {tBatal.length>0 && <div style={{fontSize:12,color:C.d,marginTop:6,borderTop:`1px dashed ${C.dL}`,paddingTop:6}}><b>Batal/Tunda:</b> {tBatal.map((o:any)=>maskName(o.patient,privacyMode)+(o.cancelReason?` (${o.cancelReason})`:""  )).join("; ")}</div>}
      </Card>
      <Card style={{background:"#EDE7F6",border:"1px solid #D1C4E9"}}>
        <SH label="④ Rencana Operasi Besok (H-1)" color="#512DA8"/>
        {!_isSaturday ? (
          /* ── Mode normal (bukan Sabtu): satu blok seperti biasa ── */
          tmrwOps.length===0 ? <div style={{fontSize:13,color:C.tL,textAlign:"center",padding:"6px 0"}}>Belum ada rencana operasi besok</div>
            : <>
              <div style={{fontSize:12,color:"#512DA8",marginBottom:10}}>Total <b>{tmrwOps.length}</b> · Pagi (&lt;14.00): <b>{tmrwOps.filter((o:any)=>o.time<"14:00").length}</b> · Sore (≥14.00): <b>{tmrwOps.filter((o:any)=>o.time>="14:00").length}</b></div>
              {tmrwOps.sort((a:any,b:any)=>a.time.localeCompare(b.time)).map((op:any,i:number)=>(
                <div key={op.id} style={{display:"flex",gap:10,marginBottom:8,paddingBottom:8,borderBottom:i<tmrwOps.length-1?"1px solid #D1C4E9":"none"}}>
                  <div style={{background:"#D1C4E9",borderRadius:8,padding:"5px 9px",fontSize:12,fontWeight:700,color:"#512DA8",flexShrink:0,minWidth:52,textAlign:"center"}}>{op.time}</div>
                  <div style={{flex:1}}>
                    <div style={{fontSize:13,fontWeight:700,color:C.t}}>{maskName(op.patient,privacyMode)}{op.age?`, ${op.age}th`:""}</div>
                    <div style={{fontSize:12,color:C.tL}}>{op.procedure}</div>
                    <div style={{fontSize:11,color:C.tL}}>{op.surgeon} / {op.anesthesiologist||"—"}</div>
                  </div>
                  <Bdg label={OT[op.opType as keyof typeof OT||"elektif"]?.label} color={OT[op.opType as keyof typeof OT||"elektif"]?.c||C.i} bg={OT[op.opType as keyof typeof OT||"elektif"]?.bg||C.iBg}/>
                </div>
              ))}
            </>
        ) : (
          /* ── Mode Sabtu: Minggu & Senin DIRENDER TERPISAH, masing-masing
               dengan filter tanggal lokal sendiri — tidak saling menimpa.
               Minggu: kalau memang ada jadwal, tetap ditampilkan apa adanya;
               statis "Tidak ada acara operasi" hanya jika benar-benar kosong. ── */
          <>
            <div style={{marginBottom:14}}>
              <div style={{fontSize:12,fontWeight:700,color:"#512DA8",textTransform:"uppercase",letterSpacing:.4,marginBottom:8}}>Rencana Hari Minggu ({fDMY(_mingguStr)})</div>
              {mingguOps.length===0 ? (
                <div style={{fontSize:13,color:C.tL,textAlign:"center",padding:"6px 0",fontStyle:"italic"}}>Tidak ada acara operasi</div>
              ) : (
                <>
                  <div style={{fontSize:12,color:"#512DA8",marginBottom:10}}>Total <b>{mingguOps.length}</b> · Pagi (&lt;14.00): <b>{mingguOps.filter((o:any)=>o.time<"14:00").length}</b> · Sore (≥14.00): <b>{mingguOps.filter((o:any)=>o.time>="14:00").length}</b></div>
                  {mingguOps.map((op:any,i:number)=>(
                    <div key={op.id} style={{display:"flex",gap:10,marginBottom:8,paddingBottom:8,borderBottom:i<mingguOps.length-1?"1px solid #D1C4E9":"none"}}>
                      <div style={{background:"#D1C4E9",borderRadius:8,padding:"5px 9px",fontSize:12,fontWeight:700,color:"#512DA8",flexShrink:0,minWidth:52,textAlign:"center"}}>{op.time}</div>
                      <div style={{flex:1}}>
                        <div style={{fontSize:13,fontWeight:700,color:C.t}}>{maskName(op.patient,privacyMode)}{op.age?`, ${op.age}th`:""}</div>
                        <div style={{fontSize:12,color:C.tL}}>{op.procedure}</div>
                        <div style={{fontSize:11,color:C.tL}}>{op.surgeon} / {op.anesthesiologist||"—"}</div>
                      </div>
                      <Bdg label={OT[op.opType as keyof typeof OT||"elektif"]?.label} color={OT[op.opType as keyof typeof OT||"elektif"]?.c||C.i} bg={OT[op.opType as keyof typeof OT||"elektif"]?.bg||C.iBg}/>
                    </div>
                  ))}
                </>
              )}
            </div>
            <div style={{borderTop:"1px dashed #D1C4E9",paddingTop:12}}>
              <div style={{fontSize:12,fontWeight:700,color:"#512DA8",textTransform:"uppercase",letterSpacing:.4,marginBottom:8}}>Rencana Hari Senin ({fDMY(_seninStr)})</div>
              {seninOps.length===0 ? (
                <div style={{fontSize:13,color:C.tL,textAlign:"center",padding:"6px 0",fontStyle:"italic"}}>Belum ada rencana operasi Hari Senin</div>
              ) : (
                <>
                  <div style={{fontSize:12,color:"#512DA8",marginBottom:10}}>Total <b>{seninOps.length}</b> · Pagi (&lt;14.00): <b>{seninOps.filter((o:any)=>o.time<"14:00").length}</b> · Sore (≥14.00): <b>{seninOps.filter((o:any)=>o.time>="14:00").length}</b></div>
                  {seninOps.map((op:any,i:number)=>(
                    <div key={op.id} style={{display:"flex",gap:10,marginBottom:8,paddingBottom:8,borderBottom:i<seninOps.length-1?"1px solid #D1C4E9":"none"}}>
                      <div style={{background:"#D1C4E9",borderRadius:8,padding:"5px 9px",fontSize:12,fontWeight:700,color:"#512DA8",flexShrink:0,minWidth:52,textAlign:"center"}}>{op.time}</div>
                      <div style={{flex:1}}>
                        <div style={{fontSize:13,fontWeight:700,color:C.t}}>{maskName(op.patient,privacyMode)}{op.age?`, ${op.age}th`:""}</div>
                        <div style={{fontSize:12,color:C.tL}}>{op.procedure}</div>
                        <div style={{fontSize:11,color:C.tL}}>{op.surgeon} / {op.anesthesiologist||"—"}</div>
                      </div>
                      <Bdg label={OT[op.opType as keyof typeof OT||"elektif"]?.label} color={OT[op.opType as keyof typeof OT||"elektif"]?.c||C.i} bg={OT[op.opType as keyof typeof OT||"elektif"]?.bg||C.iBg}/>
                    </div>
                  ))}
                </>
              )}
            </div>
          </>
        )}
      </Card>
      {/* ⑤ Siaga Dokter Anestesi — 2 kolom fixed */}
      <Card>
        <SH label="⑤ Siaga Dokter Anestesi" color={C.p}/>
        {rTmrw && <div style={{fontSize:11,color:C.s,marginBottom:8,fontWeight:600}}>✓ Terisi otomatis dari jadwal jaga {fDMY(todayDate())} — bisa diedit</div>}
        <div style={{display:"flex",gap:8}}>
          {[0,1].map(i=>(
            <div key={i} style={{flex:1}}>
              <div style={{fontSize:11,fontWeight:700,color:C.g,marginBottom:4,textTransform:"uppercase",letterSpacing:.4}}>Dokter {i+1}</div>
              <input style={iS} placeholder={`Dokter Anestesi ${i+1}`} value={(lSby.anest||["",""])[i]||""} onChange={(ev:any)=>setI("anest",i,ev.target.value)} spellCheck={false}/>
            </div>
          ))}
        </div>
      </Card>
      {/* ⑥ Siaga Cyto Dokter Anestesi — 3 kolom fixed */}
      <Card hi={C.dL}>
        <SH label="⑥ Siaga Cyto Dokter Anestesi" color="#B71C1C"/>
        {rTmrw && <div style={{fontSize:11,color:C.s,marginBottom:8,fontWeight:600}}>✓ Terisi otomatis dari jadwal jaga {fDMY(todayDate())} — bisa diedit</div>}
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          {[0,1,2].map(i=>(
            <div key={i} style={{flex:"1 1 calc(33% - 8px)",minWidth:120}}>
              <div style={{fontSize:11,fontWeight:700,color:C.g,marginBottom:4,textTransform:"uppercase",letterSpacing:.4}}>Dokter {i+1}</div>
              <input style={iS} placeholder={`Dokter Cyto ${i+1}`} value={(lSby.cyto||["","",""])[i]||""} onChange={(ev:any)=>setI("cyto",i,ev.target.value)} spellCheck={false}/>
            </div>
          ))}
        </div>
      </Card>
      {/* ⑦ Siaga Perawat — 4 kolom fixed */}
      <Card>
        <SH label="⑦ Siaga Perawat" color={C.p}/>
        {rTmrw && <div style={{fontSize:11,color:C.s,marginBottom:8,fontWeight:600}}>✓ Terisi otomatis dari jadwal jaga {fDMY(todayDate())} — bisa diedit</div>}
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          {[0,1,2,3].map(i=>(
            <div key={i} style={{flex:"1 1 calc(50% - 8px)",minWidth:140}}>
              <div style={{fontSize:11,fontWeight:700,color:C.g,marginBottom:4,textTransform:"uppercase",letterSpacing:.4}}>Perawat {i+1}</div>
              <input style={iS} placeholder={`Siaga Perawat ${i+1}`} value={(lSby.nurses||["","","",""])[i]||""} onChange={(ev:any)=>setI("nurses",i,ev.target.value)} spellCheck={false}/>
            </div>
          ))}
        </div>
      </Card>
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
            <WaBt onClick={()=>{ if(!lSet.katimPhone){showToast("Nomor WA Katim belum diisi",C.d);return;} window.open(`https://wa.me/${normalizePhone62(lSet.katimPhone)}?text=${encodeURIComponent(preview)}`,"_blank"); showToast("✓ WA Katim dibuka","#25D366"); }} style={{flex:1}}>WA Katim</WaBt>
            <WaBt onClick={()=>{ if(!lSet.grupPhone){showToast("Nomor WA Grup belum diisi",C.d);return;} window.open(`https://wa.me/${normalizePhone62(lSet.grupPhone)}?text=${encodeURIComponent(preview)}`,"_blank"); showToast("✓ WA Grup dibuka","#25D366"); }} style={{flex:1}}>WA Grup</WaBt>
          </div>
        </Card>
      )}
    </div>
  );
}
