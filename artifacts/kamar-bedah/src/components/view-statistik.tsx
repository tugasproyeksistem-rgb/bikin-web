/* src/components/view-statistik.tsx
   Tab "Statistik" — ringkasan angka, grafik operasi per dokter/bulan/kamar,
   rekapitulasi tindakan & diagnosa, dan export Excel multi-sheet. Dipindah
   dari App.tsx (langkah 6 refactor — pisah View/layar besar, batch 4),
   TIDAK ADA perubahan logika, termasuk seluruh FIX AUDIT di dalamnya. */

import { useState } from "react";
import * as XLSX from "xlsx";
import { C, OT, STS } from "../constants";
import { fNow, todayDate, sanitizeCellForExcel } from "../utils";
import { Row, Card, SH, MiniBar } from "./primitives";
import { StatistikBulanan } from "./statistik-bulanan";

export function ViewStatistik({ops, archive}: {ops: any[]; archive: any[]}) {
  const [range, setRange] = useState<"all"|"30"|"90">("all");
  const [exporting, setExporting] = useState(false);
  const COLORS = ["#1F5A52","#1565C0","#E65100","#7B1FA2","#2E7D32","#C62828","#00838F","#E91E63"];

  const allOps = [...ops, ...(archive.flatMap((a:any)=>a.ops||[]))];
  const now = new Date();
  const filtered = allOps.filter(op => {
    if(range==="all") return true;
    const d = new Date(op.date+"T00:00:00");
    const days = range==="30"?30:90;
    const diff = now.getTime()-d.getTime();
    /* PENTING: harus cek diff>=0 JUGA, bukan hanya diff<=range.
       Tanpa batas bawah ini, operasi yang TERJADWAL DI MASA DEPAN (status
       "scheduled") punya diff negatif, yang selalu <= days*ms berapa pun
       besarnya — akibatnya operasi bulan depan/tahun depan tetap lolos
       filter "30 Hari" / "90 Hari" dan mencemari statistik. */
    return diff >= 0 && diff <= days*24*60*60*1000;
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

  // Per room
  const byRoom: Record<string,number> = {};
  filtered.forEach((o:any)=>{ if(o.room) byRoom[o.room]=(byRoom[o.room]||0)+1; });
  const roomData = Object.entries(byRoom).sort((a,b)=>b[1]-a[1]).map(([label,value],i)=>({label,value,color:COLORS[i%COLORS.length]}));

  // Per month (last 6 months)
  const monthMap: Record<string,number> = {};
  filtered.forEach((o:any)=>{
    /* FIX #5: Cek tipe dan panjang string sebelum .slice() — jika o.date null/undefined
       (dari data migrasi lama), .slice() akan crash "Cannot read properties of null".
       Perlu typeof === "string" DAN panjang >= 7, bukan sekadar truthy check. */
    if(typeof o.date !== "string" || o.date.length < 7) return;
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

  // Agregasi Tindakan (procedure)
  const byProcedure: Record<string,number> = {};
  filtered.forEach((o:any)=>{ const k=(o.procedure||"").trim(); if(k) byProcedure[k]=(byProcedure[k]||0)+1; });
  const procedureData = Object.entries(byProcedure).sort((a,b)=>b[1]-a[1]);

  // Agregasi Diagnosa
  const byDiagnosis: Record<string,number> = {};
  filtered.forEach((o:any)=>{ const k=(o.diagnosis||"").trim(); if(k) byDiagnosis[k]=(byDiagnosis[k]||0)+1; });
  const diagnosisData = Object.entries(byDiagnosis).sort((a,b)=>b[1]-a[1]);

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
      /* FIX AUDIT #33 (HIGH — CSV/Excel Formula Injection, CWE-1236):
         sebelumnya kolom bebas-teks (pasien, diagnosa, tindakan, nama staf,
         dll) ditulis MENTAH ke aoa_to_sheet, padahal field ini bisa diisi
         manual atau auto-fill OCR (data eksternal). Kalau ada nilai yang
         diawali "=", "+", "-", atau "@", Excel akan menafsirkannya sebagai
         formula saat file dibuka (bisa dipakai untuk RCE/DDE atau eksfiltrasi
         data via =WEBSERVICE(...)). sanitizeCellForExcel() sudah dipakai di
         modul lembur — sekarang diterapkan juga di sini agar konsisten. */
      const detailRows = filtered.map((op:any,i:number)=>[
        i+1, op.date||"", op.time||"",
        sanitizeCellForExcel(op.patient||""), op.age||"", sanitizeCellForExcel(op.rm||""),
        OT[op.opType as keyof typeof OT]?.label||"Elektif",
        sanitizeCellForExcel(op.diagnosis||""), sanitizeCellForExcel(op.procedure||""), sanitizeCellForExcel(op.room||""),
        sanitizeCellForExcel(op.surgeon||""), sanitizeCellForExcel(op.anesthesiologist||""),
        sanitizeCellForExcel(op.assistantNurse||""), sanitizeCellForExcel(op.circulatingNurse||""),
        sanitizeCellForExcel(op.anesthesiaNurse||""), sanitizeCellForExcel(op.onloopNurse||""),
        sanitizeCellForExcel(op.rrKatim||""), sanitizeCellForExcel(op.allergy||""), op.bloodType||"",
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

      <StatistikBulanan ops={ops} archive={archive}/>

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



      {/* ── Rekapitulasi Tindakan ── */}
      {procedureData.length>0 && (
        <Card>
          <SH label="🔪 Rekapitulasi Tindakan Medis" color="#0284c7"/>
          <div style={{fontSize:11,color:C.tL,marginBottom:10}}>Total {procedureData.length} jenis tindakan · {filtered.length} operasi</div>
          {procedureData.map(([label,val],i)=>{
            const pct = filtered.length>0 ? Math.round((val/filtered.length)*100) : 0;
            const barColor = COLORS[i%COLORS.length];
            return (
              <div key={label} style={{marginBottom:8}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:3}}>
                  <div style={{fontSize:12,fontWeight:600,color:C.t,flex:1,marginRight:8,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>
                    <span style={{color:C.tL,marginRight:6,fontSize:11}}>#{i+1}</span>{label}
                  </div>
                  <div style={{display:"flex",alignItems:"center",gap:8,flexShrink:0}}>
                    <span style={{fontSize:13,fontWeight:800,color:barColor}}>{val}</span>
                    <span style={{fontSize:10,color:C.tL}}>({pct}%)</span>
                  </div>
                </div>
                <div style={{background:"#f1f5f9",borderRadius:6,height:8,overflow:"hidden"}}>
                  <div style={{height:"100%",background:barColor,borderRadius:6,width:`${pct}%`,transition:"width .4s"}}/>
                </div>
              </div>
            );
          })}
        </Card>
      )}

      {/* ── Rekapitulasi Diagnosa ── */}
      {diagnosisData.length>0 && (
        <Card>
          <SH label="🏥 Rekapitulasi Diagnosa" color="#7c3aed"/>
          <div style={{fontSize:11,color:C.tL,marginBottom:10}}>Total {diagnosisData.length} jenis diagnosa</div>
          {diagnosisData.map(([label,val],i)=>{
            const pct = filtered.length>0 ? Math.round((val/filtered.length)*100) : 0;
            const barColor = ["#7c3aed","#a855f7","#8b5cf6","#6d28d9","#9333ea","#4f46e5","#7c3aed","#c026d3"][i%8];
            return (
              <div key={label} style={{marginBottom:8}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:3}}>
                  <div style={{fontSize:12,fontWeight:600,color:C.t,flex:1,marginRight:8,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>
                    <span style={{color:C.tL,marginRight:6,fontSize:11}}>#{i+1}</span>{label}
                  </div>
                  <div style={{display:"flex",alignItems:"center",gap:8,flexShrink:0}}>
                    <span style={{fontSize:13,fontWeight:800,color:barColor}}>{val}</span>
                    <span style={{fontSize:10,color:C.tL}}>({pct}%)</span>
                  </div>
                </div>
                <div style={{background:"#f5f3ff",borderRadius:6,height:8,overflow:"hidden"}}>
                  <div style={{height:"100%",background:barColor,borderRadius:6,width:`${pct}%`,transition:"width .4s"}}/>
                </div>
              </div>
            );
          })}
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
