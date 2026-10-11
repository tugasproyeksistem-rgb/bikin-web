/* src/components/import-excel.tsx
   Komponen import jadwal operasi massal dari file Excel (.xlsx/.xls) —
   drag-drop atau klik, preview 10 baris pertama sebelum konfirmasi import.
   Dipindah dari App.tsx (langkah 6 refactor — pisah View/layar besar,
   batch 5), TIDAK ADA perubahan logika, termasuk FIX AUDIT #28. */

import { useState, useRef } from "react";
import * as XLSX from "xlsx";
import { C, ROOMS } from "../constants";
import { toISODateStrict, gId, fNow, todayDate, fDMY } from "../utils";
import { Card, SH, Btn } from "./primitives";

export function ImportExcel({ops, setOps, showToast, upsertBulkToSupa}: any) {
  const [rows, setRows] = useState<any[]>([]);
  const [fileName, setFileName] = useState("");
  const [importing, setImporting] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);

  const COLS: {label:string;key:string;required?:boolean}[] = [
    {label:"Tanggal (DD/MM/YYYY atau YYYY-MM-DD)",key:"date",required:true},
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
    /* FIX AUDIT #28 (LOW — FileReader tanpa onerror): tanpa ini, kalau file
       gagal dibaca di level OS/browser (bukan gagal parse Excel), tidak ada
       feedback ke user sama sekali — UI terlihat "menggantung" tanpa
       penjelasan. */
    reader.onerror = () => { setFileName(""); showToast("Gagal membaca file — coba pilih file lain", C.d); };
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target!.result as ArrayBuffer), {type:"array", cellDates:true});
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
            if(idx<0){ obj[col.key]=""; return; }
            /* Kolom tanggal disanitasi khusus — sel Excel bisa berupa Date object
               atau serial number, jangan langsung String() atau formatnya rusak. */
            obj[col.key] = col.key==="date" ? (toISODateStrict(row[idx])||"") : String(row[idx]??"").trim();
          });
          if(obj.date||obj.patient||obj.procedure) parsed.push(obj);
        }
        setRows(parsed);
        if(parsed.length===0) showToast("Tidak ada baris data ditemukan",C.w);
      } catch(err){showToast("Gagal membaca file Excel",C.d);}
    };
    reader.readAsArrayBuffer(file);
  };

  const doImport = async () => {
    setImporting(true);
    const valid = rows.filter(r=>r.patient&&r.date&&r.procedure);
    const ts = new Date().toISOString();
    const newOps = valid.map(r=>{
      return {
        id: gId(), status:"scheduled", reminders:[], requests:[],
        createdAt: fNow(), updated_at: ts,
        opType: r.opType||"elektif",
        date: r.date, time: r.time||"08:00",
        patient: r.patient, diagnosis: r.diagnosis||"", procedure: r.procedure,
        surgeon: r.surgeon||"", anesthesiologist: r.anesthesiologist||"",
        room: r.room||ROOMS[0], age: r.age||"", rm: r.rm||"",
        bloodType: r.bloodType||"Tidak Diketahui", allergy: r.allergy||"Tidak Ada",
        circulatingNurse: r.circulatingNurse||"", anesthesiaNurse: r.anesthesiaNurse||"",
        onloopNurse: r.onloopNurse||"", rrKatim: r.rrKatim||"",
      };
    });
    /* ── SSOT Import Flow ──────────────────────────────────────────────────
       DB-FIRST: tulis ke Supabase dulu, baru update local state setelah konfirmasi.
       Ini memastikan data persist bahkan jika browser di-refresh sesaat setelah import.
       postgres_changes akan debounce ke device lain secara otomatis.
       ──────────────────────────────────────────────────────────────────── */
    showToast(`⟳ Menyimpan ${newOps.length} jadwal ke Supabase...`,C.p);
    const res = upsertBulkToSupa
      ? await upsertBulkToSupa("kb_operasi", newOps)
      : {ok:false, error:"upsertBulkToSupa tidak tersedia"};
    if(res.ok){
      /* Insert ke local state HANYA setelah DB sukses — cegah ghost data */
      setOps((p:any[])=>{
        const existingIds = new Set(p.map((o:any)=>o.id));
        const toAdd = newOps.filter(o=>!existingIds.has(o.id));
        return [...p,...toAdd];
      });
      showToast(`✓ ${newOps.length} jadwal berhasil diimport & tersinkron`,C.s);
    } else {
      showToast(`⚠ Import gagal: ${res.error||"kesalahan tidak diketahui dari Supabase"}`,C.d);
    }
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
                <tbody>{rows.slice(0,10).map((r,i)=><tr key={i} style={{borderBottom:`1px solid ${C.b}`}}><td style={{padding:"5px 10px"}}>{fDMY(r.date)}</td><td style={{padding:"5px 10px"}}>{r.time}</td><td style={{padding:"5px 10px",fontWeight:600}}>{r.patient}</td><td style={{padding:"5px 10px"}}>{r.procedure}</td><td style={{padding:"5px 10px"}}>{r.room}</td><td style={{padding:"5px 10px"}}>{r.surgeon}</td></tr>)}</tbody>
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
