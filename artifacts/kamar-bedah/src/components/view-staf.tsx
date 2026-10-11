/* src/components/view-staf.tsx
   Tab "Staf" — CRUD data staf (dokter/perawat) + upload data master staf
   dan jadwal siaga bulanan via CSV/Excel. Dipindah dari App.tsx (langkah 6
   refactor — pisah View/layar besar, batch 5), TIDAK ADA perubahan logika,
   termasuk seluruh FIX AUDIT di dalamnya. */

import { useState, useRef } from "react";
import * as XLSX from "xlsx";
import type { ViewStafProps } from "../types";
import { C, ST, iS } from "../constants";
import { downloadBlob, parseCSV, sanitize, normalizePhone62, isValidPhone62, normalizeStaffType, toISODateStrict, gId, todayDate, tmrwDate, fD } from "../utils";
import { AmbilLogMedis } from "../services/supabase";
import { Row, Card, SH, Btn, LF, SF } from "./primitives";
import { SquaresBackground } from "./effects";

export function ViewStaf({staff,setStaff,roster,setRoster,showToast,upsertOneToSupa,deleteFromSupa,upsertBulkToSupa}: ViewStafProps) {
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
    else if(!isValidPhone62(form.phone.replace(/\s/g,""))) e.phone="Format: 62xxx (tanpa + atau spasi)";
    setErr(e); return !Object.keys(e).length;
  };
  const save = () => {
    if(!validate()) return;
    if(editing){
      const updated={...editing,...form, updated_at: new Date().toISOString()};
      setStaff((p:any[])=>p.map((s:any)=>s.id===editing.id?updated:s));
      AmbilLogMedis("MANAJEMEN_STAF","Admin","UBAH_STAF",
        `Edit data staf: ${updated.name||"-"} (${updated.type||"-"})`,
        { before: editing, after: updated }
      );
      upsertOneToSupa?.("kb_staf",updated).then((res:{ok:boolean;error?:string})=>{
        if(!res?.ok){
          setStaff((p:any[])=>p.map((s:any)=>s.id===editing.id?editing:s));
          showToast(`⚠ Gagal menyimpan staf: ${res?.error||"kesalahan tidak diketahui"}`,C.d);
        } else showToast("✓ Data staf diperbarui & tersinkron",C.s);
      });
    } else {
      const newStaf={id:gId(),...form, updated_at: new Date().toISOString()};
      AmbilLogMedis("MANAJEMEN_STAF","Admin","TAMBAH_STAF",
        `Tambah staf baru: ${newStaf.name||"-"} (${newStaf.type||"-"})`,
        { ...newStaf }
      );
      upsertOneToSupa?.("kb_staf",newStaf).then((res:{ok:boolean;error?:string})=>{
        if(!res?.ok){
          showToast(`⚠ Gagal menyimpan staf: ${res?.error||"kesalahan tidak diketahui"}`,C.d);
        } else {
          setStaff((p:any[])=>{
            if(p.some((s:any)=>s.id===newStaf.id)) return p;
            return [...p,newStaf];
          });
          showToast("✓ Staf berhasil ditambahkan & tersinkron",C.s);
        }
      });
    }
    setForm({name:"",type:"surgeon",phone:""}); setEd(null); setShow(false); setErr({});
  };
  const dlStaff = () => {
    const csv = `"nama","jabatan","nomor_wa"\n"dr. Nama Bedah Sp.B","surgeon","628112345001"\n"dr. Nama Anestesi Sp.An","anesthesiologist","628112345002"\n"Ns. Nama Instrumen","circulating","628112345003"\n"Ns. Nama P.Anestesi","anesthesia_nurse","628112345004"\n"Ns. Nama Onloop","onloop","628112345005"\n"Ns. Nama Katim","katim","628112345006"`;
    downloadBlob(csv,"template_staf.csv"); showToast("✓ Template CSV Staf diunduh",C.s);
  };
  const dlRoster = () => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
      ["Tanggal","Dokter Anestesi Jaga 1","Dokter Anestesi Jaga 2","Dokter Anestesi Cyto 1","Dokter Anestesi Cyto 2","Dokter Anestesi Cyto 3","Perawat Siaga 1","Perawat Siaga 2","Perawat Siaga 3","Perawat Siaga 4","Pembawa HP"],
      ["01/06/2025","dr. Anestesi Jaga 1 Sp.An","dr. Anestesi Jaga 2 Sp.An","dr. Cyto 1 Sp.An","dr. Cyto 2 Sp.An","dr. Cyto 3 Sp.An","Ns. Perawat 1","Ns. Perawat 2","Ns. Perawat 3","Ns. Perawat 4","Ns. Katim Jaga"],
      ["02/06/2025","dr. Anestesi Jaga 1 Sp.An","dr. Anestesi Jaga 2 Sp.An","dr. Cyto 1 Sp.An","dr. Cyto 2 Sp.An","dr. Cyto 3 Sp.An","Ns. Perawat 1","Ns. Perawat 2","Ns. Perawat 3","Ns. Perawat 4","Ns. Katim Jaga"],
    ]);
    ws["!cols"]=[{wch:12},{wch:28},{wch:28},{wch:28},{wch:28},{wch:28},{wch:22},{wch:22},{wch:22},{wch:22},{wch:22}];
    XLSX.utils.book_append_sheet(wb,ws,"Jadwal Siaga");
    XLSX.writeFile(wb,"template_jadwal_siaga_bulanan.xlsx");
    showToast("✓ Template Jadwal Siaga diunduh",C.s);
  };

  /* readFileRows: PENTING — untuk .xlsx dibaca dengan cellDates:true dan TIDAK
     dipaksa jadi string di sini. Memaksa String(cell) terlalu awal adalah
     sumber bug nyata: sel tanggal Excel yang diformat sebagai tanggal akan
     terbaca sebagai angka "serial" (mis. 45809) oleh SheetJS bila cellDates
     tidak diaktifkan, lalu String(45809) tidak akan pernah cocok dengan
     regex DD/MM/YYYY — tanggal rusak inilah yang akhirnya ditolak Supabase.
     Sel Date/number asli dipertahankan agar bisa disanitasi per-kolom oleh
     toISODateStrict() di pemanggil. */
  const readFileRows = async (f: File): Promise<any[][]> => {
    const isXlsx = f.name.match(/\.xlsx?$/i);
    if(isXlsx) {
      const buf = await f.arrayBuffer();
      const wb2 = XLSX.read(new Uint8Array(buf),{type:"array", cellDates:true});
      const ws2 = wb2.Sheets[wb2.SheetNames[0]];
      const raw: any[][] = XLSX.utils.sheet_to_json(ws2,{header:1,defval:""});
      return raw;
    }
    return parseCSV(await f.text());
  };

  const handleStaffFile = async (ev: any) => {
    const f=ev.target.files[0]; if(!f)return; ev.target.value="";
    try {
      const rows = await readFileRows(f);
      if(rows.length<2){showToast("File kosong atau format tidak sesuai",C.d);return;}
      const h=rows[0].map((x:any)=>String(x).toLowerCase().trim());
      const ni=h.findIndex((x:string)=>x.includes("nama")), ti=h.findIndex((x:string)=>x.includes("jabatan")||x.includes("type")), pi=h.findIndex((x:string)=>x.includes("wa")||x.includes("phone")||x.includes("nomor"));
      if(ni<0){showToast("Kolom 'nama' tidak ditemukan. Gunakan template.",C.d);return;}
      let skippedPhone = 0;
      const imp=rows.slice(1).map((c:any[])=>{
        const name = sanitize(String(c[ni]??""));
        if(!name) return null;
        const rawPhone = String(pi>=0?c[pi]??"":"");
        const phone = normalizePhone62(rawPhone);
        if(rawPhone.trim() && !isValidPhone62(phone)) skippedPhone++;
        /* Cocokkan dgn staf yang SUDAH ADA (nama, case-insensitive) — jika ketemu,
           PERBARUI baris yg sama (reuse id) bukan membuat staf duplikat baru.
           Tanpa ini, re-upload file staf yg sama akan menggandakan data setiap kali. */
        const existing = staff.find((s:any)=>String(s.name||"").toLowerCase().trim()===name.toLowerCase());
        return {
          id: existing ? existing.id : gId(),
          name,
          type: ti>=0 ? normalizeStaffType(String(c[ti]??"")) : "circulating",
          phone,
          updated_at: new Date().toISOString(),
        };
      }).filter(Boolean) as any[];
      if(!imp.length){showToast("Tidak ada data valid (kolom nama kosong)",C.d);return;}
      showToast(`⟳ Menyimpan ${imp.length} staf ke Supabase...`,C.p);
      const res = await upsertBulkToSupa?.("kb_staf", imp);
      if(res?.ok) {
        setStaff((p:any[])=>{
          const byId = new Map(p.map((s:any)=>[s.id,s]));
          imp.forEach((s:any)=>byId.set(s.id,s));
          return Array.from(byId.values());
        });
        const note = skippedPhone ? ` (${skippedPhone} nomor WA format tidak valid — tetap tersimpan, cek kembali)` : "";
        showToast("✓ "+imp.length+" staf berhasil diimport & tersinkron"+note,C.s);
      } else {
        showToast(`⚠ Import staf gagal: ${res?.error||"kesalahan tidak diketahui dari Supabase"}`,C.d);
      }
    } catch(e: any){showToast("Error membaca file: "+(e?.message||"format tidak dikenali"),C.d);}
  };
  const handleRosterFile = async (ev: any) => {
    const f=ev.target.files[0]; if(!f)return; ev.target.value="";
    try {
      const rows = await readFileRows(f);
      if(rows.length<2){showToast("File kosong",C.d);return;}
      const h=rows[0].map((x:any)=>String(x).toLowerCase().trim());
      const di=h.findIndex((x:string)=>x.includes("tanggal")||x.includes("date"));
      if(di<0){showToast("Kolom 'Tanggal' tidak ditemukan. Gunakan template.",C.d);return;}
      const ai1=h.findIndex((x:string)=>x.includes("jaga")&&!x.includes("cyto")&&(x.includes("1")||(!x.includes("2")&&!x.includes("3"))));
      const ai2=h.findIndex((x:string)=>x.includes("jaga")&&!x.includes("cyto")&&x.includes("2"));
      const ci1=h.findIndex((x:string)=>x.includes("cyto")&&(x.includes("1")||(!x.includes("2")&&!x.includes("3"))));
      const ci2=h.findIndex((x:string)=>x.includes("cyto")&&x.includes("2"));
      const ci3=h.findIndex((x:string)=>x.includes("cyto")&&x.includes("3"));
      const pi=h.findIndex((x:string)=>x.includes("perawat")||x.includes("nurse"));
      const hpi=h.findIndex((x:string)=>x.includes("pembawa")||x.includes("hp"));
      const ts = new Date().toISOString();
      let skippedDate = 0;
      const imp=rows.slice(1).map((r:any[])=>{
        /* toISODateStrict menangani Date object, serial number Excel, dan
           string DD/MM/YYYY atau YYYY-MM-DD sekaligus — lihat definisinya. */
        const date = toISODateStrict(r[di]);
        if(!date){ if(r[di]) skippedDate++; return null; }
        const nurses: string[]=[];
        if(pi>=0) for(let i=pi;i<Math.min(r.length,pi+4);i++){ const cell=String(r[i]??"").trim(); if(cell && !cell.toLowerCase().includes("pembawa")) nurses.push(sanitize(cell)); }
        const jaga1=sanitize(String(r[ai1>=0?ai1:0]??""));
        const jaga2=ai2>=0?sanitize(String(r[ai2]??"")):"";
        const cyto1=ci1>=0?sanitize(String(r[ci1]??"")):"";
        const cyto2=ci2>=0?sanitize(String(r[ci2]??"")):"";
        const cyto3=ci3>=0?sanitize(String(r[ci3]??"")):"";
        const anestJagaArr=[jaga1,jaga2].filter(Boolean);
        const anestCytoArr=[cyto1,cyto2,cyto3].filter(Boolean);
        /* id DETERMINISTIK dari tanggal (bukan random) — re-upload jadwal yang
           tanggalnya sama akan UPSERT baris yg sama (onConflict:"id"), bukan
           membuat baris baru yang "berhantu" (tetap ada di DB walau sudah
           tergeser di tampilan lokal) dan muncul lagi sbg duplikat di device
           lain / setelah refresh. */
        return {id:`roster_${date}`,date,anestJaga:anestJagaArr.join(", "),anestCyto:anestCytoArr.join(", "),anestJagaList:anestJagaArr,anestCytoList:anestCytoArr,nurses,pembawaHP:sanitize(String(r[hpi>=0?hpi:0]??"")),updated_at:ts};
      }).filter(Boolean) as any[];
      if(!imp.length){showToast(`Tidak ada baris valid (${skippedDate} tanggal tidak terbaca — cek format DD/MM/YYYY)`,C.d);return;}
      /* Merge lokal: timpa entry dengan tanggal yang sama */
      const oldSameDate = roster.filter((x:any)=>imp.some((n:any)=>n.date===x.date));
      const newRoster=[...roster.filter((x:any)=>!imp.find((n:any)=>n.date===x.date)),...imp];
      showToast(`⟳ Menyimpan ${imp.length} jadwal siaga ke Supabase...`,C.p);
      /* DB-first: roster harus persist sebelum UI diupdate */
      const res = await upsertBulkToSupa?.("kb_roster", imp);
      if(res?.ok) {
        setRoster(newRoster);
        /* Best-effort: bersihkan baris lama (id berbeda, mis. dari upload versi
           sebelumnya yg masih pakai id acak) utk tanggal yg sama agar tidak
           menumpuk sbg duplikat "hantu" di Supabase. */
        const orphanIds = oldSameDate.filter((x:any)=>!imp.some((n:any)=>n.id===x.id)).map((x:any)=>x.id);
        if(orphanIds.length) Promise.all(orphanIds.map((id:string)=>deleteFromSupa?.("kb_roster", id))).catch(()=>{});
        const note = skippedDate ? ` (${skippedDate} baris dilewati — tanggal tidak terbaca)` : "";
        showToast("✓ "+imp.length+" jadwal siaga tersimpan & tersinkron ke semua perangkat"+note,C.s);
      } else {
        showToast(`⚠ Upload jadwal siaga gagal: ${res?.error||"kesalahan tidak diketahui dari Supabase"}`,C.d);
      }
    } catch(e: any){showToast("Error membaca file: "+(e?.message||"format tidak dikenali"),C.d);}
  };

  return (
    <div style={{position:"relative"}}>
      <SquaresBackground opacity={0.4}/>
      <div style={{position:"relative",zIndex:1}}>
      <Row title="Manajemen Staf & Roster" right={<Btn sm onClick={()=>{setForm({name:"",type:"surgeon",phone:""});setEd(null);setErr({});setShow(true);}}>+ Tambah Staf</Btn>}/>
      <Card style={{background:C.iBg,border:"1px solid #1565C033"}}>
        <SH label="📋 Upload Data Master Staf" color={C.i}/>
        <div style={{fontSize:12,color:C.tL,marginBottom:10,lineHeight:1.6}}>Format: <b>CSV</b> atau <b>Excel (.xlsx)</b> · Kolom: nama, jabatan, nomor_wa</div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <input ref={sfRef} type="file" accept=".csv,.txt,.xlsx,.xls" onChange={handleStaffFile} style={{display:"none"}}/>
          <Btn sm outline color={C.i} onClick={()=>sfRef.current&&sfRef.current.click()}>📂 Upload CSV/Excel Staf</Btn>
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
          <input ref={rsfRef} type="file" accept=".csv,.txt,.xlsx,.xls" onChange={handleRosterFile} style={{display:"none"}}/>
          <Btn sm color="#7B1FA2" onClick={()=>rsfRef.current&&rsfRef.current.click()}>📂 Upload Jadwal Siaga (CSV/Excel)</Btn>
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
                    <Btn sm outline color={C.d} onClick={async()=>{
                      const res = await deleteFromSupa?.("kb_staf",s.id);
                      if(!res?.ok){ showToast(`⚠ Gagal menghapus staf: ${res?.error||"kesalahan tidak diketahui"}`,C.d); return; }
                      setStaff((p:any[])=>p.filter((x:any)=>x.id!==s.id));
                      AmbilLogMedis("MANAJEMEN_STAF","Admin","HAPUS_STAF",
                        `Hapus staf: ${s.name||"-"} (${s.type||"-"})`,
                        { deleted: { id: s.id, name: s.name, type: s.type } }
                      );
                      showToast("✓ Staf dihapus & tersinkron",C.d);
                    }}>Hapus</Btn>
                  </div>
                </div>
              </Card>
            ))}
          </div>
        );
      })}
    </div>
    </div>
  );
}
