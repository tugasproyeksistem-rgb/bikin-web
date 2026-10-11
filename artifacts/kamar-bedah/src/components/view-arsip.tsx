/* src/components/view-arsip.tsx
   Tab "Arsip" — log notifikasi, arsip harian, download data per bulan
   (lokal/Supabase/Dropbox), import Excel, audit trail, dan pengaturan
   backup (Supabase, Dropbox, Realtime, Privacy mode). Dipindah dari
   App.tsx (langkah 6 refactor — pisah View/layar besar, batch 6),
   TIDAK ADA perubahan logika, termasuk seluruh FIX AUDIT di dalamnya. */

import { useState } from "react";
import * as XLSX from "xlsx";
import type { ViewArsipProps, SupabaseConfig, DropboxConfig } from "../types";
import { C, OT, STS, iS, PAGE_SIZE, LOCK_MS, HOSPITAL, MON_ROOMS } from "../constants";
import { todayDate, fD, fNow, sanitizeCellForExcel, sanitizeForCSV, downloadBlob } from "../utils";
import { downloadAsWord } from "../utils/word-export";
import { monIsOK } from "../utils/monitoring-helpers";
import { SUPA_CLIENT } from "../services/supabase";
import { callDbxProxy } from "../services/dropbox";
import { Card, SH, Btn, LF, Bdg, Toggle } from "./primitives";
import { ImportExcel } from "./import-excel";

export function ViewArsip(props: ViewArsipProps) {
  const { ops, archive, notifs, lemburData, lemburPegawai, monitoringEntries, monitoringCfg } = props.data;
  const { setOps, showToast, role, upsertBulkToSupa, auditLog } = props;
  const { mode: privacyMode, setMode: setPrivacyMode } = props.privacy;
  const { cfg: supaCfg, setCfg: setSupaCfg, status: supaStatus, backingUp: supaBackingUp,
          onBackup: onSupaBackup, onRestoreOps: onSupaRestoreOps, onRestoreLembur: onSupaRestoreLembur,
          onRestoreMonitoring: onSupaRestoreMonitoring, onRestoreAll: onSupaRestoreAll } = props.supabase;
  const { cfg: dbxCfg, setCfg: setDbxCfg, status: dbxStatus, backingUp: dbxBacking,
          onBackup: onDbxBackup, onRestoreOps: onDbxRestoreOps, onRestoreLembur: onDbxRestoreLembur,
          onBackupOpsXls: onDbxBackupOpsXls, onBackupLemburXls: onDbxBackupLemburXls,
          onBackupMonitoringXls: onDbxBackupMonitoringXls } = props.dropbox;
  const { status: rtStatus, enabled: rtEnabled, setEnabled: setRtEnabled } = props.realtime;
  const [sub,setSub] = useState("notif");
  const [arsipTahun,setArsipTahun]  = useState("");
  const [arsipBulan,setArsipBulan]  = useState("");
  const [arsipTipe,setArsipTipe]    = useState("laporan_operasi");
  const [dlCloudMonth,setDlCloudMonth] = useState(()=>{const d=new Date();return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0");});
  const [dlCloudType,setDlCloudType] = useState<"ops"|"lembur"|"monitoring">("ops");
  const [dlCloudBusy,setDlCloudBusy] = useState(false);
  const nCfg: any = {cyto:{c:C.d,bg:C.dBg},"H-1":{c:C.i,bg:C.iBg},"H-1jam":{c:C.w,bg:C.wBg},wa_direct:{c:C.wa,bg:C.waBg}};

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
      const h = ["Ruangan","Tanggal","Jam","Suhu (°C)","Kelembaban (%)","Status","Petugas"];
      const rows = me
        .sort((a:any,b:any)=>
          (MON_ROOMS.indexOf(a.ruang??"")-MON_ROOMS.indexOf(b.ruang??"")) ||
          a.tanggal.localeCompare(b.tanggal) ||
          a.jam.localeCompare(b.jam)
        )
        .map((e:any)=>[
          e.ruang??monitoringCfg?.lokasiRuang??"",
          e.tanggal, e.jam, e.suhu, e.kelembaban,
          monIsOK(e.suhu,e.kelembaban,monitoringCfg)?"SESUAI":"TIDAK SESUAI",
          e.petugas
        ]);
      const ws = XLSX.utils.aoa_to_sheet([h,...rows]);
      ws["!cols"]=[{wch:18},{wch:12},{wch:8},{wch:12},{wch:14},{wch:16},{wch:22}];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Monitoring Suhu");
      XLSX.writeFile(wb, `Monitoring_Suhu${arsipTahun?"_"+arsipTahun:""}${arsipBulan?"_"+arsipBulan:""}_${todayDate()}.xlsx`);
      showToast("✓ Excel monitoring suhu diunduh",C.s);
    } else {
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
          (rec.entries||[]).forEach((e:any)=>rows.push([sanitizeCellForExcel(p.name),p.nik||"",k.replace(p.id+"_",""),e.tanggalAwal||"",e.tanggalAkhir||"",e.jamMasuk||"",e.jamKeluar||"",sanitizeCellForExcel(e.keperluanLembur||""),sanitizeCellForExcel(e.keterangan||"")]));
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

  const ymLabel = (ym:string) => {
    if(!ym) return "";
    const ML: Record<string,string> = {"01":"Jan","02":"Feb","03":"Mar","04":"Apr","05":"Mei","06":"Jun","07":"Jul","08":"Agu","09":"Sep","10":"Okt","11":"Nov","12":"Des"};
    const [y,m] = ym.split("-");
    return `${ML[m]||m} ${y}`;
  };

  const buildAndDownloadXlsx = (title:string, header:string[], rows:(string|number)[][], filename:string) => {
    const ws = XLSX.utils.aoa_to_sheet([header,...rows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, title.slice(0,31));
    XLSX.writeFile(wb, filename);
  };

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
      const h=["Ruangan","Tanggal","Jam","Suhu (°C)","Kelembaban (%)","Status","Petugas"];
      const rows=filtered.map((e:any)=>[e.ruang||"",e.tanggal||"",e.jam||"",e.suhu??"",e.kelembaban??"",monIsOK(e.suhu,e.kelembaban,monitoringCfg)?"SESUAI":"TIDAK SESUAI",e.petugas||""]);
      buildAndDownloadXlsx("Monitoring "+ymLabel(month),h,rows,`Monitoring_${month}.xlsx`);
      showToast(`✓ Excel Monitoring ${ymLabel(month)} diunduh`,C.s);
    }
  };

  const dlFromSupabase = async (month:string, type:"ops"|"lembur"|"monitoring") => {
    if(!month){showToast("Pilih bulan terlebih dahulu",C.w);return;}
    setDlCloudBusy(true);
    try {
      const tbl = "kamar_bedah_backup";
      const {data, error} = await SUPA_CLIENT.from(tbl).select("*").order("created_at",{ascending:false}).limit(1);
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

  const dlFromDropbox = async (month:string, type:"ops"|"lembur"|"monitoring") => {
    if(!month){showToast("Pilih bulan terlebih dahulu",C.w);return;}
    setDlCloudBusy(true);
    try {
      const res = await callDbxProxy({action:"restore", path: dbxCfg.path});
      if(!res.ok) throw new Error(res.msg);
      const raw = res.data;
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
            const rows=ops.map((op:any,i:number)=>[i+1,op.date,op.time,sanitizeForCSV(op.patient),sanitizeForCSV(op.age||""),sanitizeForCSV(op.rm||""),OT[op.opType as keyof typeof OT||"elektif"]&&OT[op.opType as keyof typeof OT||"elektif"].label||"Elektif",sanitizeForCSV(op.diagnosis),sanitizeForCSV(op.procedure),op.room,op.surgeon,op.anesthesiologist||"",op.assistantNurse||"",op.circulatingNurse||"",op.anesthesiaNurse||"",op.onloopNurse||"",op.rrKatim||"",sanitizeForCSV(op.allergy||""),op.bloodType,op.status]);
            downloadBlob([h,...rows].map((r:any[])=>r.map((c:any)=>typeof c==="string"&&c.startsWith('"')?c:sanitizeForCSV(c)).join(",")).join("\n"),"JadwalOK_PantiRini_"+todayDate()+".csv");
            showToast("✓ CSV berhasil diunduh",C.s);
          }} style={{marginBottom:12}}>⬇ Unduh CSV Jadwal</Btn>
          {ops.length===0 && <Card><div style={{textAlign:"center",padding:20,color:C.tL,fontSize:13}}>Belum ada data jadwal</div></Card>}
          {ops.map((op:any)=>{
            const sc=STS[op.status as keyof typeof STS]||STS.scheduled, ot=OT[op.opType as keyof typeof OT||"elektif"];
            return (
              <Card key={op.id} style={{padding:"12px 16px",marginBottom:6, ...(op.needsReview?{border:"2px solid #3B82F6",background:"#EFF6FF"}:{})}}>
                <div style={{fontSize:13,fontWeight:700,color:C.t}}>{op.patient||"—"}</div>
                <div style={{fontSize:12,color:C.tL}}>{op.procedure} · {fD(op.date)} {op.time} WIB</div>
                <div style={{marginTop:6,display:"flex",gap:5,flexWrap:"wrap"}}>
                  <Bdg label={sc.l} color={sc.c} bg={sc.bg} icon={sc.icon}/>
                  <Bdg label={ot&&ot.label||"Elektif"} color={ot&&ot.c||C.i} bg={ot&&ot.bg||C.iBg}/>
                  {op.allergy&&op.allergy!=="Tidak Ada"&&<Bdg label="Alergi" color={C.d} bg={C.dBg}/>}
                  {op.needsReview&&<Bdg label="🔍 Perlu Review (dari OCR)" color="#3B82F6" bg="#DBEAFE"/>}
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
          <Card style={{background:"#EDE7F6",border:"1px solid #B39DDB",marginBottom:14}}>
            <SH label="📥 Download Excel per Bulan" color="#4527A0"/>
            <div style={{fontSize:12,color:"#5E35B1",marginBottom:12,lineHeight:1.7}}>
              Pilih bulan &amp; jenis data, lalu download langsung sebagai file Excel (.xlsx).
              Bisa dari data <b>Lokal</b> (perangkat ini), <b>Supabase</b> (cloud), atau <b>Dropbox</b> (cloud).
            </div>
            <div style={{background:"#F3E5F5",borderRadius:8,padding:"8px 12px",marginBottom:12,fontSize:12,color:"#6A1B9A",lineHeight:1.6}}>
              ✅ <b>Multi-Device Sync Aktif:</b> Data tersinkron otomatis ke <b>Supabase</b> & <b>Dropbox</b> setiap ada perubahan.
              Buka dari HP/tablet lain = data langsung tersedia. Auto-delete backup &gt; 2 tahun.
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
              {dbxCfg?.path&&(
                <Btn full onClick={()=>dlFromDropbox(dlCloudMonth,dlCloudType)} style={{flex:1,background:dlCloudBusy?"#aaa":"#0061FF",color:"#fff",border:"none",minWidth:120,opacity:dlCloudBusy?.6:1,cursor:dlCloudBusy?"not-allowed":"pointer"}}>
                  {dlCloudBusy?"⏳ Mengunduh...":"📦 Dropbox"}
                </Btn>
              )}
            </div>
            {!supaCfg?.url&&!dbxCfg?.path&&(
              <div style={{marginTop:10,fontSize:11,color:"#7B1FA2",background:"#F3E5F5",borderRadius:8,padding:"6px 10px"}}>
                💡 Tombol Supabase &amp; Dropbox akan muncul setelah konfigurasi diisi di tab ⚙ Setelan.
              </div>
            )}
          </Card>

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
            <SH label="📋 Download File Tracking Lembur" color="#1F5A52"/>
            <div style={{fontSize:12,color:"#1F5A52",marginBottom:12,lineHeight:1.6}}>
              Unduh semua data lembur dari semua pegawai sebagai file tracking (Excel atau Word).
              Total pegawai: <b>{(lemburPegawai||[]).length}</b>.
            </div>
            {role==="admin" && <div style={{display:"flex",gap:8}}>
              <Btn full onClick={()=>downloadTrackingLembur("excel")} style={{flex:1,background:"#1F5A52",color:"#fff",border:"none"}}>
                📊 Excel Tracking Lembur
              </Btn>
              <Btn full onClick={()=>downloadTrackingLembur("word")} style={{flex:1,background:"#16685F",color:"#fff",border:"none"}}>
                📄 Word Tracking Lembur
              </Btn>
            </div>}
          </Card>
        </div>
      )}

      {sub==="import" && (
        <ImportExcel ops={ops} setOps={setOps} showToast={showToast} upsertBulkToSupa={upsertBulkToSupa}/>
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
          <Card>
            <SH label="Mode Privasi" color="#4A148C"/>
            <Toggle value={privacyMode} onChange={(v:boolean)=>{setPrivacyMode(v);showToast(v?"🔒 Mode Privasi aktif":"Mode Privasi dimatikan",v?"#7B1FA2":C.g);}} color="#4A148C" label="🔒 Sembunyikan Data Pasien" sub="Nama & RM disamarkan di layar (pesan WA tetap asli)"/>
            {privacyMode && <div style={{background:"#F3E5F5",borderRadius:8,padding:"8px 12px",marginTop:10,fontSize:12,color:"#7B1FA2",lineHeight:1.6}}>Mode aktif. Aman untuk presentasi, audit & dokumentasi sistem.</div>}
          </Card>

          <Card>
            <SH label="☁ Backup ke Supabase" color="#3ECF8E"/>
            <div style={{background:"#ECFDF5",borderRadius:8,padding:"10px 12px",marginBottom:14,fontSize:12,color:"#065F46",lineHeight:1.7,border:"1px solid #A7F3D0"}}>
              <b>✅ Koneksi Supabase aktif otomatis.</b><br/>
              Project: <code style={{background:"#D1FAE5",padding:"1px 5px",borderRadius:4}}>ezwgnpdtzcabxmimovbp.supabase.co</code><br/>
              Konfigurasi dikelola secara terpusat — tidak perlu input manual.
            </div>
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
            <div style={{background:"#F0F4F8",borderRadius:10,padding:"10px 14px",marginBottom:10,display:"flex",gap:12,flexWrap:"wrap"}}>
              <div style={{flex:1,minWidth:120}}>
                <div style={{fontSize:10,fontWeight:700,color:C.p,marginBottom:4,textTransform:"uppercase",letterSpacing:.5}}>☁ Supabase</div>
                <div style={{fontSize:11,color:C.g}}>Free: ~500MB · Auto-backup aktif</div>
                <div style={{background:C.b,borderRadius:4,height:6,marginTop:4}}>
                  <div style={{background:C.p,height:6,borderRadius:4,width:"2%"}}/>
                </div>
                <div style={{fontSize:10,color:C.tL,marginTop:2}}>~2% terpakai · Auto-delete &gt;2 tahun</div>
              </div>
              <div style={{flex:1,minWidth:120}}>
                <div style={{fontSize:10,fontWeight:700,color:"#0061FF",marginBottom:4,textTransform:"uppercase",letterSpacing:.5}}>📦 Dropbox</div>
                <div style={{fontSize:11,color:C.g}}>Free: ~2GB · Auto-backup aktif</div>
                <div style={{background:C.b,borderRadius:4,height:6,marginTop:4}}>
                  <div style={{background:"#0061FF",height:6,borderRadius:4,width:"1%"}}/>
                </div>
                <div style={{fontSize:10,color:C.tL,marginTop:2}}>~1% terpakai · JSON + Excel harian</div>
              </div>
            </div>
            {supaStatus && (
              <div style={{background:supaStatus.ok?"#ECFDF5":"#FEF2F2",borderRadius:8,padding:"8px 12px",marginBottom:10,fontSize:12,color:supaStatus.ok?"#065F46":C.d,border:`1px solid ${supaStatus.ok?"#A7F3D0":C.dL}`}}>
                {supaStatus.msg}
              </div>
            )}
            <Btn full color="#3ECF8E" onClick={onSupaBackup} disabled={supaBackingUp} style={{marginTop:4}}>
              {supaBackingUp?"⟳ Memproses...":"☁ Backup Semua Data Sekarang ke Supabase"}
            </Btn>
            <div style={{marginTop:12,background:"#F0FDF4",borderRadius:10,padding:"12px 14px",border:"1px solid #BBF7D0"}}>
              <div style={{fontSize:12,fontWeight:700,color:"#065F46",marginBottom:6}}>⬆ Pulihkan dari Supabase (backup terakhir)</div>
              <div style={{fontSize:11,color:"#047857",marginBottom:10,lineHeight:1.6}}>
                Ambil data dari backup Supabase paling baru. Pilih kategori data yang ingin dipulihkan, atau pulihkan semua sekaligus.
              </div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:8}}>
                <Btn full outline color="#065F46" onClick={onSupaRestoreOps} disabled={supaBackingUp} style={{flex:1,minWidth:130}}>
                  🗓 Pulihkan Operasi
                </Btn>
                <Btn full outline color="#065F46" onClick={onSupaRestoreLembur} disabled={supaBackingUp} style={{flex:1,minWidth:130}}>
                  ⏰ Pulihkan Lembur
                </Btn>
                <Btn full outline color="#0369A1" onClick={onSupaRestoreMonitoring} disabled={supaBackingUp} style={{flex:1,minWidth:130}}>
                  🌡 Pulihkan Monitoring
                </Btn>
              </div>
              <Btn full color="#065F46" onClick={onSupaRestoreAll} disabled={supaBackingUp} style={{background:"#065F46",color:"#fff",border:"none",width:"100%"}}>
                ♻ Pulihkan SEMUA Data dari Supabase
              </Btn>
            </div>
          </Card>

          <Card>
            <SH label="📦 Backup ke Dropbox" color="#0061FF"/>
            <div style={{background:"#EFF6FF",borderRadius:8,padding:"10px 12px",marginBottom:14,fontSize:12,color:"#1D4ED8",lineHeight:1.8,border:"1px solid #BFDBFE"}}>
              <b>✅ Koneksi Dropbox aktif otomatis.</b><br/>
              Token terkonfigurasi secara terpusat — tidak perlu input manual.<br/>
              Path backup: <code style={{background:"#DBEAFE",padding:"1px 4px",borderRadius:3}}>/KamarBedahPantiRini/backup.json</code>
            </div>
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
            <Btn full color="#0061FF" onClick={onDbxBackup} disabled={dbxBacking} style={{background:"#0061FF",color:"#fff",border:"none",marginTop:4}}>
              {dbxBacking?"⟳ Menyimpan...":"📦 Backup Semua Data (JSON) ke Dropbox"}
            </Btn>
            <div style={{marginTop:10,background:"#EEF2FF",borderRadius:10,padding:"12px 14px"}}>
              <div style={{fontSize:12,fontWeight:700,color:"#3730A3",marginBottom:8}}>📊 Backup per Kategori (Excel) ke Dropbox</div>
              <div style={{fontSize:11,color:"#6366F1",marginBottom:10,lineHeight:1.6}}>
                Backup data Operasi atau Lembur secara terpisah sebagai file Excel (.xlsx) ke Dropbox.
                File tersimpan di folder yang sama dengan path backup utama.
              </div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                <Btn full outline color="#3730A3" onClick={onDbxBackupOpsXls} disabled={dbxBacking} style={{flex:1,minWidth:140}}>
                  🗓 Backup Operasi (Excel)
                </Btn>
                <Btn full outline color="#3730A3" onClick={onDbxBackupLemburXls} disabled={dbxBacking} style={{flex:1,minWidth:140}}>
                  ⏰ Backup Lembur (Excel)
                </Btn>
                <Btn full outline color="#0369A1" onClick={onDbxBackupMonitoringXls} disabled={dbxBacking} style={{flex:1,minWidth:140}}>
                  🌡 Backup Monitoring (Excel)
                </Btn>
              </div>
            </div>
            <div style={{marginTop:10}}>
              <div style={{fontSize:11,fontWeight:700,color:"#1D4ED8",marginBottom:6}}>⬇ Pulihkan dari Dropbox (pilih data yang ingin dipulihkan):</div>
              <div style={{display:"flex",gap:8}}>
                <Btn full outline color="#0061FF" onClick={onDbxRestoreOps} disabled={dbxBacking} style={{flex:1}}>
                  🗓 Jadwal Operasi
                </Btn>
                <Btn full outline color="#0061FF" onClick={onDbxRestoreLembur} disabled={dbxBacking} style={{flex:1}}>
                  ⏰ Data Lembur
                </Btn>
              </div>
            </div>
          </Card>

          <Card>
            <SH label="💾 Backup & Recovery Lokal" color={C.w}/>
            <div style={{fontSize:12,color:C.tL,marginBottom:12,lineHeight:1.6}}>Ekspor semua data lokal sebagai file JSON untuk backup darurat.</div>
            <Btn full outline color={C.w} onClick={()=>{
              const data={exportedAt:fNow(),ops,archive,notifs};
              downloadBlob(JSON.stringify(data,null,2),"BACKUP_OR_PantiRini_"+todayDate()+".json","application/json");
              showToast("✓ File JSON berhasil diunduh",C.s);
            }}>⬇ Ekspor Semua Data (JSON)</Btn>
          </Card>

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
