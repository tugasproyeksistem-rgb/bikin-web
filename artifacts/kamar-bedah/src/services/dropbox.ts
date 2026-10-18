/* src/services/dropbox.ts
   Service backup/restore Dropbox — semua panggilan lewat Supabase Edge
   Function "dropbox-proxy" (token Dropbox asli hanya hidup di server,
   TIDAK PERNAH menyentuh browser). Dipindah dari App.tsx (langkah 4
   refactor — pisah services), TIDAK ADA perubahan logika. */

import * as XLSX from "xlsx";
import type { DropboxConfig } from "../types";
import { SUPA_CLIENT } from "./supabase";

/* Semua backup/restore Dropbox TIDAK LAGI bicara langsung ke Dropbox API
   dari client — melainkan lewat Supabase Edge Function "dropbox-proxy"
   (lihat supabase/functions/dropbox-proxy/index.ts). Token asli hanya
   hidup sebagai secret di server (Deno.env), tidak pernah dikirim ke /
   disimpan di browser. Client hanya memegang DBX_PROXY_SECRET — sebuah
   shared secret terbatas yang HANYA bisa memicu 2 aksi spesifik
   (backup/restore 1 folder) lewat proxy ini, BUKAN kunci akses Dropbox
   penuh. Lihat catatan di vite-env.d.ts / .env.local untuk cara isi nilainya. */
export const DBX_PROXY_SECRET = (import.meta as any).env?.VITE_PROXY_SHARED_SECRET || "";
export const DBX_PROXY_FN     = "dropbox-proxy";

export const defaultDbxCfg: DropboxConfig = { path:"/KamarBedahPantiRini/backup.json", autoBackup:true, backupInterval:30, realtimeBackup:true, lastBackup:null, autoExcelBackup:true, lastExcelBackup:null, lastExcelBackupTs:null };

/* Helper tunggal: semua panggilan ke proxy lewat sini agar konsisten
   (header, error handling, parsing respons) */
export async function callDbxProxy(body: Record<string, any>): Promise<{ok:boolean; data?:any; msg:string}> {
  try {
    const {data, error} = await SUPA_CLIENT.functions.invoke(DBX_PROXY_FN, {
      body,
      headers: {"x-proxy-secret": DBX_PROXY_SECRET},
    });
    if(error) return {ok:false, msg:"Proxy error: "+(error.message||"tidak diketahui")};
    if(!data?.ok) return {ok:false, msg:data?.msg||"Gagal — proxy mengembalikan respons tidak sukses"};
    return {ok:true, data:data.data, msg:data.msg||"✓ Berhasil"};
  } catch(e:any){ return {ok:false, msg:"Gagal menghubungi proxy: "+(e?.message||"network error")}; }
}

export async function dropboxUpload(cfg: DropboxConfig, data: any): Promise<{ok:boolean;msg:string}> {
  const d: any = data || {};
  if(!(d.ops?.length) && !(d.staff?.length) && !(d.archive?.length))
    return {ok:false, msg:"Backup Dropbox dibatalkan: data lokal masih kosong (belum termuat)"};
  return callDbxProxy({action:"backup", path:cfg.path, payload:data});
}

export async function dropboxDownload(cfg: DropboxConfig): Promise<{ok:boolean;data?:any;msg:string}> {
  return callDbxProxy({action:"restore", path:cfg.path});
}

/* Upload Excel per-kategori ke Dropbox (lewat proxy — payload dikirim base64) */
/* SATU-SATUNYA jalur upload Excel ke Dropbox: semua tab memakai fungsi ini.
   Menerima ArrayBuffer atau Uint8Array (SheetJS maupun ExcelJS), menolak file kosong. */
export async function dropboxUploadBytes(path: string, data: ArrayBuffer | Uint8Array): Promise<{ok:boolean;msg:string}> {
  try {
    const u8 = data instanceof Uint8Array ? data : new Uint8Array(data);
    if (!u8.length) return {ok:false, msg:"File Excel kosong, tidak diunggah ke Dropbox"};
    const CHUNK = 8192;
    let bin = "";
    for (let i = 0; i < u8.length; i += CHUNK) bin += String.fromCharCode(...u8.subarray(i, i + CHUNK));
    return callDbxProxy({action:"upload_excel", path, base64: btoa(bin)});
  } catch(e:any){ return {ok:false, msg:"Gagal menyiapkan file Excel: "+(e?.message||"unknown error")}; }
}

/* Workbook SheetJS (Laporan Harian, Monitoring, dll.) */
export async function dropboxUploadExcel(path: string, wb: any): Promise<{ok:boolean;msg:string}> {
  try {
    const buf = XLSX.write(wb, {type:"array", bookType:"xlsx"}) as ArrayBuffer;
    return dropboxUploadBytes(path, buf);
  } catch(e:any){ return {ok:false, msg:"Gagal menyiapkan file Excel: "+(e?.message||"unknown error")}; }
}
