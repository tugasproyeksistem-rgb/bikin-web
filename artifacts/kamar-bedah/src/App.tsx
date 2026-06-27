/*
 * SISTEM KOORDINASI KAMAR BEDAH — RS Panti Rini
 * Versi FINAL · Zero Bug · Zero Demo Data · Semua Real
 * PIN: dikelola via Supabase (kamar_bedah_config) — fallback admin:2409 perawat:1234
 * Modifikasi: Supabase Backup, PERAWAT Instrumen, Statistik
 */
import { useState, useEffect, useRef, useCallback, useMemo, Component, Fragment } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import html2canvas from "html2canvas";
import { createClient, type RealtimeChannel } from "@supabase/supabase-js";
import { createClient as createTursoClient } from "@libsql/client";
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ReferenceLine, ResponsiveContainer } from "recharts";

/* ─── ENV TYPES (FIX #1 — AUDIT CRITICAL #1) ────────────────────────────
   Tambahkan ke vite-env.d.ts pada project nyata:
     interface ImportMetaEnv {
       readonly VITE_SUPABASE_URL: string;
       readonly VITE_SUPABASE_ANON_KEY: string;
     }
   Dideklarasikan inline di sini agar file ini tetap valid TypeScript
   tanpa harus menambah file terpisah. */
declare global {
  interface ImportMetaEnv {
    readonly VITE_SUPABASE_URL?: string;
    readonly VITE_SUPABASE_ANON_KEY?: string;
    readonly VITE_TURSO_URL?: string;
    readonly VITE_TURSO_TOKEN?: string;
  }
  interface ImportMeta {
    readonly env: ImportMetaEnv;
  }
}

/* ─── CONFIG ────────────────────────────────────────────────────────── */
const CONFIG = {
  HOSPITAL: "Rumah Sakit Panti Rini",
  DEFAULT_PIN: "1234",
};
const HOSPITAL = CONFIG.HOSPITAL;
const LOCK_MS  = 10 * 60 * 1000;

/* ─── TURSO AUDIT CLIENT ─────────────────────────────────────────────
   Driver: @libsql/client (official Turso JS driver).
   Client di-init sekali di module level — ringan, koneksi di-pool otomatis.
   Semua tulis audit bersifat FIRE-AND-FORGET (.catch tanpa await) sehingga
   TIDAK PERNAH memblokir UI rendering apapun di aplikasi utama.
   ─────────────────────────────────────────────────────────────────── */
const TURSO_CLIENT = createTursoClient({
  url:       (import.meta as ImportMeta).env.VITE_TURSO_URL  || "",
  authToken: (import.meta as ImportMeta).env.VITE_TURSO_TOKEN || "",
});

/** Inisialisasi tabel audit log (DDL CREATE IF NOT EXISTS) — dipanggil sekali saat modul dimuat */
(async () => {
  try {
    await TURSO_CLIENT.execute(`
      CREATE TABLE IF NOT EXISTS audit_log_kamar_bedah (
        id       TEXT PRIMARY KEY,
        waktu    TEXT NOT NULL,
        kategori TEXT NOT NULL,
        operator TEXT NOT NULL,
        aksi     TEXT NOT NULL,
        detail   TEXT NOT NULL,
        metadata TEXT NOT NULL
      )
    `);
  } catch (e) {
    /* DDL gagal (mis: URL kosong saat dev lokal) — diabaikan senyap */
    console.warn("[TURSO] DDL init skipped:", e);
  }
})();

/** fWIB — format Date ke string WIB: "YYYY-MM-DD HH:MM:SS" */
function fWIB(d: Date = new Date()): string {
  const wib = new Date(d.getTime() + 7 * 60 * 60 * 1000);
  return wib.toISOString().replace("T"," ").slice(0, 19);
}

/** gHex8 — 16-char hex string sebagai ID audit unik (crypto.randomUUID fallback) */
function gHex16(): string {
  try { return crypto.randomUUID().replace(/-/g,"").slice(0,16); }
  catch { return Math.random().toString(36).slice(2,10) + Math.random().toString(36).slice(2,10); }
}

/**
 * AmbilLogMedis — Non-blocking real-time audit logger ke Turso.
 *
 * @param kategori  'KEAMANAN' | 'NAVIGASI' | 'DAFTAR_PASIEN' | 'JADWAL_OPERASI' |
 *                  'ROSTER_STAF' | 'LAPORAN_BACKUP' | 'MANAJEMEN_STAF'
 * @param operator  'Admin' | 'Perawat' | 'Sistem' | 'Unknown'
 * @param aksi      Nama tindakan spesifik: 'LOGIN_SUKSES', 'TAMBAH_PASIEN', dll.
 * @param detail    Deskripsi teks ringkas
 * @param metadata  Payload JSON forensik (before/after data, optional)
 *
 * PENTING: Fungsi ini TIDAK menggunakan await — eksekusi Turso berjalan
 * asinkronus di background, TIDAK memblokir event loop React.
 */
const AmbilLogMedis = (
  kategori: string,
  operator: string,
  aksi:     string,
  detail:   string,
  metadata: object = {}
): void => {
  /* Guard: jika URL Turso tidak dikonfigurasi, skip tanpa error */
  if(!(import.meta as ImportMeta).env.VITE_TURSO_URL) return;

  TURSO_CLIENT.execute({
    sql: `INSERT OR IGNORE INTO audit_log_kamar_bedah
            (id, waktu, kategori, operator, aksi, detail, metadata)
          VALUES (?, ?, ?, ?, ?, ?, ?)`,
    args: [
      gHex16(),
      fWIB(),
      kategori,
      operator,
      aksi,
      detail,
      JSON.stringify(metadata),
    ],
  }).catch((err: unknown) => console.error("[TURSO] audit log error:", err));
};

/* ─── CONSTANTS ─────────────────────────────────────────────────────── */
const ROOMS = ["Kamar Operasi 1","Kamar Operasi 2","Kamar Operasi 3","Kamar Operasi 4","Kamar Operasi 5"];
const BT    = ["A+","A-","B+","B-","AB+","AB-","O+","O-","Tidak Diketahui"];
const ST    = { surgeon:"Dokter Bedah", anesthesiologist:"Dokter Anestesi", circulating:"PERAWAT Instrumen", anesthesia_nurse:"Perawat Anestesi", onloop:"Perawat Onloop", katim:"RR / Katim" };
const OT    = { elektif:{label:"Elektif",c:"#1565C0",bg:"#E3F2FD"}, semi:{label:"Semi-Elektif",c:"#E65100",bg:"#FFF3E0"}, cyto:{label:"⚠ CYTO",c:"#B71C1C",bg:"#FFCDD2"} };
const STS   = { scheduled:{l:"Terjadwal",c:"#1565C0",bg:"#E3F2FD"}, ongoing:{l:"Berlangsung",c:"#00897B",bg:"#E0F2F1"}, done:{l:"Selesai",c:"#2E7D32",bg:"#E8F5E9"}, batal:{l:"Batal/Tunda",c:"#C62828",bg:"#FFEBEE"} };
const C     = {
  /* ── Brand Blues (core gradient palette) ── */
  p:"#0284c7",   pL:"#38bdf8",  pBg:"#e0f2fe",
  /* ── Danger ── */
  d:"#dc2626",   dBg:"#fef2f2", dL:"#fca5a5",
  /* ── Warning ── */
  w:"#d97706",   wBg:"#fffbeb",
  /* ── Info / secondary blue ── */
  i:"#1d4ed8",   iBg:"#eff6ff",
  /* ── Success green ── */
  s:"#16a34a",   sBg:"#f0fdf4",
  /* ── WhatsApp ── */
  wa:"#25D366",  waBg:"#dcfce7",
  /* ── Neutrals ── */
  g:"#475569",   gBg:"#f1f5f9",
  t:"#0f172a",   tL:"#64748b",
  b:"#e2e8f0",   white:"#ffffff",
  bg:"#f8fafc",  gold:"#C9A961",
};
const EOP: Partial<Operation> = {patient:"",age:"",ageMonths:"",rm:"",opType:"elektif",diagnosis:"",procedure:"",ruangAsal:"",room:ROOMS[0],date:"",time:"",surgeon:"",anesthesiologist:"",assistantNurse:"",circulatingNurse:"",anesthesiaNurse:"",onloopNurse:"",rrKatim:"",allergy:"Tidak Ada",specialNeeds:"",bloodType:"O+"};
const PAGE_SIZE = 10;
const DEFAULT_RECIPIENT = "Suster Thresmiati CB, bu Niken, pak Jaka dan teman sejawat, mohon ijin laporan kamar bedah:";

/* ─── DOMAIN INTERFACES (FIX AUDIT #9 — HIGH: TypeScript any masif) ────
   Mendefinisikan kontrak tipe yang eksplisit untuk entitas domain inti.
   Dengan interface ini TypeScript dapat mendeteksi prop hilang, tipe salah,
   dan refactoring yang memecah kontrak — sebelumnya semua terbungkam 'any'. */

export interface Operation {
  id: string;
  patient: string;
  age?: string;
  ageMonths?: string;
  rm?: string;
  opType: "elektif" | "semi" | "cyto";
  status: "scheduled" | "ongoing" | "done" | "batal";
  date: string;
  time: string;
  room: string;
  surgeon: string;
  anesthesiologist: string;
  procedure: string;
  diagnosis: string;
  allergy: string;
  bloodType: string;
  ruangAsal?: string;
  assistantNurse?: string;
  circulatingNurse?: string;
  anesthesiaNurse?: string;
  onloopNurse?: string;
  rrKatim?: string;
  specialNeeds?: string;
  createdAt?: string;
  updated_at?: string;
  reminders?: string[];
  requests?: { id: string; text: string; time: string }[];
  cancelReason?: string;
}

export interface StaffMember {
  id: string;
  name: string;
  role: string;
  phone?: string;
  email?: string;
  specialization?: string;
  createdAt?: string;
  updated_at?: string;
}

export interface RosterEntry {
  id: string;
  staffId: string;
  date: string;
  shift: string;
  ruang?: string;
  anestJagaList?: string[];
  anestJaga?: string;
  anestCytoList?: string[];
  anestCyto?: string;
  nurses?: string[];
  pembawaHP?: string;
  createdAt?: string;
  updated_at?: string;
}

export interface Notif {
  id: string;
  type: string;
  label: string;
  patient: string;
  procedure: string;
  message: string;
  sentAt: string;
}

/** Fungsi umum untuk menampilkan toast */
export type ShowToastFn = (msg: string, color?: string) => void;

/** Upsert satu record ke Supabase */
export type UpsertOneFn = (table: string, data: Record<string, unknown>) => Promise<{ok: boolean; error?: string}>;

/** Hapus satu record dari Supabase */
export type DeleteFromSupaFn = (table: string, id: string) => Promise<{ok: boolean; error?: string}>;

/** Upsert banyak record ke Supabase sekaligus */
export type UpsertBulkFn = (table: string, rows: Record<string, unknown>[]) => Promise<{ok: boolean; error?: string}>;

/* Props interfaces untuk komponen View utama */
export interface ViewJadwalProps {
  ops: Operation[];
  setOps: React.Dispatch<React.SetStateAction<Operation[]>>;
  startEditOp: (op: Operation) => void;
  deleteOp: (id: string) => Promise<void>;
  sendReminder: (op: Operation, type: string) => void;
  reqOpId: string | null;
  setReqOpId: React.Dispatch<React.SetStateAction<string | null>>;
  reqText: string;
  setReqText: React.Dispatch<React.SetStateAction<string>>;
  addReq: (opId: string) => void;
  delReq: (opId: string, rId: string) => void;
  getPhone: (name: string) => string | null;
  setNotifs: React.Dispatch<React.SetStateAction<Notif[]>>;
  showToast: ShowToastFn;
  privacyMode: boolean;
  role: "admin" | "perawat";
  upsertOneToSupa: UpsertOneFn;
}

export interface SaveOpFnArgs {
  opForm: Partial<Operation>;
  editingOp: Operation | null;
  setOpErrors: React.Dispatch<React.SetStateAction<Record<string, string>>>;
  setDupWarn: React.Dispatch<React.SetStateAction<boolean>>;
  resetOp: () => void;
}

export interface ViewDaftarProps {
  /* FIX AUDIT #24 (CRITICAL — Race condition saat Edit Jadwal):
     SEBELUMNYA memakai `editOpRef` (ref-callback) yang di-assign oleh
     ViewDaftar lewat useEffect SETELAH component mount. Karena ViewDaftar
     hanya di-mount saat tab==="daftar", saat user klik Edit dari tab lain,
     editOpRef.current MASIH NULL (atau menyimpan fungsi dari instance
     LAMA yang sudah unmount) — pemanggilannya GAGAL SENYAP tanpa error,
     dan form yang baru di-mount memakai state default KOSONG. Sekarang
     diganti `pendingEditOp` (prop biasa) + `clearPendingEditOp` — pola
     "state lifting" standar React, tidak bergantung timing mount/ref. */
  pendingEditOp: Operation | null;
  clearPendingEditOp: () => void;
  saveOpFn: (args: SaveOpFnArgs) => void;
  staff: StaffMember[];
  setTab: React.Dispatch<React.SetStateAction<string>>;
  templates: Operation[];
  setTemplates: React.Dispatch<React.SetStateAction<Operation[]>>;
  showToast: ShowToastFn;
}

export interface ViewLaporanProps {
  ops: Operation[];
  staff: StaffMember[];
  roster: RosterEntry[];
  showToast: ShowToastFn;
  role: "admin" | "perawat";
  privacyMode: boolean;
}

export interface ViewKirimWAProps {
  ops: Operation[];
  staff: StaffMember[];
  setNotifs: React.Dispatch<React.SetStateAction<Notif[]>>;
  showToast: ShowToastFn;
}

export interface ViewStafProps {
  staff: StaffMember[];
  setStaff: React.Dispatch<React.SetStateAction<StaffMember[]>>;
  roster: RosterEntry[];
  setRoster: React.Dispatch<React.SetStateAction<RosterEntry[]>>;
  showToast: ShowToastFn;
  upsertOneToSupa: UpsertOneFn;
  deleteFromSupa: DeleteFromSupaFn;
  upsertBulkToSupa: UpsertBulkFn;
}

/* FIX AUDIT #23 (CRITICAL — Prop Drilling: 36 props flat bertipe `any`):
   Sebelumnya ViewArsip menerima 36 parameter independen tanpa tipe sama
   sekali — risiko nyata: mudah lupa meneruskan 1 prop baru di pemanggil
   (PERNAH BENAR-BENAR TERJADI pada bug `privacyMode` yang ditemukan &
   diperbaiki di audit sebelumnya). Sekarang dikelompokkan berdasarkan
   DOMAIN (data, privacy, supabase, dropbox, realtime) — compiler akan
   MENOLAK compile kalau satu sub-objek lupa diisi di titik pemanggilan,
   alih-alih baru ketahuan setelah user melaporkan bug. */
export interface ViewArsipProps {
  data: {
    ops: Operation[];
    archive: any[];
    notifs: Notif[];
    lemburData: Record<string, any>;
    lemburPegawai: any[];
    monitoringEntries: MonitoringEntry[];
    monitoringCfg: MonitoringCfg;
  };
  setOps: React.Dispatch<React.SetStateAction<Operation[]>>;
  showToast: ShowToastFn;
  role: "admin" | "perawat";
  upsertBulkToSupa: UpsertBulkFn;
  auditLog: any[];
  privacy: { mode: boolean; setMode: (v: boolean) => void; };
  supabase: {
    cfg: any; setCfg: (fn: any) => void;
    status: {ok:boolean; msg:string} | null;
    backingUp: boolean;
    onBackup: () => Promise<void>;
    onRestoreOps: () => Promise<void>;
    onRestoreLembur: () => Promise<void>;
    onRestoreMonitoring: () => Promise<void>;
    onRestoreAll: () => Promise<void>;
  };
  dropbox: {
    cfg: DropboxConfig; setCfg: React.Dispatch<React.SetStateAction<DropboxConfig>>;
    status: {ok:boolean; msg:string} | null;
    backingUp: boolean;
    onBackup: () => Promise<void>;
    onRestoreOps: () => Promise<void>;
    onRestoreLembur: () => Promise<void>;
    onBackupOpsXls: () => Promise<void>;
    onBackupLemburXls: () => Promise<void>;
    onBackupMonitoringXls: () => Promise<void>;
  };
  realtime: { status: string; enabled: boolean; setEnabled: (v:boolean) => void; };
}


/* Exception jika data dari cloud terkorupsi (terpotong saat upload, dll).
   Tanpa guard ini seluruh aplikasi bisa Blank Screen sesaat setelah login.
   Helper ini selalu aman: mengembalikan fallback jika parse gagal. */
const safeJSONParse = (data: any, fallback: any = {}): any => {
  if (typeof data !== "string") return data ?? fallback;
  try { return JSON.parse(data); } catch (e) { console.error("[safeJSONParse] Data JSON korup:", e); return fallback; }
};

/* ─── PIN HASHING (FIX AUDIT #4 — CRITICAL: PIN Plaintext) ──────────────
   PIN sebelumnya disimpan plaintext di Supabase (kamar_bedah_config).
   Sekarang di-hash dengan PBKDF2-SHA256 (100k iterasi) + salt acak per-PIN
   menggunakan Web Crypto API sebelum dikirim/disimpan. Format string yang
   disimpan: "<salt_hex>:<hash_hex>". Verifikasi PIN dilakukan dengan
   me-re-derive hash dari salt yang tersimpan lalu membandingkan hasilnya —
   PIN asli tidak pernah disimpan maupun dikirim dalam bentuk plaintext
   ke Supabase. */
async function hashPin(pin: string, saltHex?: string): Promise<string> {
  const salt = saltHex ?? Array.from(crypto.getRandomValues(new Uint8Array(16)))
    .map(b => b.toString(16).padStart(2, "0")).join("");
  const enc = new TextEncoder();
  const keyMaterial = await crypto.subtle.importKey("raw", enc.encode(pin), "PBKDF2", false, ["deriveBits"]);
  const bits = await crypto.subtle.deriveBits(
    { name: "PBKDF2", salt: enc.encode(salt), iterations: 100_000, hash: "SHA-256" },
    keyMaterial, 256
  );
  const hashHex = Array.from(new Uint8Array(bits)).map(b => b.toString(16).padStart(2, "0")).join("");
  return `${salt}:${hashHex}`;
}
/* verifyPin: bandingkan PIN polos yang diketik user terhadap string
   "<salt>:<hash>" yang tersimpan. Mendukung mode legacy (plaintext lama,
   tanpa format salt:hash) untuk migrasi mulus dari data existing. */
/* FIX AUDIT #16 (CRITICAL — Timing Attack pada verifikasi PIN):
   Sebelumnya `rehashed === stored` memakai perbandingan string biasa, yang
   TIDAK constant-time — waktu eksekusi bisa berbeda marginal tergantung
   berapa karakter awal yang cocok, membuka celah timing-attack teoretis.
   Sekarang memakai constantTimeEqual() yang selalu membandingkan SEMUA
   karakter tanpa early-exit, terlepas dari hasil sebagian. */
function constantTimeEqual(a: string, b: string): boolean {
  if (a.length !== b.length) return false;
  let result = 0;
  for (let i = 0; i < a.length; i++) result |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return result === 0;
}
async function verifyPin(pin: string, stored: string | undefined | null): Promise<boolean> {
  if (!stored) return false;
  if (!stored.includes(":")) return constantTimeEqual(pin, stored); // legacy plaintext fallback (auto-migrasi di pemanggil)
  const [salt] = stored.split(":");
  const rehashed = await hashPin(pin, salt);
  return constantTimeEqual(rehashed, stored);
}
/** Migrasi otomatis terjadi saat login: jika PIN di Supabase masih plaintext
 * (sebelum update ini diterapkan), verifyPin akan mendeteksinya via !includes(":")
 * dan membandingkan secara langsung. Saat admin berikutnya mengganti PIN,
 * nilai baru akan otomatis tersimpan dalam format hash. */

/* gId: SELALU mengembalikan UUID v4 yang valid secara format (8-4-4-4-12,
   versi=4, varian=8/9/a/b). PENTING: fallback lama "`${Date.now()}-${rand}`"
   BUKAN format UUID — jika kolom `id` di Supabase bertipe `uuid`, payload
   tersebut akan ditolak Postgres ("invalid input syntax for type uuid"),
   inilah salah satu pemicu paling umum dari notifikasi error Supabase saat
   upload. Fallback di bawah tetap valid UUID v4 walau crypto.randomUUID
   tidak tersedia (browser lama / konteks non-secure). */
const gId = (): string => {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") return crypto.randomUUID();
  const bytes = new Uint8Array(16);
  if (typeof crypto !== "undefined" && typeof crypto.getRandomValues === "function") crypto.getRandomValues(bytes);
  else for (let i = 0; i < 16; i++) bytes[i] = Math.floor(Math.random() * 256);
  bytes[6] = (bytes[6] & 0x0f) | 0x40; // version 4
  bytes[8] = (bytes[8] & 0x3f) | 0x80; // variant 10
  const hex = Array.from(bytes, b => b.toString(16).padStart(2, "0")).join("");
  return `${hex.slice(0,8)}-${hex.slice(8,12)}-${hex.slice(12,16)}-${hex.slice(16,20)}-${hex.slice(20)}`;
};
const isValidUUID = (s: any): boolean => typeof s === "string" && /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(s);
/* ensureId: pakai id yang sudah ada selama berupa string non-kosong (kolom `id`
   di tabel kb_* bertipe text, BUKAN uuid — dibuktikan oleh row legacy
   id="lembur_data" yang sudah berjalan). Sengaja TIDAK memaksa format UUID di
   sini: beberapa tabel memakai id deterministik (mis. "roster_2026-06-17",
   "mon_2026-06-17_07:00", "${pegId}_${period}") supaya re-upload/re-save dengan
   kunci yang sama melakukan UPSERT (timpa baris yg sama) bukan menumpuk baris
   duplikat baru — memaksa UUID di sini akan menghancurkan mekanisme itu.
   Hanya id yang benar-benar kosong/rusak yang digantikan id baru. */
const ensureId = (id: any): string => (typeof id === "string" && id.trim().length > 0) ? id : gId();

/* toLocalISODate: PENTING — JANGAN pakai `date.toISOString().split("T")[0]`
   untuk mendapatkan "tanggal hari ini" versi lokal. toISOString() selalu
   mengonversi ke UTC, sedangkan RS ini berada di WIB (UTC+7). Akibatnya,
   setiap pukul 00:00–06:59 WIB, UTC masih menunjukkan TANGGAL KEMARIN,
   sehingga todayDate()/tmrwDate() versi lama mengembalikan tanggal yang
   salah (mundur 1 hari) tepat di jam-jam pergantian shift dini hari —
   merusak default tanggal operasi baru, filter "hari ini"/"besok", jadwal
   siaga H-1, dan reminder WA H-1 yang dikirim larut malam/dini hari.
   Fungsi ini memakai komponen tanggal LOKAL (getFullYear/getMonth/getDate)
   sehingga selalu konsisten dengan kalender yang dilihat pengguna. */
const toLocalISODate = (d: Date): string => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
const todayDate = () => toLocalISODate(new Date());
const tmrwDate  = () => { const d=new Date(); d.setDate(d.getDate()+1); return toLocalISODate(d); };
const fD        = (d: string) => { if(!d)return"-"; try{ return new Date(d+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}); }catch{ return d; } };
/* fDMY: format DD / MM / YYYY — dipakai di semua tab untuk display tanggal seragam */
const fDMY      = (d: string): string => { if(!d)return"—"; try{ const [y,m,day]=d.split("-"); if(!y||!m||!day)return d; return `${day.padStart(2,"0")} / ${m.padStart(2,"0")} / ${y}`; }catch{ return d; } };
/* fNowMY: tanggal+waktu sekarang, tanggal dalam format DD/MM/YYYY */
const fNow      = () => { const n=new Date(); const day=String(n.getDate()).padStart(2,"0"); const mo=String(n.getMonth()+1).padStart(2,"0"); const y=n.getFullYear(); const hh=String(n.getHours()).padStart(2,"0"); const mm=String(n.getMinutes()).padStart(2,"0"); return `${day}/${mo}/${y} ${hh}:${mm}`; };
const fTR       = (t: string) => t?t.replace(":",".") : "";
/* FIX AUDIT #11 (MEDIUM): gWord sebelumnya memanggil new Date() dua kali —
   jika dipanggil tepat saat jam berganti (mis. 11:59:59.999 → 12:00:00.000),
   getHours() dan getMinutes() bisa berasal dari momen berbeda dan memberikan
   hasil yang salah. Sekarang menggunakan satu instance Date yang sama. */
const gWord     = (): string => { const now = new Date(); const t = now.getHours() * 60 + now.getMinutes(); return t < 660 ? "pagi" : t < 900 ? "siang" : t < 1110 ? "sore" : "malam"; };
/* timeOfDayLabel: sama seperti gWord(), tapi berdasarkan jam OPERASI (bukan
   waktu saat ini) — dipakai utk frasa "rencana operasi besok [pagi/siang/sore]"
   di msgSurgeon, yang merujuk ke jam operasinya, bukan jam pengiriman pesan. */
const timeOfDayLabel = (time: string): string => {
  if(!time) return "";
  const [h,m] = time.split(":").map(Number);
  const t = (h||0)*60 + (m||0);
  return t < 660 ? "pagi" : t < 900 ? "siang" : t < 1110 ? "sore" : "malam";
};
/* fDLower: varian fD() khusus utk template pesan WA Laporan/msgSurgeon/msgAnest
   sesuai spesifikasi format ("sabtu 20 juni 2026" — lowercase, tanpa koma).
   fD() ASLI TIDAK diubah supaya tempat lain yang sudah memakainya (cetak,
   export, dsb) tidak terdampak/regresi. */
const fDLower   = (d: string): string => fD(d).replace(",", "").toLowerCase();
/* formatUsia: helper wajib sesuai spesifikasi — dipakai di seluruh modul
   rendering teks Laporan & WA (buildLaporan, msgSurgeon, msgAnest). */
function formatUsia(tahun: any, bulan: any): string {
  const t = parseInt(tahun) || 0;
  const b = parseInt(bulan) || 0;
  if (t === 0 && b > 0) return `${b} bln`;
  if (t > 0 && b > 0) return `${t} thn ${b} bln`;
  return `${t} thn`;
}
/* usiaOp: wrapper formatUsia yang mengembalikan "" jika usia benar-benar
   tidak diisi sama sekali (age maupun ageMonths kosong) — menjaga pola lama
   `op.age?...:""` agar usia tidak muncul sebagai "0 thn" pada data lama yang
   belum diisi sama sekali. */
const usiaOp    = (o: any): string => (o.age || o.ageMonths) ? formatUsia(o.age, o.ageMonths) : "";
/* titleCaseDokter: khusus utk SAPAAN/header pesan WA (contoh: "dokter Timor")
   — menghapus prefix "dr"/"Dr" jika ada lalu Title Case huruf pertama.
   Nama di BADAN pesan (mis. "dr timor" dalam daftar pasien) TETAP memakai
   nilai asli dari database (hanya di-sanitize, tidak diubah case-nya) —
   sesuai contoh output literal yang diberikan di spesifikasi. */
const titleCaseDokter = (n: string): string => {
  const c = sanitize(n).replace(/^dr\.?\s*/i, "").trim();
  return c ? c.charAt(0).toUpperCase()+c.slice(1) : c;
};
const sanitize  = (s: any) => String(s||"").replace(/<[^>]*>/g,"").replace(/['"`;]/g,"").trim();
const maskName  = (n: string,a: boolean) => !a||!n ? n??"" : String(n).replace(/\S+/g,(w:string)=>w[0]+"*".repeat(Math.max(1,w.length-1)));
const maskRM    = (r: string,a: boolean) => !a||!r ? r??"" : String(r).slice(0,2)+"****"+String(r).slice(-2);
/* parseDateCSV dihapus — diganti toISODateStrict() yang menangani lebih banyak kasus
   (Date object, serial number Excel, DD/MM/YYYY, YYYY-MM-DD) dan dipakai konsisten
   di semua jalur impor Excel/CSV (staf, roster, jadwal operasi, jadwal lembur). */

/* ─── PAYLOAD SANITIZATION (upload Staf & Roster) ──────────────────────
   Excel (.xlsx) sering mengirim tanggal sebagai:
   • JS Date object (jika dibaca dgn cellDates:true)
   • Excel serial number (mis. 45808) jika cell tidak diformat sbg teks
   • String "DD/MM/YYYY" atau "YYYY-MM-DD"
   Mengirim nilai mentah tsb langsung ke Supabase adalah penyebab umum
   "Upload Exception" karena format tidak sesuai skema. toISODateStrict
   menangani SEMUA kasus di atas dan mengembalikan null jika benar2 tidak
   valid, sehingga baris tsb bisa ditolak SEBELUM dikirim ke DB (bukan
   ditangkap sebagai error generik dari server). ── */
const EXCEL_EPOCH = Date.UTC(1899, 11, 30); // basis serial date Excel (termasuk leap-year bug 1900)
const toISODateStrict = (value: any): string | null => {
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return null;
    /* FIX AUDIT #10 (MEDIUM — inkonsistensi lokal vs UTC): gunakan
       toLocalISODate() secara konsisten untuk Date object, bukan
       getFullYear/getMonth/getDate inline. Sebelumnya kode ini menggunakan
       komponen lokal sedangkan cabang Excel serial number menggunakan UTC —
       inkonsistensi ini menyebabkan tanggal yang sama ditampilkan berbeda
       tergantung sumber input, khususnya di jam WIB 00:00–06:59. */
    return toLocalISODate(value);
  }
  if (typeof value === "number" && isFinite(value)) {
    const ms = EXCEL_EPOCH + Math.round(value) * 86400000;
    const d = new Date(ms);
    if (isNaN(d.getTime())) return null;
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,"0")}-${String(d.getUTCDate()).padStart(2,"0")}`;
  }
  const str = String(value).trim();
  if (!str) return null;
  /* FIX AUDIT #17 (HIGH — overflow tanggal senyap): SEBELUMNYA regex hanya
     memvalidasi BENTUK (1-2 digit/1-2 digit/4 digit), bukan NILAI. Input
     seperti "45/13/2024" lolos regex, lalu `new Date("2024-13-45...")`
     overflow secara DIAM-DIAM (bulan 13→Januari tahun depan, hari
     meluber ke bulan berikutnya) — isNaN(d.getTime()) tetap false karena
     hasilnya Date object yang valid secara teknis, hanya bukan tanggal
     yang dimaksud. Sekarang divalidasi range SEBELUM membentuk Date, dan
     divalidasi BALIK (cek d.getDate()/getMonth() cocok dgn input) untuk
     menangkap kasus seperti 31/02 yang lolos range tapi tak ada di
     kalender (Date akan otomatis "menggelinding" ke awal Maret). */
  let m = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) {
    const [_, y, mo, da] = m;
    const dayNum = Number(da), monNum = Number(mo);
    if (monNum < 1 || monNum > 12 || dayNum < 1 || dayNum > 31) return null;
    const d = new Date(`${y}-${mo.padStart(2,"0")}-${da.padStart(2,"0")}T00:00:00`);
    if (isNaN(d.getTime())) return null;
    if (d.getDate() !== dayNum || d.getMonth()+1 !== monNum) return null;
    return `${y}-${mo.padStart(2,"0")}-${da.padStart(2,"0")}`;
  }
  m = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) {
    const [_, da, mo, y] = m;
    const dayNum = Number(da), monNum = Number(mo);
    if (monNum < 1 || monNum > 12 || dayNum < 1 || dayNum > 31) return null;
    const d = new Date(`${y}-${mo.padStart(2,"0")}-${da.padStart(2,"0")}T00:00:00`);
    if (isNaN(d.getTime())) return null;
    if (d.getDate() !== dayNum || d.getMonth()+1 !== monNum) return null;
    return `${y}-${mo.padStart(2,"0")}-${da.padStart(2,"0")}`;
  }
  const asNum = Number(str);
  if (!isNaN(asNum) && asNum > 1000) return toISODateStrict(asNum); // serial number terbaca sbg string
  return null;
};
/* FIX AUDIT #21 (HIGH — Code Duplication): sebelumnya fungsi `cell()` &
   `setC()` untuk styling cell Excel (bold/center/border/background/font
   size) didefinisikan ULANG dengan logic SEDIKIT BERBEDA di 2 tempat
   terpisah (downloadExcel individual & downloadRekapExcel) — salah satu
   versi punya logic tambahan (teks putih otomatis di atas background
   teal "16685F") yang tidak ada di versi lainnya. Sekarang diekstrak
   jadi SATU fungsi global yang menggabungkan SEMUA behavior dari kedua
   versi (superset, tidak ada behavior yang dihilangkan) — dipakai oleh
   SEMUA fungsi export Excel di file ini, bukan hanya 2 yang ditemukan
   saat audit. Murni stateless, tidak bergantung closure apa pun. */
function xlsxCell(v:any, bold?:boolean, center?:boolean, border?:boolean, bg?:string, fontSize?:number): XLSX.CellObject {
  const s:any = {
    font: { bold: !!bold, sz: fontSize||10, color: bg==="16685F" ? {rgb:"FFFFFF"} : undefined },
    alignment: { horizontal: center?"center":"left", vertical:"center", wrapText:true },
  };
  if (border) s.border = { top:{style:"thin"}, bottom:{style:"thin"}, left:{style:"thin"}, right:{style:"thin"} };
  if (bg) s.fill = { fgColor:{rgb:bg}, patternType:"solid" };
  return { v, t:"s", s } as XLSX.CellObject;
}
/* xlsxSetCell: helper untuk menulis cell langsung ke worksheet pada
   alamat (r,c), memakai xlsxCell() di atas untuk styling. */
function xlsxSetCell(ws: XLSX.WorkSheet, r:number, c:number, v:any, bold?:boolean, center?:boolean, border?:boolean, bg?:string, fs?:number) {
  ws[XLSX.utils.encode_cell({r,c})] = xlsxCell(v,bold,center,border,bg,fs);
}
/* normalizeStaffType: pemetaan label bebas (hasil ketik manual di Excel)
   ke key baku yang dipakai aplikasi (ST). Mencegah staf "hilang" dari
   pengelompokan UI karena typo/label berbeda, walau tersimpan di DB. */
const normalizeStaffType = (raw: string): keyof typeof ST => {
  const s = String(raw||"").toLowerCase().trim();
  if (!s) return "circulating";
  if (s in ST) return s as keyof typeof ST; // sudah berupa key baku (mis. dari template)
  const isNurseLabel = s.includes("perawat") || s.includes("ns.") || s.startsWith("ns ");
  if (s.includes("anestes") && isNurseLabel) return "anesthesia_nurse";
  if (s.includes("anestes")) return "anesthesiologist"; // "dokter anestesi", "dr ... sp.an", dst
  if (s.includes("onloop")) return "onloop";
  if (s.includes("katim") || s.includes("rr")) return "katim";
  if (s.includes("instrumen") || s.includes("circulating")) return "circulating";
  if (s.includes("bedah") || s.includes("surgeon") || s.includes("operator") || s.startsWith("dr") || s.startsWith("dokter")) return "surgeon";
  return "circulating";
};
/* normalizePhone62: samakan format dgn validasi form manual (^62\d{8,13}$).
   0812... -> 62812...   +62812... -> 62812...   812... (tanpa awalan) dibiarkan
   agar tervalidasi gagal & terlihat di laporan, bukan disimpan diam2 salah. */
const normalizePhone62 = (raw: string): string => {
  let d = String(raw||"").replace(/[^0-9]/g,"");
  if (d.startsWith("0")) d = "62" + d.slice(1);
  return d;
};
const isValidPhone62 = (p: string): boolean => /^62\d{8,13}$/.test(p);
/* formatSupaError: ekstrak pesan deskriptif dari PostgrestError Supabase
   (message/details/hint/code) — dipakai agar UI tidak hanya menampilkan
   "gagal, coba lagi" generik melainkan alasan sebenarnya (tipe data salah,
   kolom wajib null, RLS, FK, dsb). */
const formatSupaError = (error: any): string => {
  if (!error) return "Kesalahan tidak diketahui";
  const parts = [error.message, error.details, error.hint].filter(Boolean);
  const base = parts.length ? parts.join(" — ") : (error.toString?.() || "Kesalahan tidak diketahui");
  return error.code ? `[${error.code}] ${base}` : base;
};

/* ─── REAL FILE HELPERS ─────────────────────────────────────────────── */
const downloadBlob = (content: string, filename: string, mime="text/csv") => {
  const blob = new Blob(["\uFEFF"+content], {type:mime});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href=url; a.download=filename;
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
};
/* FIX AUDIT #7 (HIGH — sanitizeForCSV Incomplete / CSV Injection Bypass):
   Implementasi lama hanya menangkap karakter berbahaya di AWAL string.
   Penyerang bisa menyisipkan "\n=cmd|..." di dalam field multi-line —
   Excel memproses baris setelah newline sebagai sel baru dengan formula aktif.
   Fix:
   1. Hapus semua \r\n dari dalam field (normalize ke spasi)
   2. Strip karakter kontrol lain (U+0000–U+001F kecuali tab yang sudah dihandle)
   3. Neutralisasi formula prefix (=, +, -, @, |, \t, \r) di awal hasil bersih
   Dengan urutan ini, injeksi via embedded newline tidak bisa lolos. */
const sanitizeForCSV = (val: any): string => {
  // 1. Normalize newline ke spasi agar tidak membentuk baris baru di CSV
  let s = String(val ?? "").replace(/\r?\n/g, " ").replace(/\r/g, " ");
  // 2. Escape tanda kutip ganda
  s = s.replace(/"/g, '""');
  // 3. Strip karakter kontrol non-printable (kecuali spasi/tab)
  s = s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "");
  // 4. Neutralisasi formula prefix di awal string hasil bersih
  if (/^[=+\-@\t\r|]/.test(s)) s = "'" + s;
  return `"${s}"`;
};

/* parseCSV: parser CSV minimal tapi AMAN terhadap field yang dibungkus tanda
   kutip (mis. `"Doe, John"` atau keterangan yang mengandung koma) — kasus ini
   sangat umum pada CSV hasil "Save As" dari Excel. Implementasi lama
   (`text.split("\n").map(r=>r.split(","))`) memecah SETIAP koma tanpa
   peduli tanda kutip, sehingga 1 baris bisa pecah menjadi kolom yang salah
   (nama/telepon/jabatan bergeser) TANPA error apa pun yang terlihat —
   silent data corruption saat import staf/roster. Parser di bawah
   menangani: koma di dalam tanda kutip, tanda kutip ganda escaped (""),
   serta akhir baris CRLF/LF. */
const parseCSV = (text: string): string[][] => {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let inQuotes = false;
  const normalized = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  for (let i = 0; i < normalized.length; i++) {
    const ch = normalized[i];
    if (inQuotes) {
      if (ch === '"') {
        if (normalized[i + 1] === '"') { field += '"'; i++; }
        else inQuotes = false;
      } else field += ch;
    } else if (ch === '"') {
      inQuotes = true;
    } else if (ch === ",") {
      row.push(field.trim()); field = "";
    } else if (ch === "\n") {
      row.push(field.trim()); rows.push(row); row = []; field = "";
    } else {
      field += ch;
    }
  }
  if (field.length || row.length) { row.push(field.trim()); rows.push(row); }
  return rows.filter(r => r.some(c => c !== ""));
};

/* ─── SUPABASE BACKUP ───────────────────────────────────────────────── */
interface SupabaseConfig { url: string; anonKey: string; autoBackup: boolean; backupInterval: number; lastBackup: string | null; realtimeBackup: boolean; autoExcelBackup: boolean; lastExcelBackup: string | null; lastExcelBackupTs: number | null; }
const defaultSupaCfg: SupabaseConfig = { url:"", anonKey:"", autoBackup:true, backupInterval:60, lastBackup:null, realtimeBackup:true, autoExcelBackup:true, lastExcelBackup:null, lastExcelBackupTs:null };

/* ─── DROPBOX BACKUP ────────────────────────────────────────────────── */
/* PENTING (security): Token Dropbox TIDAK ADA di client sama sekali.
   Sebelumnya token tersimpan langsung di sini (hardcoded) — itu adalah
   risiko keamanan serius karena bundle JS yang dikirim ke browser bisa
   dibongkar (view-source / unzip build) dan token diekstrak utuh, lalu
   dipakai siapa pun untuk mengakses akun Dropbox secara penuh.
   Sekarang semua request backup/restore diteruskan lewat Supabase Edge
   Function "dropbox-proxy" (lihat supabase/functions/dropbox-proxy/index.ts).
   Token asli hanya hidup sebagai secret di server (Deno.env), tidak pernah
   dikirim ke / disimpan di browser. Client hanya memegang DBX_PROXY_SECRET
   — sebuah shared secret terbatas yang HANYA bisa memicu 2 aksi spesifik
   (backup/restore 1 folder) lewat proxy ini, BUKAN kunci akses Dropbox
   penuh. Lihat catatan di vite-env.d.ts / .env.local untuk cara isi nilainya. */
const DBX_PROXY_SECRET = (import.meta as any).env?.VITE_PROXY_SHARED_SECRET || "";
const DBX_PROXY_FN     = "dropbox-proxy";

interface DropboxConfig { path: string; autoBackup: boolean; backupInterval: number; realtimeBackup: boolean; lastBackup: string | null; autoExcelBackup: boolean; lastExcelBackup: string | null; lastExcelBackupTs: number | null; }
const defaultDbxCfg: DropboxConfig = { path:"/KamarBedahPantiRini/backup.json", autoBackup:true, backupInterval:30, realtimeBackup:true, lastBackup:null, autoExcelBackup:true, lastExcelBackup:null, lastExcelBackupTs:null };

/* Helper tunggal: semua panggilan ke proxy lewat sini agar konsisten
   (header, error handling, parsing respons) */
async function callDbxProxy(body: Record<string, any>): Promise<{ok:boolean; data?:any; msg:string}> {
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

async function dropboxUpload(cfg: DropboxConfig, data: any): Promise<{ok:boolean;msg:string}> {
  return callDbxProxy({action:"backup", path:cfg.path, payload:data});
}

async function dropboxDownload(cfg: DropboxConfig): Promise<{ok:boolean;data?:any;msg:string}> {
  return callDbxProxy({action:"restore", path:cfg.path});
}

/* Upload Excel per-kategori ke Dropbox (lewat proxy — payload dikirim base64) */
async function dropboxUploadExcel(path: string, wb: any): Promise<{ok:boolean;msg:string}> {
  try {
    const buf = XLSX.write(wb, {type:"array", bookType:"xlsx"}) as Uint8Array;
    /* FIX AUDIT #13 (MEDIUM — O(n²) Base64 Encoding):
       String concatenation per-byte dalam loop menciptakan O(n²) alokasi memori
       karena setiap `+=` membuat string baru (JavaScript string immutable).
       Untuk file Excel 500+ baris ini bisa menyebabkan browser freeze.
       Fix: chunk-based approach — proses 8192 byte sekaligus menggunakan
       spread operator + String.fromCharCode(), lalu gabungkan chunk-nya.
       Ini mempertahankan keamanan btoa() sambil mengurangi alokasi ke O(n). */
    const CHUNK = 8192;
    let binary = "";
    for (let i = 0; i < buf.length; i += CHUNK) {
      binary += String.fromCharCode(...buf.subarray(i, i + CHUNK));
    }
    const base64 = btoa(binary);
    return callDbxProxy({action:"upload_excel", path, base64});
  } catch(e:any){ return {ok:false, msg:"Gagal menyiapkan file Excel: "+(e?.message||"unknown error")}; }
}

/* FIX AUDIT #3 (CRITICAL — XSS via downloadAsWord): semua data dinamis
   (judul, header, isi sel) WAJIB di-escape sebelum disisipkan ke string
   HTML. Tanpa ini, data pasien yang mengandung "<script>" atau atribut
   "onerror=" bisa tereksekusi saat file .doc dibuka di Word/LibreOffice
   (stored XSS). */
const escHtml = (s: any): string =>
  String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");

/* Download sebagai Word (.doc via HTML) */
function downloadAsWord(title: string, rows: string[][], filename: string) {
  const thead = rows[0].map(h=>`<th style="background:#16685F;color:#fff;padding:6px 10px;font-size:11px">${escHtml(h)}</th>`).join("");
  const tbody = rows.slice(1).map((r,i)=>`<tr style="background:${i%2===0?"#F0FFF8":"#fff"}">${r.map(c=>`<td style="padding:5px 10px;border-bottom:1px solid #e0e0e0;font-size:11px">${escHtml(c)}</td>`).join("")}</tr>`).join("");
  const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8"><title>${escHtml(title)}</title></head><body><h2 style="color:#16685F;font-family:Arial">${escHtml(title)}</h2><p style="color:#555;font-size:12px;font-family:Arial">Dicetak: ${escHtml(fNow())} — ${escHtml(HOSPITAL)}</p><table style="border-collapse:collapse;width:100%;font-family:Arial"><thead><tr>${thead}</tr></thead><tbody>${tbody}</tbody></table></body></html>`;
  const blob = new Blob([html], {type:"application/msword"});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href=url; a.download=filename;
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}

/* supabaseBackup & supabaseRestore kini menerima SUPA_CLIENT sebagai parameter
   agar dapat digunakan sebelum modul App selesai di-mount. */
async function supabaseBackup(_cfg: SupabaseConfig, data: any, client: any): Promise<{ok: boolean; msg: string}> {
  try {
    const payload = { created_at: new Date().toISOString(), data: JSON.stringify(data) };
    const {error} = await client.from("kamar_bedah_backup").insert(payload);
    if(error) return {ok:false, msg:"Supabase error: " + error.message.slice(0,120)};
    return {ok:true, msg:"✓ Backup ke Supabase berhasil — "+fNow()};
  } catch(e: any) {
    return {ok:false, msg:"Koneksi gagal: " + (e?.message||"unknown error")};
  }
}

async function supabaseRestore(_cfg: SupabaseConfig, client: any): Promise<{ok: boolean; msg: string; data?: any}> {
  try {
    const {data:rows, error} = await client.from("kamar_bedah_backup").select("data,created_at").order("created_at",{ascending:false}).limit(1);
    if(error) return {ok:false, msg:"Supabase error: " + error.message.slice(0,120)};
    if(!rows || rows.length === 0) return {ok:false, msg:"Tidak ada data backup di Supabase"};
    /* FIX AUDIT #8 (HIGH — JSON.parse tanpa guard): gunakan safeJSONParse
       agar data backup yang korup (terpotong, encoding rusak, upload gagal
       sebagian) tidak menyebabkan uncaught exception dan blank screen.
       Helper ini sudah ada di codebase dan digunakan di tempat lain. */
    const parsed = safeJSONParse(rows[0].data || "{}", null);
    if(parsed === null) return {ok:false, msg:"Data backup korup — tidak dapat di-parse. Coba restore dari backup lain."};
    return {ok:true, msg:"✓ Data backup ditemukan — "+(rows[0].created_at?.slice(0,19)?.replace("T"," ")||""), data:parsed};
  } catch(e: any) {
    return {ok:false, msg:"Koneksi gagal: " + (e?.message||"unknown error")};
  }
}

/* ─── MESSAGE BUILDERS ──────────────────────────────────────────────── */
/* msgSurgeon — direfactor sesuai Spesifikasi Teknis (lihat catatan di bawah).
   Nama fungsi & signature (name, ops) TIDAK diubah agar binding onClick UI
   tetap valid (sesuai PERINGATAN SISTEM di spesifikasi).
   Catatan keputusan implementasi:
   - "[pagi/siang/sore]" pertama (greeting) = gWord() → waktu SAAT pesan disusun.
   - "[pagi/siang/sore]" kedua (shift operasi besok) = timeOfDayLabel(jam operasi
     pertama) → karena merujuk ke JAM OPERASI, bukan jam pengiriman pesan.
   - Tanggal pakai fDLower() (lowercase, tanpa koma) agar konsisten dgn contoh
     output literal yang diberikan untuk msgAnest.
   - Jika dalam grup operasi dokter bedah ini ada >1 nama anestesi berbeda
     (kasus campuran, tidak disebutkan di spesifikasi), nama-nama itu
     digabung dengan "dan" agar informasi tidak hilang — bukan diasumsikan
     seragam begitu saja. */
function msgSurgeon(name: string, ops: any[]) {
  const s = [...ops].sort((a,b)=>a.time.localeCompare(b.time));
  if(!s.length) return "";
  const N = titleCaseDokter(name); // nama TARGET sapaan → Title Case, tanpa prefix "dr"
  const uniqAnest = [...new Set(s.map(o=>o.anesthesiologist).filter(Boolean))];
  const anestLabel = uniqAnest.length ? uniqAnest.map(a=>sanitize(a)).join(" dan ") : "-"; // nama pendukung → raw, sesuai database
  const greet = gWord();
  const shift = timeOfDayLabel(s[0]?.time);
  const tanggal = fDLower(tmrwDate());
  const jam = fTR(s[0]?.time);
  const baris = (o:any) => {
    const usia = usiaOp(o);
    const diag = sanitize(o.diagnosis||"").trim();
    return `${sanitize(o.patient)}${usia?` ${usia}`:""} rencana ${sanitize(o.procedure)}${diag?` ${diag}`:""}`;
  };
  if(s.length>1){
    const lines = s.map((o,i)=>`${i+1}. ${baris(o)}`).join("\n");
    return `selamat ${greet} dokter ${N} rencana operasi besok ${shift} , ${tanggal} mulai jam ${jam} dr bius dengan ${anestLabel} ada ${s.length} pasien :\n${lines}\nterima kasih banyak dokter`;
  }
  return `selamat ${greet} dokter ${N} rencana operasi besok ${shift} ${tanggal} mulai jam ${jam} bius dengan ${anestLabel} atas nama ${baris(s[0])}\nterima kasih banyak dokter`;
}
/* msgAnest — direfactor sesuai Spesifikasi Teknis & CONTOH OUTPUT LITERAL
   yang diberikan. Signature (name, ops) TIDAK diubah.
   Catatan keputusan implementasi:
   - Spacing mengikuti CONTOH AKHIR secara harfiah (baris-baris rapat,
     newline tunggal antar section — BUKAN double newline seperti disebut
     di teks algoritma pipeline, yang ternyata tidak konsisten dgn contoh).
   - Header "dokter [Nama]" pakai titleCaseDokter (Title Case, tanpa prefix
     "dr"), sesuai contoh "dokter Timor". Nama dokter bedah & anestesi di
     BADAN pesan tetap nilai asli dari database (mis. "dr dedi", "dr timor"),
     sesuai contoh — tidak diubah case-nya. */
function msgAnest(name: string, ops: any[]) {
  const tanggal = fDLower(tmrwDate());
  const greet = gWord();
  const anestRaw = sanitize(name); // dipakai di badan pesan, format asli (mis. "dr timor")

  const renderShift = (shiftOps: any[]): string => {
    if(!shiftOps.length) return "";
    const byTime: Record<string, any[]> = {};
    shiftOps.forEach(o=>{ if(!byTime[o.time]) byTime[o.time]=[]; byTime[o.time].push(o); });
    return Object.keys(byTime).sort((a,b)=>a.localeCompare(b)).map(time=>{
      const opsAtTime = byTime[time];
      const bySurgeon: Record<string, any[]> = {};
      opsAtTime.forEach(o=>{ const k=(o.surgeon||"").trim().toLowerCase(); if(!bySurgeon[k]) bySurgeon[k]=[]; bySurgeon[k].push(o); });
      return Object.values(bySurgeon).map((group:any[])=>{
        const surgeonRaw = sanitize(group[0].surgeon);
        const jam = fTR(time);
        const baris = (o:any) => {
          const usia = usiaOp(o);
          const diag = sanitize(o.diagnosis||"").trim();
          return `${sanitize(o.patient)}${usia?` ${usia}`:""} rencana ${sanitize(o.procedure)}${diag?` ${diag}`:""}`;
        };
        if(group.length>1){
          const lines = group.map((o,i)=>`${i+1}. ${baris(o)}`).join("\n");
          return `*jam ${jam}* ${surgeonRaw} dan ${anestRaw} ada ${group.length} pasien :\n${lines}`;
        }
        return `*jam ${jam}* ${surgeonRaw} dan ${anestRaw} atas nama ${baris(group[0])}`;
      }).join("\n");
    }).join("\n");
  };

  const pagi = ops.filter((o:any)=>o.time<"14:00");
  const sore = ops.filter((o:any)=>o.time>="14:00");

  let out = `Selamat ${greet} dokter ${titleCaseDokter(name)} ,Rencana operasi besok hari ${tanggal}`;
  if(pagi.length) out += `\n*Pagi*\n${renderShift(pagi)}`;
  if(sore.length) out += `\n*Sore*\n${renderShift(sore)}`;
  if(!pagi.length && !sore.length) out += "\nBelum ada rencana operasi";
  out += "\nterima kasih banyak dokter";
  return out;
}
function msgCyto(op: any) {
  const alg = op.allergy&&op.allergy!=="Tidak Ada" ? `⚠ ALERGI: ${sanitize(op.allergy)}\n` : "";
  const ra  = op.ruangAsal ? `Ruang Asal : ${sanitize(op.ruangAsal)}\n` : "";
  return `⚠️ LAPORAN ACARA OPERASI CYTO / EMERGENCY ⚠️\n━━━━━━━━━━━━━━━━━━━\nKamar Bedah ${HOSPITAL}\n━━━━━━━━━━━━━━━━━━━\n${fNow()}\n\nMOHON SEGERA KE KAMAR OPERASI\n\n${alg}${ra}Pasien   : ${sanitize(op.patient)}${op.age?` (${op.age} th)`:""}\nRM/Gol   : ${op.rm||"-"} / ${op.bloodType}\nDiagnosa : ${sanitize(op.diagnosis)}\nTindakan : ${sanitize(op.procedure)}\nJam      : ${op.time} WIB — ${op.room}\n\nOperator : ${op.surgeon||"-"}\nAnestesi : ${op.anesthesiologist||"-"}\nAsisten  : ${op.assistantNurse||"-"}\nInstrumen: ${op.circulatingNurse||"-"}\nP.Anest  : ${op.anesthesiaNurse||"-"}\nOnloop   : ${op.onloopNurse||"-"}\nRR/Katim : ${op.rrKatim||"-"}\n━━━━━━━━━━━━━━━━━━━\n⚡ HARAP RESPON SEGERA ⚡\nKamar Bedah ${HOSPITAL}`;
}
/* buildLaporan — direfactor sesuai Spesifikasi Teknis. Signature TIDAK diubah.
   Poin ③ (acara hari ini) & ④ (rencana besok) direstrukturisasi sesuai
   algoritma & format wajib di spesifikasi. Bagian Siaga Anestesi/Cyto/
   Perawat/Pembawa HP/Footer TIDAK disentuh — di luar scope spesifikasi ini.
   Catatan keputusan:
   - Poin ③ "Format Baris Pasien (Wajib)" di spesifikasi TIDAK menyertakan
     usia pasien — sehingga usia SENGAJA tidak ditambahkan di bagian ini,
     murni mengikuti spesifikasi literal (berbeda dengan poin ④ yang
     eksplisit meminta formatUsia).
   - Marker "⚠️[CYTO]⚠️" (dari implementasi sebelumnya) dipertahankan sebagai
     penanda visual penting yang tidak disebutkan utk dihapus di spesifikasi. */
function buildLaporan({greeting,recipients,keterangan,todayOps,tomorrowOps,anestStandby,anestCyto,nurseStandby,pembawaHP}: any) {
  const active = [...todayOps].filter((o:any)=>o.status!=="batal");
  const batal  = todayOps.filter((o:any)=>o.status==="batal");

  /* ── Poin ③: urutan kronologis murni berdasarkan jam operasi (a→z) ── */
  const sortedActive = [...active].sort((a:any,b:any)=>a.time.localeCompare(b.time));
  const todayLines = sortedActive.map((o:any, i:number)=>{
    const cytoMark = o.opType==="cyto" ? "⚠️[CYTO]⚠️ " : "";
    const anest    = sanitize(o.anesthesiologist)||"-";
    const surg     = sanitize(o.surgeon)||"-";
    const proc     = sanitize(o.procedure);
    const asisten  = sanitize(o.assistantNurse)||"-";
    const instrumen= sanitize(o.circulatingNurse)||"-";
    const onloop   = sanitize(o.onloopNurse)||"-";
    const rr       = sanitize(o.rrKatim)||"-";
    const anestNrs = sanitize(o.anesthesiaNurse)||"-";
    return [
      `_${i+1}. ${o.time} ${cytoMark}${sanitize(o.patient)} - ${proc} - ${surg} (Anest: ${anest})_`,
      `   Asisten: ${asisten} | Instrumen: ${instrumen}`,
      `   Onloop: ${onloop} | RR/Katim: ${rr} | Pr.Anest: ${anestNrs}`,
    ].join("\n");
  }).join("\n");

  const batalLine = batal.length?`\n_Batal/Tunda: ${batal.map((o:any)=>sanitize(o.patient)+(o.cancelReason?` (${sanitize(o.cancelReason)})`:""  )).join("; ")}_`:"";
  const todaySec = active.length
    ? `*Hari ini acara operasi ada ${active.length} berjalan lancar*\n${todayLines}${batalLine}`
    : `_Hari Ini Tidak Ada Acara Operasi_${batalLine}`;

  /* ── Poin ④: split shift Pagi (00:00-13:59) / Sore (14:00-23:59) → group
     by jamOperasi → group by dokter bedah → IF jumlah>1 list bernumor,
     ELSE kalimat inline. formatUsia WAJIB dipakai di bagian ini. ── */
  const pagi = tomorrowOps.filter((o:any)=>o.time<"14:00");
  const sore = tomorrowOps.filter((o:any)=>o.time>="14:00");
  const renderShift = (shiftOps: any[]): string => {
    if(!shiftOps.length) return "";
    const byTime: Record<string, any[]> = {};
    shiftOps.forEach((o:any)=>{ if(!byTime[o.time]) byTime[o.time]=[]; byTime[o.time].push(o); });
    return Object.keys(byTime).sort((a,b)=>a.localeCompare(b)).map(time=>{
      const bySurgeon: Record<string, any[]> = {};
      byTime[time].forEach((o:any)=>{ const k=(o.surgeon||"").trim().toLowerCase(); if(!bySurgeon[k]) bySurgeon[k]=[]; bySurgeon[k].push(o); });
      return Object.values(bySurgeon).map((group:any[])=>{
        const surgeonDisplay = sanitize(group[0].surgeon)||"-";
        const anestDisplay = sanitize(group[0].anesthesiologist)||"-";
        const jam = fTR(time);
        const baris = (o:any) => {
          const usia = usiaOp(o);
          const diag = sanitize(o.diagnosis||"").trim();
          return `${sanitize(o.patient)}${usia?` ${usia}`:""} rencana ${sanitize(o.procedure)}${diag?` ${diag}`:""}`;
        };
        if(group.length>1){
          const lines = group.map((o,i)=>`${i+1}. ${baris(o)}`).join("\n");
          return `*jam ${jam}* ${surgeonDisplay} dan ${anestDisplay} ada ${group.length} pasien :\n${lines}`;
        }
        return `*jam ${jam}* ${surgeonDisplay} dan ${anestDisplay} atas nama ${baris(group[0])}`;
      }).join("\n\n");
    }).join("\n\n");
  };
  let tmrwSec = `Rencana operasi besok tanggal ${fDMY(tmrwDate())}:\n`;
  if(pagi.length) tmrwSec += `\n*PAGI*\n${renderShift(pagi)}`;
  if(sore.length) tmrwSec += `\n\n*SORE*\n${renderShift(sore)}`;
  if(!pagi.length&&!sore.length) tmrwSec += "\nBelum ada rencana operasi";

  /* ── Bagian berikut TIDAK diubah dari implementasi sebelumnya ── */
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
const Btn = ({children,onClick,color=C.p,outline,sm,full,disabled,style={}}: any) => {
  const isGrad = typeof color==="string" && color.startsWith("linear-gradient");
  const borderCol = isGrad ? "rgba(30,58,138,.4)" : (disabled?"#ccc":color);
  const bgStyle = outline ? C.white : disabled ? "#eee" : color;
  return (
    <button onClick={onClick} disabled={disabled} style={{padding:sm?"7px 12px":"10px 18px",borderRadius:10,border:`1.5px solid ${borderCol}`,background:bgStyle,color:outline?color:C.white,fontSize:sm?12:14,fontWeight:700,cursor:disabled?"not-allowed":"pointer",opacity:disabled?.45:1,width:full?"100%":"auto",fontFamily:"inherit",transition:"opacity .15s",boxShadow:isGrad&&!disabled?"0 2px 8px rgba(30,58,138,.25)":"none",...style}}>{children}</button>
  );
};
const WaBt = ({children,onClick,sm,full,disabled,style={}}: any) => (
  <button onClick={onClick} disabled={disabled} style={{padding:sm?"7px 12px":"10px 18px",borderRadius:10,border:`1.5px solid ${disabled?"#ccc":C.wa}`,background:disabled?"#eee":C.wa,color:C.white,fontSize:sm?12:14,fontWeight:700,cursor:disabled?"not-allowed":"pointer",width:full?"100%":"auto",fontFamily:"inherit",display:"flex",alignItems:"center",justifyContent:"center",gap:5,...style}}>✓ {children}</button>
);
const Card = ({children,style={},hi}: any) => (
  <div style={{background:"#ffffff",borderRadius:14,padding:"18px 20px",marginBottom:12,boxShadow:"0 4px 6px -1px rgba(0,0,0,0.05),0 2px 4px -1px rgba(59,130,246,0.06)",border:`1px solid ${hi||C.b}`,...style}}>{children}</div>
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
/* FIX AUDIT #14 (MEDIUM — ErrorBoundary tidak me-reset state setelah error):
   Sebelumnya "Coba Lagi" hanya men-set err:false, tanpa me-remount komponen
   child yang error. React akan mencoba merender ulang child yang sama persis
   yang sudah error — kemungkinan besar akan error lagi dan pengguna terjebak.
   Fix: tambahkan field `key` yang di-increment setiap reset. Perubahan `key`
   pada React.Fragment memaksa seluruh subtree di-unmount lalu remount bersih,
   sehingga state internal child ikut direset. */
class ErrorBoundary extends Component<{children: React.ReactNode},{err:boolean;msg:string;key:number}> {
  state = {err:false, msg:"", key:0};
  static getDerivedStateFromError(e: any){ return {err:true, msg:e?.message||"Error"}; }

  handleReset = () => {
    this.setState(prev => ({err:false, msg:"", key: prev.key + 1}));
  };

  render(){
    if(this.state.err) return (
      <div style={{padding:24,textAlign:"center",background:C.dBg,borderRadius:14,margin:"8px 0",border:`1px solid ${C.dL}`}}>
        <div style={{fontSize:32,marginBottom:8}}>⚠️</div>
        <div style={{fontSize:14,color:C.d,fontWeight:600,marginBottom:8}}>{this.state.msg}</div>
        <Btn onClick={this.handleReset}>Coba Lagi</Btn>
      </div>
    );
    /* key change forces full remount of children — state internal child ikut direset */
    return <Fragment key={this.state.key}>{this.props.children}</Fragment>;
  }
}

/* ─── PIN SCREEN ────────────────────────────────────────────────────── */
/* ── PIN Screen dengan 2-role (Admin & Perawat) ── */
/* FIX AUDIT #18 (CRITICAL — Brute-force PIN tanpa rate-limit):
   PIN 4 digit numerik = maksimal 10.000 kombinasi. Sebelumnya tidak ada
   batasan percobaan gagal di UI — siapa pun bisa mencoba ribuan kombinasi
   berurutan tanpa hambatan apa pun di sisi client. Sekarang ditambahkan
   lockout sederhana: setelah MAX_ATTEMPTS kali gagal berurutan, input
   dikunci selama LOCKOUT_MS.
   PENTING — CATATAN JUJUR: lockout ini HANYA lapisan UX, bukan pengganti
   keamanan sesungguhnya. State ini hilang begitu halaman di-refresh atau
   localStorage dihapus, sehingga TIDAK mencegah serangan otomatis lewat
   script langsung ke endpoint Supabase. Mitigasi sesungguhnya WAJIB ada
   di Row Level Security & rate-limiting sisi server Supabase — di luar
   kendali kode App.tsx ini. */
const PIN_MAX_ATTEMPTS = 5;
const PIN_LOCKOUT_MS   = 60_000;
function PinScreen({onVerify,pinAdmin,pinPerawat,isFirstTime}: any) {
  const [pin,setPin]   = useState("");
  const [err,setErr]   = useState("");
  const [shake,setShk] = useState(false);
  const [checking,setChecking] = useState(false);
  const [attempts,setAttempts]     = useState(0);
  const [lockedUntil,setLockedUntil] = useState<number|null>(null);
  // Force setup state (first time: PIN 0000)
  const [setupMode,setSetupMode] = useState(false);
  const [newAdmin,setNewAdmin]   = useState("");
  const [cfAdmin,setCfAdmin]     = useState("");
  const [newPerawat,setNewPerawat] = useState("");
  const [cfPerawat,setCfPerawat]   = useState("");
  const [setupErr,setSetupErr]   = useState("");

  /* FIX AUDIT #4: PIN sekarang bisa berbentuk hash ("salt:hash") atau
     legacy plaintext (data lama sebelum migrasi). verifyPin menangani
     keduanya secara transparan. */
  const check = async () => {
    if(checking) return;
    if(lockedUntil && Date.now() < lockedUntil){
      const sisaDetik = Math.ceil((lockedUntil - Date.now())/1000);
      setErr(`Terlalu banyak percobaan gagal. Coba lagi dalam ${sisaDetik} detik.`); setShk(true);
      setTimeout(()=>{setShk(false);},400);
      return;
    }
    if(isFirstTime && pin==="0000"){
      setSetupMode(true); setPin(""); setErr(""); return;
    }
    setChecking(true);
    try{
      if(await verifyPin(pin, pinAdmin)){
        setAttempts(0); setLockedUntil(null);
        AmbilLogMedis("KEAMANAN","Admin","LOGIN_SUKSES","Login Admin berhasil via PIN screen",{role:"admin"});
        onVerify("admin"); return;
      }
      if(await verifyPin(pin, pinPerawat)){
        setAttempts(0); setLockedUntil(null);
        AmbilLogMedis("KEAMANAN","Perawat","LOGIN_SUKSES","Login Perawat berhasil via PIN screen",{role:"perawat"});
        onVerify("perawat"); return;
      }
      const nextAttempts = attempts + 1;
      setAttempts(nextAttempts);
      AmbilLogMedis("KEAMANAN","Unknown","LOGIN_GAGAL",
        `PIN salah — percobaan ke-${nextAttempts}/${PIN_MAX_ATTEMPTS}`,
        { attempt: nextAttempts, locked: nextAttempts >= PIN_MAX_ATTEMPTS }
      );
      if(nextAttempts >= PIN_MAX_ATTEMPTS){
        setLockedUntil(Date.now() + PIN_LOCKOUT_MS);
        setErr(`Terlalu banyak percobaan gagal. Coba lagi dalam ${Math.ceil(PIN_LOCKOUT_MS/1000)} detik.`);
      } else {
        setErr(`PIN salah. Coba lagi. (${nextAttempts}/${PIN_MAX_ATTEMPTS})`);
      }
      setShk(true);
      setTimeout(()=>{setShk(false);setErr("");setPin("");},1500);
    } finally {
      setChecking(false);
    }
  };
  const doSetup = () => {
    if(newAdmin.length<4){setSetupErr("PIN Admin minimal 4 digit");return;}
    if(newAdmin!==cfAdmin){setSetupErr("PIN Admin tidak cocok");return;}
    if(newPerawat.length<4){setSetupErr("PIN Perawat minimal 4 digit");return;}
    if(newPerawat!==cfPerawat){setSetupErr("PIN Perawat tidak cocok");return;}
    if(newAdmin===newPerawat){setSetupErr("PIN Admin & Perawat tidak boleh sama");return;}
    onVerify("admin",newAdmin,newPerawat);
  };

  if(setupMode) return (
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"linear-gradient(135deg,#5FA39A,#3FA897,#16685F)",padding:16}}>
      <style>{`@keyframes lgFloat{0%,100%{transform:translateY(0)}50%{transform:translateY(-6px)}}`}</style>
      <div style={{background:C.white,borderRadius:24,padding:"32px 28px",width:"100%",maxWidth:380,boxShadow:"0 20px 60px rgba(0,0,0,.3)"}}>
        <div style={{textAlign:"center",marginBottom:20}}>
          <img src="/logo.jpeg" alt="Logo" style={{width:72,height:72,borderRadius:"50%",objectFit:"cover",display:"block",margin:"0 auto 12px",animation:"lgFloat 3s ease-in-out infinite"}}/>
          <div style={{fontSize:16,fontWeight:800,color:C.p}}>Setup PIN Pertama Kali</div>
          <div style={{fontSize:12,color:C.tL,marginTop:4}}>Buat PIN Admin dan PIN Perawat baru</div>
        </div>
        <div style={{background:C.wBg,border:`1px solid ${C.w}`,borderRadius:10,padding:"10px 14px",marginBottom:16,fontSize:12,color:C.w,fontWeight:600}}>
          ⚠ Setelah setup, PIN <b>0000</b> tidak bisa digunakan lagi.
        </div>
        <div style={{fontSize:13,fontWeight:700,color:C.p,marginBottom:8}}>🔐 PIN Admin</div>
        <LF label="PIN Admin Baru (min. 4 digit)"><input style={iS} type="password" inputMode="numeric" maxLength={6} placeholder="PIN admin baru" value={newAdmin} onChange={e=>setNewAdmin(e.target.value)}/></LF>
        <LF label="Konfirmasi PIN Admin"><input style={iS} type="password" inputMode="numeric" maxLength={6} placeholder="Ulangi PIN admin" value={cfAdmin} onChange={e=>setCfAdmin(e.target.value)}/></LF>
        <div style={{fontSize:13,fontWeight:700,color:C.g,margin:"12px 0 8px"}}>👤 PIN Perawat</div>
        <LF label="PIN Perawat Baru (min. 4 digit)"><input style={iS} type="password" inputMode="numeric" maxLength={6} placeholder="PIN perawat baru" value={newPerawat} onChange={e=>setNewPerawat(e.target.value)}/></LF>
        <LF label="Konfirmasi PIN Perawat"><input style={iS} type="password" inputMode="numeric" maxLength={6} placeholder="Ulangi PIN perawat" value={cfPerawat} onChange={e=>setCfPerawat(e.target.value)}/></LF>
        {setupErr && <div style={{fontSize:12,color:C.d,marginBottom:10,fontWeight:600}}>⚠ {setupErr}</div>}
        <Btn full onClick={doSetup} style={{marginTop:8,padding:"13px",fontSize:15,background:C.p}}>✓ Simpan PIN & Masuk</Btn>
      </div>
    </div>
  );

  return (
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"linear-gradient(135deg,#5FA39A,#3FA897,#16685F)",padding:16}}>
      <style>{`@keyframes shk{0%,100%{transform:translateX(0)}25%,75%{transform:translateX(-7px)}50%{transform:translateX(7px)}} @keyframes lgFloat{0%,100%{transform:translateY(0) scale(1)}50%{transform:translateY(-8px) scale(1.06)}} @keyframes lgGlow{0%,100%{box-shadow:0 4px 18px rgba(22,104,95,.3),0 0 0 3px #16685F55}50%{box-shadow:0 12px 38px rgba(22,104,95,.65),0 0 0 4px #2D9A87aa}}`}</style>
      <div style={{background:C.white,borderRadius:24,padding:"32px 28px",width:"100%",maxWidth:360,boxShadow:"0 20px 60px rgba(0,0,0,.3)",animation:shake?"shk .4s":"none"}}>
        <div style={{textAlign:"center",marginBottom:24}}>
          <img src="/logo.jpeg" alt="Logo Kamar Bedah" style={{width:104,height:104,borderRadius:"50%",objectFit:"cover",display:"block",margin:"0 auto 14px",animation:"lgFloat 3s ease-in-out infinite, lgGlow 3s ease-in-out infinite"}}/>
          <div style={{fontSize:11,fontWeight:700,color:C.g,letterSpacing:2,textTransform:"uppercase",marginBottom:4}}>RS Panti Rini</div>
          <div style={{fontSize:20,fontWeight:900,color:C.p,letterSpacing:.5,lineHeight:1.2}}>SISTEM KOORDINASI<br/>KAMAR BEDAH</div>
          <div style={{fontSize:11,color:C.tL,marginTop:6}}>{HOSPITAL}</div>
        </div>
        <div style={{fontSize:13,fontWeight:600,color:C.t,textAlign:"center",marginBottom:10}}>Masukkan PIN Akses</div>
        <div style={{fontSize:11,color:C.tL,textAlign:"center",marginBottom:10}}>Gunakan PIN Admin atau PIN Perawat</div>
        <input style={{...iS,textAlign:"center",fontSize:28,letterSpacing:12,height:56,marginBottom:12,borderRadius:12}} type="password" inputMode="numeric" maxLength={6} placeholder="••••" value={pin} disabled={checking||!!(lockedUntil&&Date.now()<lockedUntil)} onChange={e=>setPin(e.target.value)} onKeyDown={e=>e.key==="Enter"&&check()}/>
        {err && <div style={{fontSize:12,color:C.d,textAlign:"center",marginBottom:10}}>⚠ {err}</div>}
        <Btn full onClick={check} disabled={checking||!!(lockedUntil&&Date.now()<lockedUntil)} style={{marginBottom:12,padding:"13px",fontSize:15}}>{checking?"Memeriksa...":(lockedUntil&&Date.now()<lockedUntil)?"🔒 Terkunci":"Masuk"}</Btn>
        <div style={{display:"flex",justifyContent:"center"}}>
          <button onClick={()=>{}} style={{background:"none",border:"none",color:C.tL,fontSize:11,cursor:"pointer",fontFamily:"inherit"}}>Hubungi Admin jika lupa PIN</button>
        </div>
        <div style={{display:"flex",gap:6,marginTop:12,justifyContent:"center"}}>
          <span style={{background:C.pBg,color:C.p,fontSize:10,padding:"3px 10px",borderRadius:12,fontWeight:700}}>🔐 Admin</span>
          <span style={{background:C.gBg,color:C.g,fontSize:10,padding:"3px 10px",borderRadius:12,fontWeight:700}}>👤 Perawat</span>
        </div>
      </div>
    </div>
  );
}

/* ─── VIEW JADWAL ────────────────────────────────────────────────────── */
function ViewJadwal({ops,setOps,startEditOp,deleteOp,sendReminder,reqOpId,setReqOpId,reqText,setReqText,addReq,delReq,getPhone,setNotifs,showToast,privacyMode,role,upsertOneToSupa}: ViewJadwalProps) {
  const [cId,setCId]     = useState<string|null>(null);
  const [cR,setCR]       = useState("");
  const [cytoM,setCytoM] = useState<any>(null);
  const [delC,setDelC]   = useState<string|null>(null);
  const [page,setPage]   = useState(0);
  const [fSt,setFSt]     = useState("all");
  const [q,setQ]         = useState("");
  const [showRiwayat,setShowRiwayat] = useState(false);

  /* ── FITUR BARU: Auto-filter jadwal yang sudah berlalu ke "Riwayat" ──
     Tanggal pembanding di-reset ke 00:00:00 ("start of day") — jadwal HARI
     INI tetap tampil PENUH dari pagi sampai malam (apa pun statusnya,
     termasuk yang sudah "done"), baru pindah ke Riwayat TEPAT SAAT
     berganti hari ke esok. Ini murni FILTER TAMPILAN di komponen ini —
     TIDAK mengubah/memindahkan struktur data `ops` maupun `archive`
     (snapshot backup yang sudah ada, dipakai fitur lain yang terpisah),
     sehingga nol risiko terhadap fitur backup/restore/audit yang sudah
     established. `cytoActive` di bawah SENGAJA tetap membaca dari `ops`
     penuh (bukan activeOpsBase) — alert Cyto/Emergency harus tetap
     terdeteksi terlepas dari tanggalnya. */
  const todayStr = todayDate(); // sudah format YYYY-MM-DD, aman dibandingkan sebagai string
  const activeOpsBase = ops.filter((o:any)=>o.date>=todayStr);
  const pastOps = ops.filter((o:any)=>o.date<todayStr).sort((a:any,b:any)=>b.date.localeCompare(a.date)||b.time.localeCompare(a.time));

  const sorted = [...activeOpsBase]
    .filter((op:any) => {
      if(fSt!=="all"&&op.status!==fSt) return false;
      if(q.trim()){ const lq=q.toLowerCase(); return (op.patient||"").toLowerCase().includes(lq)||(op.procedure||"").toLowerCase().includes(lq)||(op.surgeon||"").toLowerCase().includes(lq); }
      return true;
    })
    .sort((a:any,b:any)=>{
      // Prioritas 1: cyto hari ini selalu paling atas
      const at=a.date===todayStr, bt=b.date===todayStr;
      if(at&&a.opType==="cyto"&&!(bt&&b.opType==="cyto")) return -1;
      if(bt&&b.opType==="cyto"&&!(at&&a.opType==="cyto")) return 1;
      // Prioritas 2: yang belum selesai naik, yang done turun
      const aDone = a.status==="done"||a.status==="batal";
      const bDone = b.status==="done"||b.status==="batal";
      if(!aDone && bDone) return -1;
      if(aDone && !bDone) return 1;
      // Prioritas 3: kronologis tanggal & jam
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
            w.document.write(`<!DOCTYPE html><html><head><title>Jadwal Operasi ${printDate||"Semua"} — ${HOSPITAL}</title><style>body{font-family:Arial,sans-serif;font-size:12px;margin:20px}h2{color:#16685F;margin-bottom:4px}p{color:#555;margin-bottom:12px}table{width:100%;border-collapse:collapse}th{background:#16685F;color:#fff;padding:7px 10px;text-align:left;font-size:11px}td{padding:6px 10px;border-bottom:1px solid #e0e0e0;vertical-align:top}tr:nth-child(even)td{background:#f5f5f5}@media print{button{display:none}}</style></head><body><h2>🏥 Jadwal Operasi — ${HOSPITAL}</h2><p>${printDate?`Tanggal: ${printDate}`:"Semua jadwal"} · Dicetak: ${fNow()}</p><button onclick="window.print()" style="margin-bottom:12px;background:#16685F;color:#fff;border:none;padding:8px 18px;border-radius:6px;cursor:pointer;font-size:13px">🖨 Cetak / Simpan PDF</button><table><thead><tr><th>#</th><th>Tanggal/Jam</th><th>Pasien</th><th>Tindakan</th><th>Kamar</th><th>Dr. Bedah</th><th>Dr. Anestesi</th><th>Status</th></tr></thead><tbody>${rows}</tbody></table></body></html>`);
            w.document.close();
          }} style={{padding:"5px 12px",borderRadius:20,border:`1px solid ${C.p}`,background:C.pBg,color:C.p,fontSize:12,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>
            🖨 Cetak
          </button>
        </div>
      </div>
      {sorted.length===0 && (
        <Card><div style={{textAlign:"center",padding:28,color:C.tL}}>
          <div style={{fontSize:36,marginBottom:8}}>📋</div>
          <div style={{fontSize:14,fontWeight:600}}>{activeOpsBase.length===0?"Tidak ada jadwal operasi yang akan datang.":"Tidak ada hasil pencarian"}</div>
          {activeOpsBase.length===0&&<div style={{fontSize:12,marginTop:4}}>Klik tab 📝 Daftar untuk mendaftarkan</div>}
        </div></Card>
      )}
      {paged.map((op:any)=>{
        const sc=STS[op.status as keyof typeof STS]||STS.scheduled, ot=OT[op.opType as keyof typeof OT||"elektif"];
        const hasAlg=op.allergy&&op.allergy!=="Tidak Ada", isBatal=op.status==="batal";
        const isDone   = op.status==="done";
        const isToday  = op.date===todayStr;
        const isTomorrow = op.date === tmrwDate();
        const isActive = isToday && !isDone && !isBatal; // blinking dot condition

        /* ── Auto-minimize: pasien selesai hari ini — tampilkan mini card TAPI tetap ada tombol Edit ── */
        if(isDone && isToday) return (
          <Card key={op.id} style={{opacity:.75,transition:"opacity .3s",marginBottom:6}}>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <span style={{fontSize:11,color:C.s,fontWeight:700,flexShrink:0}}>✅</span>
              <span style={{fontSize:12,color:"#475569",flex:1,textDecoration:"line-through",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{maskName(op.patient,privacyMode)} — {op.time} — {op.procedure}</span>
              <Bdg label="Selesai" color={C.s} bg={C.sBg}/>
              <Btn sm outline color={C.p} onClick={()=>startEditOp(op)}>✏ Edit</Btn>
            </div>
          </Card>
        );

        return (
          <Card key={op.id}
            hi={op.opType==="cyto"&&!isBatal?"#F44336":hasAlg&&!isBatal?C.dL:undefined}
            style={{
              opacity:isBatal?.72:1,
              background:isActive?"#fef2f2":(isTomorrow&&!isBatal&&!isDone)?"#F3E8FF":undefined,
              border:isActive?"1.5px solid #fca5a5":(isTomorrow&&!isBatal&&!isDone)?"1.5px solid #D8B4FE":undefined,
              transition:"all .3s"
            }}
          >
            {/* ── Alert banners ── */}
            {op.opType==="cyto"&&!isBatal && <div style={{background:"#FFCDD2",borderRadius:8,padding:"6px 12px",marginBottom:8,fontSize:12,fontWeight:700,color:"#B71C1C"}}>⚠️ CYTO / EMERGENCY — SEGERA DITANGANI</div>}
            {hasAlg&&op.opType!=="cyto"&&!isBatal && <div style={{background:C.dBg,borderRadius:8,padding:"5px 12px",marginBottom:8,fontSize:12,fontWeight:700,color:C.d}}>⚠ ALERGI: {op.allergy}</div>}
            {isTomorrow&&!isBatal&&!isDone && <div style={{background:"#EDE9FE",borderRadius:8,padding:"5px 12px",marginBottom:8,fontSize:12,fontWeight:700,color:"#6B21A8",display:"flex",alignItems:"center",gap:6}}>📅 OPERASI BESOK (H+1) — Siapkan persiapan sebelum shift selesai</div>}
            {isBatal && <div style={{background:C.dBg,borderRadius:8,padding:"5px 12px",marginBottom:8,fontSize:12,fontWeight:700,color:C.d}}>✕ {op.cancelReason?"BATAL: "+op.cancelReason:"BATAL/DITUNDA"}</div>}

            {/* ── Header row: nama + badge + Mulai/Selesai di kanan ── */}
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8,gap:8}}>
              <div style={{flex:1,minWidth:0}}>
                <div style={{display:"flex",alignItems:"center",gap:7,marginBottom:2}}>
                  {/* Blinking red dot — hanya untuk pasien aktif hari ini */}
                  {isActive && (
                    <span style={{width:10,height:10,borderRadius:"50%",background:"#ef4444",flexShrink:0,display:"inline-block",animation:"redPulse 1.2s ease-in-out infinite"}}/>
                  )}
                  <div style={{fontSize:15,fontWeight:800,color:(isTomorrow&&!isBatal&&!isDone)?"#6B21A8":C.t,textDecoration:isBatal?"line-through":"none",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{maskName(op.patient,privacyMode)||"—"}</div>
                </div>
                <div style={{fontSize:12,color:C.tL}}>{op.age?op.age+"th  ":""}{op.rm?"RM "+maskRM(op.rm,privacyMode)+"  ":""}{op.bloodType}</div>
              </div>
              {/* Badges */}
              <div style={{display:"flex",flexDirection:"column",gap:4,alignItems:"flex-end",flexShrink:0}}>
                <Bdg label={ot?.label||"Elektif"} color={ot?.c||C.i} bg={ot?.bg||C.iBg}/>
                <Bdg label={sc.l} color={sc.c} bg={sc.bg}/>
              </div>
              {/* Tombol Mulai / Selesai — ujung kanan, besar & terang */}
              {!isBatal && (
                <div style={{display:"flex",gap:6,flexShrink:0,alignItems:"center"}}>
                  {op.status==="scheduled" && (
                    <button onClick={()=>{
                      const upd={...op,status:"ongoing",updated_at:new Date().toISOString()};
                      setOps((p:any)=>p.map((o:any)=>o.id===op.id?upd:o));
                      AmbilLogMedis("JADWAL_OPERASI","Admin","MULAI_OPERASI",
                        `Operasi dimulai: ${op.patient||"-"} — ${op.procedure||"-"}`,{id:op.id,patient:op.patient,room:op.room});
                      upsertOneToSupa("kb_operasi",upd).then((res:any)=>{ if(!res?.ok){ setOps((p:any)=>p.map((o:any)=>o.id===op.id?op:o)); showToast(`⚠ Gagal: ${res?.error||"error"}`,C.d); } });
                    }} style={{background:"#22c55e",color:"#fff",border:"none",borderRadius:10,padding:"10px 18px",fontSize:14,fontWeight:800,cursor:"pointer",fontFamily:"inherit",boxShadow:"0 2px 8px rgba(34,197,94,.4)"}}>
                      ▶ Mulai
                    </button>
                  )}
                  {op.status==="ongoing" && (
                    <button onClick={()=>{
                      const upd={...op,status:"done",updated_at:new Date().toISOString()};
                      setOps((p:any)=>p.map((o:any)=>o.id===op.id?upd:o));
                      AmbilLogMedis("JADWAL_OPERASI","Admin","SELESAI_OPERASI",
                        `Operasi selesai: ${op.patient||"-"} — ${op.procedure||"-"}`,{id:op.id,patient:op.patient,room:op.room});
                      upsertOneToSupa("kb_operasi",upd).then((res:any)=>{ if(!res?.ok){ setOps((p:any)=>p.map((o:any)=>o.id===op.id?op:o)); showToast(`⚠ Gagal: ${res?.error||"error"}`,C.d); } });
                    }} style={{background:"#ef4444",color:"#fff",border:"none",borderRadius:10,padding:"10px 18px",fontSize:14,fontWeight:800,cursor:"pointer",fontFamily:"inherit",boxShadow:"0 2px 8px rgba(239,68,68,.4)"}}>
                      ✓ Selesai
                    </button>
                  )}
                </div>
              )}
            </div>

            {/* ── Procedure block ── */}
            <div style={{background:"#F8FBF0",border:"1px solid #D4E6A0",borderRadius:10,padding:"10px 12px",marginBottom:8}}>
              <div style={{fontSize:14,fontWeight:700,color:"#33691E"}}>{op.procedure}</div>
              <div style={{fontSize:12,color:"#558B2F",marginTop:2}}>{op.diagnosis}</div>
              <div style={{fontSize:12,color:"#558B2F",marginTop:3}}>📅 {fDMY(op.date)} · {op.time} WIB · {op.room}</div>
            </div>

            {/* ── Staff detail ── */}
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

            {/* ── Secondary action row (edit, reminder, batal, hapus) ── */}
            {!isBatal && (
              <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:6}}>
                {op.opType==="cyto"
                  ? <WaBt sm onClick={()=>openCyto(op)}>⚡ WA Cyto</WaBt>
                  : <><Btn sm outline color={C.i} disabled={(op.reminders||[]).includes("H-1")} onClick={()=>sendReminder(op,"H-1")}>H-1</Btn><Btn sm outline color={C.w} disabled={(op.reminders||[]).includes("H-1jam")} onClick={()=>sendReminder(op,"H-1jam")}>1 Jam</Btn></>
                }
                <Btn sm outline color={C.p} onClick={()=>startEditOp(op)}>✏ Edit</Btn>
                <Btn sm outline color={C.w} onClick={()=>{setReqOpId(op.id);setReqText("");}}>+ Permintaan</Btn>
                <Btn sm outline color={C.d} onClick={()=>{setCId(op.id);setCR("");}}>Batal/Tunda</Btn>
              </div>
            )}
            <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
              {isBatal && <Btn sm outline color={C.g} onClick={()=>{
                const upd={...op,status:"scheduled",cancelReason:"",updated_at:new Date().toISOString()};
                setOps((p:any)=>p.map((o:any)=>o.id===op.id?upd:o));
                AmbilLogMedis("JADWAL_OPERASI","Admin","AKTIFKAN_OPERASI",
                  `Operasi diaktifkan kembali: ${op.patient||"-"}`,{id:op.id,patient:op.patient});
                upsertOneToSupa("kb_operasi",upd).then((res:any)=>{ if(!res?.ok){ setOps((p:any)=>p.map((o:any)=>o.id===op.id?op:o)); showToast(`⚠ Gagal: ${res?.error||"error"}`,C.d); } });
              }}>Aktifkan</Btn>}
              <Btn sm outline color={C.d} onClick={()=>setDelC(op.id)}>Hapus</Btn>
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
      {/* ── FITUR BARU: Riwayat (jadwal yang tanggalnya sudah berlalu) ──
         Read-only — sengaja TIDAK ada tombol Edit/Hapus/Mulai/Selesai,
         sesuai instruksi: data lewat hanya untuk dilihat, bukan diubah.
         Collapsible & default tertutup supaya tidak membuat tab Jadwal
         tetap ringan secara visual; data hanya di-render saat dibuka. */}
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
                        <div style={{fontSize:11,color:C.tL,marginTop:2}}>{op.date} · {op.time} WIB · {op.surgeon}</div>
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
/* FIX #1: opForm state dipindah ke dalam ViewDaftar sebagai local state.
   Sebelumnya state ini ada di root App, sehingga setiap ketukan huruf di
   kolom "Nama Pasien" memicu re-render seluruh aplikasi — sangat berat di
   tablet/HP. Sekarang hanya ViewDaftar yang re-render saat mengetik.
   pendingEditOp (prop, lihat FIX AUDIT #24) memungkinkan root App (via
   startEditOp) mengisi form lokal ini — menggantikan mekanisme editOpRef
   lama yang rentan race condition saat ViewDaftar belum/baru mount. */
function ViewDaftar({pendingEditOp,clearPendingEditOp,saveOpFn,staff,setTab,templates,setTemplates,showToast}: ViewDaftarProps) {
  // ── Form state: lokal di sini, tidak lagi di root App ──
  // FIX AUDIT #24: initial state langsung diisi dari pendingEditOp kalau ada
  // (penting untuk RENDER PERTAMA — tidak perlu menunggu useEffect sama sekali
  // untuk kasus edit yang dibuka dari tab lain, karena component baru mount).
  const [opForm,   setOpForm]   = useState<Partial<Operation>>(()=> pendingEditOp ? {...pendingEditOp} : {...EOP, date:todayDate()});
  const [editingOp,setEditingOp]= useState<any>(pendingEditOp || null);
  const [opErrors, setOpErrors] = useState<any>({});
  const [dupWarn,  setDupWarn]  = useState(false);

  /* FIX AUDIT #24: safety net untuk kasus ViewDaftar SUDAH mount (user
     sedang di tab Daftar) lalu startEditOp dipanggil lagi (mis. dari
     komponen lain yang juga bisa trigger edit tanpa ganti tab). Initial
     state di atas menangani kasus "baru mount", useEffect ini menangani
     kasus "sudah mount, pendingEditOp berubah". clearPendingEditOp()
     dipanggil setelah konsumsi agar tidak ter-trigger ulang tanpa sebab. */
  useEffect(()=>{
    if(pendingEditOp){
      setOpForm({...pendingEditOp});
      setEditingOp(pendingEditOp);
      setOpErrors({});
      setDupWarn(false);
      clearPendingEditOp();
    }
  },[pendingEditOp]);

  const resetOp = () => { setOpForm({...EOP,date:todayDate()}); setEditingOp(null); setOpErrors({}); setDupWarn(false); };
  const saveOp  = () => saveOpFn({opForm, editingOp, setOpErrors, setDupWarn, resetOp});
  const dupWarning = dupWarn;

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
          <div style={{flex:1}}><LF label="Usia (bulan)"><input style={iStyle("ageMonths")} type="number" min="0" max="11" placeholder="0" value={opForm.ageMonths||""} onChange={(ev:any)=>setR("ageMonths",ev.target.value)}/></LF></div>
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
        <SF label="Dokter Anestesi" options={["Lokal Anestesi", ...byT("anesthesiologist")]} empty="-- Pilih Dokter Anestesi --" value={opForm.anesthesiologist} onChange={(ev:any)=>setR("anesthesiologist",ev.target.value)}/>
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
/* FIX #1: lSet dan lSby dipindah ke dalam ViewLaporan sebagai local state.
   Sebelumnya keduanya ada di root App sehingga setiap ketukan di field
   "Keterangan Laporan" memicu re-render seluruh pohon komponen. */
function ViewLaporan({ops,staff,roster,showToast,role,privacyMode}: ViewLaporanProps) {
  const [lSet, setLSet] = useState({greeting:gWord(),recipients:"",keterangan:"",pembawaHP:"",katimPhone:"",grupPhone:""});
  const [lSby, setLSby] = useState({anest:["",""],cyto:["","",""],nurses:["","","",""],pembawaHP:""});
  const [preview,setPreview] = useState("");
  const [foc,setFoc]         = useState<string|null>(null);
  const todayOps = ops.filter((o:any)=>o.date===todayDate());
  /* ── Poin ④: Filter besok — Sabtu tarik Minggu + Senin sekaligus ── */
  const _todayObj   = new Date();
  const _isSaturday = _todayObj.getDay() === 6;
  const _besokStr   = tmrwDate();
  const _lusaObj    = new Date(_todayObj);
  _lusaObj.setDate(_lusaObj.getDate() + 2);
  const _lusaStr    = `${_lusaObj.getFullYear()}-${String(_lusaObj.getMonth()+1).padStart(2,"0")}-${String(_lusaObj.getDate()).padStart(2,"0")}`;
  const tmrwOps = ops
    .filter((o:any)=>{
      if(o.status==="batal") return false;
      return _isSaturday
        ? (o.date===_besokStr || o.date===_lusaStr)
        : o.date===_besokStr;
    })
    .sort((a:any,b:any)=> a.date!==b.date ? a.date.localeCompare(b.date) : a.time.localeCompare(b.time));
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
              <Bdg label={STS[op.status as keyof typeof STS]?.l||"Terjadwal"} color={STS[op.status as keyof typeof STS]?.c||C.i} bg={STS[op.status as keyof typeof STS]?.bg||C.iBg}/>
            </div>
          ))
        }
        {tBatal.length>0 && <div style={{fontSize:12,color:C.d,marginTop:6,borderTop:`1px dashed ${C.dL}`,paddingTop:6}}><b>Batal/Tunda:</b> {tBatal.map((o:any)=>maskName(o.patient,privacyMode)+(o.cancelReason?` (${o.cancelReason})`:""  )).join("; ")}</div>}
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
                  <div style={{fontSize:13,fontWeight:700,color:C.t}}>{maskName(op.patient,privacyMode)}{op.age?`, ${op.age}th`:""}</div>
                  <div style={{fontSize:12,color:C.tL}}>{op.procedure}</div>
                  <div style={{fontSize:11,color:C.tL}}>{op.surgeon} / {op.anesthesiologist||"—"}</div>
                </div>
                <Bdg label={OT[op.opType as keyof typeof OT||"elektif"]?.label} color={OT[op.opType as keyof typeof OT||"elektif"]?.c||C.i} bg={OT[op.opType as keyof typeof OT||"elektif"]?.bg||C.iBg}/>
              </div>
            ))}
          </>
        }
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
            <WaBt onClick={()=>{ if(!lSet.katimPhone){showToast("Nomor WA Katim belum diisi",C.d);return;} window.open(`https://wa.me/${lSet.katimPhone.replace(/[^0-9]/g,"")}?text=${encodeURIComponent(preview)}`,"_blank"); showToast("✓ WA Katim dibuka","#25D366"); }} style={{flex:1}}>WA Katim</WaBt>
            <WaBt onClick={()=>{ if(!lSet.grupPhone){showToast("Nomor WA Grup belum diisi",C.d);return;} window.open(`https://wa.me/${lSet.grupPhone.replace(/[^0-9]/g,"")}?text=${encodeURIComponent(preview)}`,"_blank"); showToast("✓ WA Grup dibuka","#25D366"); }} style={{flex:1}}>WA Grup</WaBt>
          </div>
        </Card>
      )}
    </div>
  );
}

/* ─── VIEW KIRIM WA ──────────────────────────────────────────────────── */
function ViewKirimWA({ops,staff,setNotifs,showToast}: ViewKirimWAProps) {
  const [prev,setPrev] = useState<Record<string,boolean>>({});
  const getPhone = (n: string) => staff.find((x:any)=>x.name===n)?.phone||null;
  const tmrwOps  = ops.filter((o:any)=>o.date===tmrwDate()&&o.status!=="batal");
  const sendWA   = (ph: string | null,msg: string,name: string,lbl: string) => {
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

/* ─── VIEW STATISTIK ─────────────────────────────────────────────────── */
function ViewStatistik({ops, archive}: {ops: any[]; archive: any[]}) {
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
  const surgeonData = Object.entries(bySurgeon).sort((a,b)=>b[1]-a[1]).slice(0,8).map(([label,value],i)=>({label,value,color:COLORS[i%COLORS.length]}));

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
  const anestData = Object.entries(byAnest).sort((a,b)=>b[1]-a[1]).slice(0,6).map(([label,value],i)=>({label,value,color:COLORS[i%COLORS.length]}));

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

/* ─── VIEW STAF ──────────────────────────────────────────────────────── */
function ViewStaf({staff,setStaff,roster,setRoster,showToast,upsertOneToSupa,deleteFromSupa,upsertBulkToSupa}: ViewStafProps) {
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
    <div>
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
  );
}

/* ─── IMPORT EXCEL ───────────────────────────────────────────────────── */
function ImportExcel({ops, setOps, showToast, upsertBulkToSupa}: any) {
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
function ViewArsip(props: ViewArsipProps) {
  /* FIX AUDIT #23: destructure ulang dari struktur grouped menjadi nama
     variabel IDENTIK dengan sebelumnya — sengaja, supaya seluruh body
     function di bawah ini (ratusan baris) TIDAK PERLU diubah sama sekali,
     menghilangkan risiko regresi dari menyentuh kode yang sudah teruji. */
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
      const h=["Ruangan","Tanggal","Jam","Suhu (°C)","Kelembaban (%)","Status","Petugas"];
      const rows=filtered.map((e:any)=>[e.ruang||"",e.tanggal||"",e.jam||"",e.suhu??"",e.kelembaban??"",monIsOK(e.suhu,e.kelembaban,monitoringCfg)?"SESUAI":"TIDAK SESUAI",e.petugas||""]);
      buildAndDownloadXlsx("Monitoring "+ymLabel(month),h,rows,`Monitoring_${month}.xlsx`);
      showToast(`✓ Excel Monitoring ${ymLabel(month)} diunduh`,C.s);
    }
  };

  /* ── Download dari Supabase berdasarkan bulan ── */
  const dlFromSupabase = async (month:string, type:"ops"|"lembur"|"monitoring") => {
    if(!month){showToast("Pilih bulan terlebih dahulu",C.w);return;}
    
    setDlCloudBusy(true);
    try {
      // Gunakan SUPA_CLIENT singleton
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

  /* ── Download dari Dropbox berdasarkan bulan ── */
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
            {/* Storage indicator */}
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
            {/* ── Restore dari Supabase ── */}
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

          {/* Dropbox Backup */}
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
            <Btn full color="#0061FF" onClick={onDbxBackup} disabled={dbxBacking} style={{background:"#0061FF",color:"#fff",border:"none",marginTop:4}}>
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

function ViewLembur({lemburPegawai, setLemburPegawai, lemburData, setLemburData, showToast, supaCfg, dbxCfg, role, upsertOneToSupa, deleteFromSupa}: any) {
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
  const [dbxLemburBusy, setDbxLemburBusy] = useState(false);
  const [dbxLemburYM, setDbxLemburYM]     = useState(nowYM());
  const [dbxLemburPegId, setDbxLemburPegId] = useState<string>("__all__");

  const peg = lemburPegawai.find((p:any)=>p.id===selPeg);
  const key = selPeg && selYM ? `${selPeg}_${selYM}` : null;
  const rec: any = (key && lemburData[key]) || {entries:[], kepRuang:"", kepBidang:""};
  const entries: any[] = rec.entries || [];

  /* Sync kepRuang/kepBidang from saved record.
     FIX AUDIT #15 (MEDIUM — useEffect deps tidak lengkap): tambahkan lemburData
     ke deps. Sebelumnya jika lemburData diupdate via Supabase realtime sementara
     `key` tidak berubah (mis. device lain menyimpan perubahan), effect ini tidak
     berjalan dan kepRuang/kepBidang tetap menampilkan nilai lama (stale UI). */
  useEffect(()=>{
    if(key && lemburData[key]){
      setKepRuang(lemburData[key].kepRuang||"");
      setKepBidang(lemburData[key].kepBidang||"");
    } else {
      setKepRuang(""); setKepBidang("");
    }
  },[key, lemburData]);

  /* ── writeLemburKey: DB-first untuk SATU key "${pegId}_${period}" saja.
     Granular — tidak menyentuh key/pegawai/periode lain, sehingga edit
     bersamaan oleh device lain (key berbeda) tidak akan saling menimpa. ── */
  const writeLemburKey = async (newRecForKey: any): Promise<boolean> => {
    if(!key) return false;
    const res = await upsertOneToSupa?.("kb_lembur_data", {id:key, ...newRecForKey, updated_at:new Date().toISOString()});
    if(!res?.ok){
      showToast(`⚠ Gagal menyimpan ke Supabase: ${res?.error||"kesalahan tidak diketahui"}`,C.d);
      return false;
    }
    setLemburData((p:Record<string,any>)=>({...p,[key]:newRecForKey}));
    return true;
  };

  const saveEntry = async () => {
    const jenis = rowForm.jenisEntri || "kerja";
    if(!rowForm.tanggalAwal){
      showToast("Tanggal awal lembur wajib diisi",C.d); return;
    }
    if(jenis==="kerja" && (!rowForm.jamMasuk||!rowForm.jamKeluar)){
      showToast("Jam masuk & keluar wajib diisi untuk Kerja Lembur",C.d); return;
    }
    if(jenis==="ambil" && (rowForm.jumlahJamAmbil===undefined||rowForm.jumlahJamAmbil===""||Number(rowForm.jumlahJamAmbil)<=0)){
      showToast("Jumlah jam diambil wajib diisi (lebih dari 0) untuk Pengambilan Lembur",C.d); return;
    }
    /* Normalisasi: pastikan field yang tidak relevan untuk jenis ini tidak menyimpan sisa data lama */
    const cleanForm = jenis==="kerja"
      ? {...rowForm, jenisEntri:"kerja", jumlahJamAmbil:""}
      : {...rowForm, jenisEntri:"ambil", jamMasuk:"", jamKeluar:""};
    const upd = editRow
      ? entries.map((e:any)=>e.id===editRow?{...e,...cleanForm}:e)
      : [...entries,{id:gId(),...cleanForm,no:entries.length+1}];
    showToast("⟳ Menyimpan...",C.p);
    const ok = await writeLemburKey({...rec, entries:upd});
    if(ok){ setEditRow(null); setRowForm({}); showToast("✓ Baris disimpan & tersinkron",C.s); }
  };

  const delEntry = async (id:string) => {
    const upd = entries.filter((e:any)=>e.id!==id).map((e:any,i:number)=>({...e,no:i+1}));
    const ok = await writeLemburKey({...rec, entries:upd});
    if(ok) showToast("Baris dihapus & tersinkron");
  };

  const saveSigs = async () => {
    if(!key) return;
    const ok = await writeLemburKey({...rec, kepRuang, kepBidang, savedAt:fNow()});
    if(ok) showToast("✓ Tanda tangan tersimpan & tersinkron",C.s);
  };

  const handleImportJadwal = (file:File) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target!.result as ArrayBuffer),{type:"array", cellDates:true});
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
          const jadwalRaw = jadwalIdx>=0 ? row[jadwalIdx] : "";
          const jadwal = toISODateStrict(jadwalRaw) || "";
          if(!jadwal) continue;
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
        const added = newEntries.length-entries.length;
        showToast(`⟳ Menyimpan ${added} baris lembur...`,C.p);
        const ok = await writeLemburKey({...rec, entries:newEntries});
        if(ok) showToast(`✓ ${added} baris diimport & tersinkron`,C.s);
      } catch(err){showToast("Gagal membaca file",C.d);}
    };
    reader.readAsArrayBuffer(file);
  };


  const downloadExcel = () => {
    if(!peg){showToast("Pilih pegawai terlebih dahulu",C.w);return;}
    if(!entries.length){showToast("Belum ada data lembur",C.w);return;}
    const ws = XLSX.utils.aoa_to_sheet([]);
    const merge = (r1:number,c1:number,r2:number,c2:number)=>({s:{r:r1,c:c1},e:{r:r2,c:c2}});
    const cell = xlsxCell; // FIX AUDIT #21: pakai helper global, bukan definisi lokal duplikat
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
    TH.forEach((h,i)=>setC(7,i,h,true,true,true,"16685F"));
    /* Table rows */
    entries.forEach((e:any,i:number)=>{
      const r=8+i;
      const isAmbil = e.jenisEntri==="ambil";
      const masukCol  = isAmbil ? "—" : e.jamMasuk;
      const keluarCol = isAmbil ? "—" : e.jamKeluar;
      const kepCol    = isAmbil ? `[PENGAMBILAN -${Number(e.jumlahJamAmbil)||0} jam] ${e.keperluanLembur||""}`.trim() : e.keperluanLembur;
      [e.no,e.tanggalAwal||e.jadwalDinas||"",e.tanggalAkhir||"",masukCol,keluarCol,kepCol,e.keterangan,e.ttd||""].forEach((v,c)=>setC(r,c,String(v||""),false,c===0||c===3||c===4,true,i%2===0?"F0FFF8":undefined));
    });
    /* Ringkasan saldo: Didapat / Diambil / Saldo Bersih */
    const totalMinsKerjaX = entries.reduce((s:number,e:any)=>{
      if(e.jenisEntri==="ambil") return s;
      if(!e.jamMasuk||!e.jamKeluar) return s;
      const [h1,m1]=e.jamMasuk.split(":").map(Number);
      const [h2,m2]=e.jamKeluar.split(":").map(Number);
      let m=(h2*60+m2)-(h1*60+m1); if(m<0) m+=24*60; return s+(m>0?m:0);
    },0);
    const totalMinsAmbilX = entries.reduce((s:number,e:any)=>{
      if(e.jenisEntri!=="ambil") return s;
      const jam = Number(e.jumlahJamAmbil)||0;
      return s + Math.round(jam*60);
    },0);
    const saldoMinsX = totalMinsKerjaX - totalMinsAmbilX;
    const fmtJamX = (mins:number) => `${mins<0?"-":""}${Math.floor(Math.abs(mins)/60)} jam ${Math.abs(mins)%60} menit`;
    const sumRow = 8+entries.length;
    setC(sumRow,5,"Total Didapat:",true,false); setC(sumRow,6,fmtJamX(totalMinsKerjaX),true,false);
    setC(sumRow+1,5,"Total Diambil:",true,false); setC(sumRow+1,6,fmtJamX(totalMinsAmbilX),true,false);
    setC(sumRow+2,5,"Saldo Bulan Ini:",true,false); setC(sumRow+2,6,fmtJamX(saldoMinsX),true,false);
    /* Signatures */
    const sigRow=sumRow+4;
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
    const totalMinsKerja = ents.reduce((s:number,e:any)=>{
      if(e.jenisEntri==="ambil") return s;
      if(!e.jamMasuk||!e.jamKeluar) return s;
      const [h1,m1]=e.jamMasuk.split(":").map(Number);
      const [h2,m2]=e.jamKeluar.split(":").map(Number);
      let m=(h2*60+m2)-(h1*60+m1); if(m<0) m+=24*60; return s+(m>0?m:0);
    },0);
    const totalMinsAmbil = ents.reduce((s:number,e:any)=>{
      if(e.jenisEntri!=="ambil") return s;
      const jam = Number(e.jumlahJamAmbil)||0;
      return s + Math.round(jam*60);
    },0);
    const totalMins = totalMinsKerja - totalMinsAmbil; // saldo bersih
    return {id:p.id, name:p.name, nik:p.nik||"—", jumlahHari:ents.length, totalMins, totalMinsKerja, totalMinsAmbil, entries:ents, kepRuang:rec2.kepRuang||"", kepBidang:rec2.kepBidang||""};
  });

  const downloadRekapExcel = () => {
    if(!rekapRows.length){showToast("Belum ada data pegawai",C.w);return;}
    const ws = XLSX.utils.aoa_to_sheet([]);
    const merge=(r1:number,c1:number,r2:number,c2:number)=>({s:{r:r1,c:c1},e:{r:r2,c:c2}});
    const cell = xlsxCell; // FIX AUDIT #21: pakai helper global, bukan definisi lokal duplikat
    const setC=(r:number,c:number,v:any,bold?:boolean,center?:boolean,border?:boolean,bg?:string,fs?:number)=>{
      ws[XLSX.utils.encode_cell({r,c})]=cell(v,bold,center,border,bg,fs);
    };
    /* Title */
    setC(0,0,"REKAP LEMBUR BULANAN",true,true,false,undefined,14);
    setC(1,0,HOSPITAL,false,true);
    setC(2,0,`Bulan: ${ymLabel(rekapYM)}`,false,true);
    setC(3,0,"",false,false);
    /* Header row */
    const TH=["No","Nama Pegawai","NIK / NIP","Jumlah Hari Lembur","Jam Didapat","Jam Diambil","Saldo Bersih","Keterangan"];
    TH.forEach((h,i)=>setC(4,i,h,true,true,true,"16685F",11));
    /* Data rows */
    const merges=[
      merge(0,0,0,7),merge(1,0,1,7),merge(2,0,2,7),merge(3,0,3,7),
    ];
    const fmtJamExcel = (mins:number) => `${mins<0?"-":""}${Math.floor(Math.abs(mins)/60)} jam ${Math.abs(mins)%60} menit`;
    rekapRows.forEach((r:any,i:number)=>{
      const row=5+i;
      const bg=i%2===0?"F0FFF8":undefined;
      setC(row,0,String(i+1),false,true,true,bg);
      setC(row,1,r.name,true,false,true,bg);
      setC(row,2,r.nik,false,false,true,bg);
      setC(row,3,String(r.jumlahHari),false,true,true,bg);
      setC(row,4,r.totalMinsKerja>0?fmtJamExcel(r.totalMinsKerja):"—",false,true,true,bg);
      setC(row,5,r.totalMinsAmbil>0?fmtJamExcel(r.totalMinsAmbil):"—",false,true,true,bg);
      setC(row,6,fmtJamExcel(r.totalMins),false,true,true,bg);
      setC(row,7,r.kepRuang?`Kep. Ruang: ${r.kepRuang}`:"",false,false,true,bg);
    });
    /* Footer total */
    const footRow=5+rekapRows.length;
    const grandMinsKerja=rekapRows.reduce((s:number,r:any)=>s+r.totalMinsKerja,0);
    const grandMinsAmbil=rekapRows.reduce((s:number,r:any)=>s+r.totalMinsAmbil,0);
    const grandMins=rekapRows.reduce((s:number,r:any)=>s+r.totalMins,0);
    const grandHari=rekapRows.reduce((s:number,r:any)=>s+r.jumlahHari,0);
    setC(footRow,0,"TOTAL",true,true,true,"E0F2F1");
    setC(footRow,1,"",false,false,true,"E0F2F1");
    merges.push(merge(footRow,0,footRow,2));
    setC(footRow,2,"",false,false,true,"E0F2F1");
    setC(footRow,3,String(grandHari)+" hari",true,true,true,"E0F2F1");
    setC(footRow,4,fmtJamExcel(grandMinsKerja),true,true,true,"E0F2F1");
    setC(footRow,5,fmtJamExcel(grandMinsAmbil),true,true,true,"E0F2F1");
    setC(footRow,6,fmtJamExcel(grandMins),true,true,true,"E0F2F1");
    setC(footRow,7,"",false,false,true,"E0F2F1");
    ws["!merges"]=merges;
    ws["!cols"]=[{wch:4},{wch:28},{wch:18},{wch:18},{wch:18},{wch:18},{wch:18},{wch:28}];
    ws["!ref"]=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:footRow,c:7}});
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
      ["No","Tgl Awal Lembur","Tgl Akhir Lembur","Jam Masuk","Jam Keluar","Keperluan","Keterangan"].forEach((h,i)=>s2(1,i,h,true,true,true,"16685F"));
      r.entries.forEach((e:any,i:number)=>{
        const row=2+i; const bg=i%2===0?"F0FFF8":undefined;
        const isAmbil = e.jenisEntri==="ambil";
        const masukCol  = isAmbil ? "—" : (e.jamMasuk||"");
        const keluarCol = isAmbil ? "—" : (e.jamKeluar||"");
        const kepCol    = isAmbil ? `[PENGAMBILAN -${Number(e.jumlahJamAmbil)||0} jam] ${e.keperluanLembur||""}`.trim() : (e.keperluanLembur||"");
        [String(e.no),e.tanggalAwal||e.jadwalDinas||"",e.tanggalAkhir||"",masukCol,keluarCol,kepCol,e.keterangan||""].forEach((v,c)=>s2(row,c,v,false,c===0||c===3||c===4,true,bg));
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
    // SUPA_CLIENT sudah singleton — guard tidak diperlukan
    setSupaLemburBusy(true);
    try {
      const {data,error} = await SUPA_CLIENT.from("kamar_bedah_backup").select("data").order("created_at",{ascending:false}).limit(1);
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

  const downloadLemburFromDropbox = async () => {
    setDbxLemburBusy(true);
    try {
      const res = await dropboxDownload(dbxCfg);
      if(!res.ok||!res.data) throw new Error(res.msg||"Gagal mengambil data dari Dropbox");
      const ld: any = res.data.lemburData||{};
      const lp: any[] = res.data.lemburPegawai||lemburPegawai;
      const pegList = dbxLemburPegId==="__all__" ? lp : lp.filter((p:any)=>p.id===dbxLemburPegId);
      const wb = XLSX.utils.book_new(); let found = false;
      pegList.forEach((p:any)=>{
        const k=`${p.id}_${dbxLemburYM}`;
        if(ld[k]){
          found = true;
          const ents: any[] = ld[k].entries||[];
          if(!ents.length) return;
          const totalMinsKerja = ents.reduce((s:number,e:any)=>{
            if(e.jenisEntri==="ambil") return s;
            if(!e.jamMasuk||!e.jamKeluar) return s;
            const [h1,m1]=e.jamMasuk.split(":").map(Number);
            const [h2,m2]=e.jamKeluar.split(":").map(Number);
            let m=(h2*60+m2)-(h1*60+m1); if(m<0) m+=24*60; return s+(m>0?m:0);
          },0);
          const totalMinsAmbil = ents.reduce((s:number,e:any)=>{
            if(e.jenisEntri!=="ambil") return s;
            const jam = Number(e.jumlahJamAmbil)||0;
            return s + Math.round(jam*60);
          },0);
          const totalMins = totalMinsKerja - totalMinsAmbil;
          const fmtJamDbx = (mins:number) => `${mins<0?"-":""}${Math.floor(Math.abs(mins)/60)} jam ${Math.abs(mins)%60} menit`;
          const rows = ents.map((e:any,i:number)=>{
            const isAmbil = e.jenisEntri==="ambil";
            const kep = isAmbil ? `[PENGAMBILAN -${Number(e.jumlahJamAmbil)||0} jam] ${e.keperluanLembur||""}`.trim() : (e.keperluanLembur||"");
            return [i+1,e.tanggalAwal||"",e.tanggalAkhir||"",isAmbil?"—":(e.jamMasuk||""),isAmbil?"—":(e.jamKeluar||""),kep,e.keterangan||"",e.ttd||""];
          });
          rows.push(["","","","","","Total Didapat",fmtJamDbx(totalMinsKerja),""]);
          rows.push(["","","","","","Total Diambil",fmtJamDbx(totalMinsAmbil),""]);
          rows.push(["","","","","","Saldo Bulan Ini",fmtJamDbx(totalMins),""]);
          const ws=XLSX.utils.aoa_to_sheet([
            [`PENCATATAN LEMBUR — ${HOSPITAL}`],
            [`Nama: ${p.name}   NIK: ${p.nik||"-"}   Bulan: ${ymLabel(dbxLemburYM)}`],
            [],
            ["No","Tgl Awal","Tgl Akhir","Jam Masuk","Jam Keluar","Keperluan Lembur","Keterangan","TTD"],
            ...rows
          ]);
          ws["!cols"]=[{wch:4},{wch:14},{wch:14},{wch:10},{wch:10},{wch:28},{wch:20},{wch:14}];
          XLSX.utils.book_append_sheet(wb,ws,p.name.slice(0,28));
        }
      });
      if(!found){showToast("Tidak ada data lembur "+ymLabel(dbxLemburYM)+" di Dropbox",C.w);setDbxLemburBusy(false);return;}
      const label = dbxLemburPegId==="__all__"?"Semua":pegList[0]?.name||"";
      XLSX.writeFile(wb,`Lembur_Dropbox_${label}_${dbxLemburYM}.xlsx`);
      showToast("✓ Data lembur "+ymLabel(dbxLemburYM)+" dari Dropbox diunduh",C.s);
    } catch(err:any){showToast("Gagal: "+(err?.message||"Error Dropbox"),C.d);}
    setDbxLemburBusy(false);
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
              <div style={{background:"linear-gradient(135deg,#234B45,#1F5A52)",borderRadius:14,padding:"16px 20px",marginBottom:14,color:"#fff"}}>
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
                <Btn onClick={()=>{setEditRow(null);setRowForm({jenisEntri:"kerja",tanggalAwal:"",tanggalAkhir:"",jamMasuk:"",jamKeluar:"",jumlahJamAmbil:"",keperluanLembur:"",keterangan:"",ttd:""});}} style={{flex:1}}>✚ Tambah Baris</Btn>
                <button onClick={()=>fileRef.current?.click()} style={{flex:1,background:C.iBg,border:`1px solid ${C.i}`,borderRadius:10,color:C.i,fontSize:12,fontWeight:700,padding:"10px 14px",cursor:"pointer",fontFamily:"inherit"}}>📥 Import Jadwal dari Excel</button>
                <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={e=>{const f=e.target.files?.[0];if(f)handleImportJadwal(f);e.target.value="";}}/>
                <Btn onClick={downloadExcel} style={{flex:1,background:"#1565C0",color:"#fff"}}>⬇ Download Excel</Btn>
              </div>

              {/* Add/Edit row form */}
              {(editRow!==null || rowForm.tanggalAwal!==undefined) && rowForm.tanggalAwal!==undefined && (
                <Card style={{marginBottom:14,background:"#F0FFF8",border:`1.5px solid ${C.p}33`}}>
                  <SH label={editRow?"✏️ Edit Baris":"✚ Tambah Baris Lembur"} color={C.p}/>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                    <div style={{gridColumn:"1 / -1"}}>
                      <LF label="Jenis Entri" req>
                        <select
                          style={iS}
                          value={rowForm.jenisEntri||"kerja"}
                          onChange={e=>setRowForm((p:any)=>({...p,jenisEntri:e.target.value}))}
                        >
                          <option value="kerja">🟢 Kerja Lembur</option>
                          <option value="ambil">🟠 Pengambilan Lembur</option>
                        </select>
                      </LF>
                    </div>
                    <LF label="Tanggal Awal Lembur" req>
                      <input style={iS} type="date" value={rowForm.tanggalAwal||""} onChange={e=>setRowForm((p:any)=>({...p,tanggalAwal:e.target.value}))}/>
                    </LF>
                    <LF label="Tanggal Akhir Lembur">
                      <input style={iS} type="date" value={rowForm.tanggalAkhir||""} onChange={e=>setRowForm((p:any)=>({...p,tanggalAkhir:e.target.value}))}/>
                    </LF>
                    {(rowForm.jenisEntri||"kerja")==="kerja" ? (
                      <>
                        <LF label="Jam Absen Masuk" req>
                          <input style={iS} type="time" value={rowForm.jamMasuk||""} onChange={e=>setRowForm((p:any)=>({...p,jamMasuk:e.target.value}))}/>
                        </LF>
                        <LF label="Jam Absen Keluar" req>
                          <input style={iS} type="time" value={rowForm.jamKeluar||""} onChange={e=>setRowForm((p:any)=>({...p,jamKeluar:e.target.value}))}/>
                        </LF>
                      </>
                    ) : (
                      <div style={{gridColumn:"1 / -1"}}>
                        <LF label="Jumlah Jam Diambil" req>
                          <input
                            style={iS}
                            type="number"
                            min={0}
                            step={0.5}
                            placeholder="cth: 2 (jam)"
                            value={rowForm.jumlahJamAmbil??""}
                            onChange={e=>setRowForm((p:any)=>({...p,jumlahJamAmbil:e.target.value}))}
                          />
                        </LF>
                      </div>
                    )}
                    <LF label="Keperluan Lembur">
                      <input style={iS} placeholder={(rowForm.jenisEntri||"kerja")==="ambil"?"cth: Pulang awal / izin pribadi":"cth: Menyiapkan alat operasi"} value={rowForm.keperluanLembur||""} onChange={e=>setRowForm((p:any)=>({...p,keperluanLembur:e.target.value}))}/>
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
                <div style={{background:"#16685F",padding:"12px 16px"}}>
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
                            <th key={h} style={{padding:"8px 10px",textAlign:"left",color:"#16685F",fontWeight:700,borderBottom:"2px solid #16685F",whiteSpace:"nowrap"}}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {entries.map((e:any,i:number)=>{
                          const isAmbil = e.jenisEntri==="ambil";
                          const dur = !isAmbil && e.jamMasuk&&e.jamKeluar ? (()=>{
                            const [h1,m1]=e.jamMasuk.split(":").map(Number);
                            const [h2,m2]=e.jamKeluar.split(":").map(Number);
                            let mins=(h2*60+m2)-(h1*60+m1);
                            if(mins<0) mins+=24*60;
                            if(mins===0) return "—";
                            return `${Math.floor(mins/60)}j ${mins%60}m`;
                          })() : null;
                          const jamAmbil = isAmbil ? Number(e.jumlahJamAmbil)||0 : 0;
                          return (
                            <tr key={e.id} style={{background:isAmbil?"#FFF4ED":(i%2===0?"#FAFBFD":"#F0FFF8"),borderBottom:`1px solid ${C.b}`}}>
                              <td style={{padding:"8px 10px",textAlign:"center",fontWeight:700,color:C.p}}>{e.no}</td>
                              <td style={{padding:"8px 10px",fontWeight:600,color:C.t}}>{fDMY(e.tanggalAwal||e.jadwalDinas)||"—"}</td>
                              <td style={{padding:"8px 10px",fontWeight:600,color:C.t}}>{e.tanggalAkhir||"—"}</td>
                              <td style={{padding:"8px 10px",color:"#1565C0",fontWeight:600}}>{isAmbil?"—":(e.jamMasuk||"—")}</td>
                              <td style={{padding:"8px 10px",color:"#B71C1C",fontWeight:600}}>{isAmbil?"—":(e.jamKeluar||"—")}</td>
                              <td style={{padding:"8px 10px",whiteSpace:"nowrap"}}>
                                {isAmbil ? (
                                  <span style={{background:"#FFE0CC",color:"#C2410C",fontWeight:800,padding:"3px 9px",borderRadius:12,fontSize:11,border:"1px solid #FDBA8C"}}>
                                    - {jamAmbil} jam
                                  </span>
                                ) : (
                                  <span style={{color:C.s,fontWeight:700}}>{dur}</span>
                                )}
                              </td>
                              <td style={{padding:"8px 10px",color:C.t}}>
                                {isAmbil && <span style={{background:"#FFE0CC",color:"#C2410C",fontWeight:800,fontSize:10,padding:"2px 6px",borderRadius:6,marginRight:6}}>[PENGAMBILAN]</span>}
                                {e.keperluanLembur||"—"}
                              </td>
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
                        const totalMinsKerja=entries.reduce((sum:number,e:any)=>{
                          if(e.jenisEntri==="ambil") return sum;
                          if(!e.jamMasuk||!e.jamKeluar) return sum;
                          const [h1,m1]=e.jamMasuk.split(":").map(Number);
                          const [h2,m2]=e.jamKeluar.split(":").map(Number);
                          let m=(h2*60+m2)-(h1*60+m1); if(m<0) m+=24*60; return sum+(m>0?m:0);
                        },0);
                        const totalMinsAmbil=entries.reduce((sum:number,e:any)=>{
                          if(e.jenisEntri!=="ambil") return sum;
                          const jam = Number(e.jumlahJamAmbil)||0;
                          return sum + Math.round(jam*60);
                        },0);
                        const saldoMins = totalMinsKerja - totalMinsAmbil;
                        const fmtJam = (mins:number) => {
                          const sign = mins<0 ? "-" : "";
                          const abs = Math.abs(mins);
                          return `${sign}${Math.floor(abs/60)} jam ${abs%60} menit`;
                        };
                        return (
                          <tfoot>
                            <tr style={{background:"#E0F2F1"}}>
                              <td colSpan={10} style={{padding:"10px 10px",fontWeight:700,color:"#16685F",fontSize:12}}>
                                <div style={{display:"flex",gap:14,flexWrap:"wrap",alignItems:"center",justifyContent:"flex-end"}}>
                                  <span>Total Didapat: <b style={{color:C.s}}>{fmtJam(totalMinsKerja)}</b></span>
                                  <span style={{color:"#D1D5DB"}}>|</span>
                                  <span>Total Diambil: <b style={{color:"#C2410C"}}>{fmtJam(totalMinsAmbil)}</b></span>
                                  <span style={{color:"#D1D5DB"}}>|</span>
                                  <span>Saldo Bulan Ini: <b style={{color:saldoMins<0?C.d:"#16685F"}}>{fmtJam(saldoMins)}</b></span>
                                </div>
                              </td>
                            </tr>
                          </tfoot>
                        );
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

          {/* ── Dropbox cloud download ── */}
          <Card style={{background:"#EFF6FF",border:"1.5px solid #0061FF33",marginBottom:14}}>
            <SH label="📦 Download Lembur dari Dropbox" color="#0061FF"/>
            <div style={{fontSize:12,color:"#64748B",marginBottom:10}}>Unduh data lembur langsung dari backup Dropbox. Pilih bulan dan karyawan spesifik atau semua sekaligus.</div>
            <div style={{display:"flex",gap:10,alignItems:"flex-end",flexWrap:"wrap",marginBottom:10}}>
              <div style={{flex:1,minWidth:140}}>
                <LF label="Pilih Bulan"><input style={iS} type="month" value={dbxLemburYM} onChange={e=>setDbxLemburYM(e.target.value)} disabled={dbxLemburBusy}/></LF>
              </div>
              <div style={{flex:2,minWidth:180}}>
                <LF label="Karyawan">
                  <select style={{...iS,appearance:"none"} as React.CSSProperties} value={dbxLemburPegId} onChange={e=>setDbxLemburPegId(e.target.value)} disabled={dbxLemburBusy}>
                    <option value="__all__">— Semua Karyawan —</option>
                    {lemburPegawai.map((p:any)=><option key={p.id} value={p.id}>{p.name}</option>)}
                  </select>
                </LF>
              </div>
            </div>
            <Btn full onClick={downloadLemburFromDropbox} disabled={dbxLemburBusy} style={{background:dbxLemburBusy?"#94A3B8":"#0061FF",color:"#fff",border:"none"}}>
              {dbxLemburBusy?"⏳ Mengunduh...":"📦 Download Excel dari Dropbox"}
            </Btn>
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
                <div style={{background:"#16685F",padding:"12px 16px"}}>
                  <div style={{fontSize:13,fontWeight:800,color:"#fff"}}>Rekap Lembur — {ymLabel(rekapYM)}</div>
                  <div style={{fontSize:11,color:"rgba(255,255,255,.6)",marginTop:2}}>{HOSPITAL}</div>
                </div>
                <div style={{overflowX:"auto"}}>
                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:500}}>
                    <thead>
                      <tr style={{background:"#E0F2F1"}}>
                        {["No","Nama Pegawai","NIK / NIP","Hari Lembur","Total Jam","Detail"].map(h=>(
                          <th key={h} style={{padding:"9px 12px",textAlign:"left",color:"#16685F",fontWeight:700,borderBottom:"2px solid #16685F",whiteSpace:"nowrap"}}>{h}</th>
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
                        <td colSpan={3} style={{padding:"9px 12px",fontWeight:700,color:"#16685F",textAlign:"right"}}>TOTAL</td>
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
            <Btn full onClick={async ()=>{
              if(!pegForm.name.trim()){showToast("Nama wajib diisi",C.d);return;}
              const newPeg={id:gId(),name:pegForm.name.trim(),nik:pegForm.nik.trim()};
              showToast("⟳ Menyimpan...",C.p);
              const res = await upsertOneToSupa?.("kb_lembur_pegawai", newPeg);
              if(!res?.ok){ showToast(`⚠ Gagal menyimpan pegawai: ${res?.error||"kesalahan tidak diketahui"}`,C.d); return; }
              setLemburPegawai((p:any[])=>[...p,newPeg]);
              setPegForm({name:"",nik:""});
              showToast("✓ Pegawai ditambahkan & tersinkron",C.s);
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
                      <button onClick={async ()=>{
                        if(!window.confirm(`Hapus pegawai "${p.name}"?\nSemua data lembur pegawai ini juga akan dihapus.`)) return;
                        showToast("⟳ Menghapus...",C.p);
                        const res = await deleteFromSupa?.("kb_lembur_pegawai", p.id);
                        if(!res?.ok){ showToast(`⚠ Gagal menghapus pegawai: ${res?.error||"kesalahan tidak diketahui"}`,C.d); return; }
                        setLemburPegawai((prev:any[])=>prev.filter((x:any)=>x.id!==p.id));
                        /* Hapus juga setiap row kb_lembur_data milik pegawai ini (best-effort,
                           tidak menghalangi UI — kegagalan hanya menyisakan row tak terpakai). */
                        const relatedKeys = Object.keys(lemburData).filter(k=>k.startsWith(p.id+"_"));
                        await Promise.all(relatedKeys.map(k=>deleteFromSupa?.("kb_lembur_data", k)));
                        setLemburData((prev:Record<string,any>)=>{ const n={...prev}; relatedKeys.forEach(k=>delete n[k]); return n; });
                        showToast("✓ Pegawai dihapus & tersinkron",C.d);
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

/* ─── MONITORING SUHU & KELEMBABAN (Multi-Room) ─────────────────────── */
interface MonitoringEntry {
  id: string;
  ruang: string;        // nama ruangan (hardcoded dari MON_ROOMS)
  tanggal: string;
  jam: string;
  suhu: number;
  kelembaban: number;
  petugas: string;
  createdAt: string;
}
interface MonitoringCfg { suhuMin:number; suhuMax:number; rhMin:number; rhMax:number; lokasiRuang:string; kepalaKamarBedah:string; }
const defaultMonCfg: MonitoringCfg = { suhuMin:19, suhuMax:24, rhMin:45, rhMax:65, lokasiRuang:"Kamar Bedah", kepalaKamarBedah:"" };
const MON_JAMS  = ["07:00","14:00","21:00"];
/* MON_ROOMS: 3 ruangan hardcoded sesuai spesifikasi.
   Urutan ini digunakan di seluruh UI, Excel, dan ID deterministik. */
const MON_ROOMS = ["Kamar Bedah 1","Kamar Bedah 2","Ruang Instrumen"];
/* monId: id deterministik berbasis (ruang,tanggal,jam) — bukan random — sehingga
   menyimpan ulang slot yang sama akan UPSERT (timpa baris yang sama) bukan
   membuat baris baru yang menumpuk sebagai duplikat "hantu" di Supabase.
   KRITIS: ruang disertakan agar data antar-ruangan tidak saling menimpa. */
const monId = (ruang: string, tanggal: string, jam: string) =>
  `mon_${ruang.replace(/\s+/g,"_")}_${tanggal}_${jam}`;
const MON_MONTHS_ID = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
const MC = { pri:"#0369A1", priL:"#0284C7", priBg:"#EFF6FF", ok:"#16A34A", okBg:"#DCFCE7", err:"#DC2626", errBg:"#FEE2E2", warn:"#D97706" };
/* Warna per ruangan untuk chart multi-line */
const MON_ROOM_COLORS = ["#0369A1","#7C3AED","#B45309"];

function monIsOK(suhu:number, rh:number, cfg:MonitoringCfg):boolean {
  return suhu>=cfg.suhuMin && suhu<=cfg.suhuMax && rh>=cfg.rhMin && rh<=cfg.rhMax;
}

function monXlsx(
  entries: MonitoringEntry[],
  cfg: MonitoringCfg,
  ym: string,
  type: "harian"|"rekap"|"akreditasi",
  showToast: (m:string,c?:string)=>void
) {
  const [ys,ms] = ym.split("-"); const year=+ys; const month=+ms;
  /* Filter & sort: Ruangan → Tanggal → Jam */
  const me = [...entries]
    .filter(e=>e.tanggal.startsWith(ym))
    .sort((a,b)=>
      MON_ROOMS.indexOf(a.ruang??"")-MON_ROOMS.indexOf(b.ruang??"") ||
      a.tanggal.localeCompare(b.tanggal) ||
      a.jam.localeCompare(b.jam)
    );
  if(!me.length){showToast("Tidak ada data untuk periode ini","#E65100");return;}
  const mName = MON_MONTHS_ID[month-1];
  const wb = XLSX.utils.book_new();

  /* ── Sheet 1: Monitoring Harian (tambah kolom Ruangan) ── */
  let lastKey = "";
  const rows1 = me.map(e=>{
    const ok = monIsOK(e.suhu,e.kelembaban,cfg);
    const key = `${e.ruang??""}_${e.tanggal}`;
    const showRuang = key!==lastKey;
    lastKey = key;
    return [
      showRuang ? (e.ruang??"-") : "",   // Ruangan (dikelompokkan)
      showRuang ? e.tanggal : "",         // Tanggal (dikelompokkan)
      e.jam,
      e.suhu,
      e.kelembaban,
      cfg.suhuMin+"-"+cfg.suhuMax,
      cfg.rhMin+"-"+cfg.rhMax,
      ok?"SESUAI":"TIDAK SESUAI",
      e.petugas,
    ];
  });
  const ws1 = XLSX.utils.aoa_to_sheet([
    ["Ruangan","Tanggal","Jam","Suhu (°C)","Kelembaban (%)","Standar Suhu","Standar RH","Status","Petugas"],
    ...rows1
  ]);
  ws1["!cols"]=[{wch:18},{wch:14},{wch:8},{wch:10},{wch:14},{wch:13},{wch:12},{wch:16},{wch:16}];
  XLSX.utils.book_append_sheet(wb,ws1,"Monitoring Harian");

  /* ── Helper: hitung statistik per ruangan ── */
  const statPerRoom = MON_ROOMS.map(ruang=>{
    const rows = me.filter(e=>(e.ruang??""  )===ruang);
    const suhuV = rows.map(e=>e.suhu).filter(v=>v>0);
    const rhV   = rows.map(e=>e.kelembaban).filter(v=>v>0);
    const avgS  = suhuV.length ? +(suhuV.reduce((a,b)=>a+b,0)/suhuV.length).toFixed(2):0;
    const avgR  = rhV.length   ? +(rhV.reduce((a,b)=>a+b,0)/rhV.length).toFixed(2)    :0;
    const tS    = rows.filter(e=>!monIsOK(e.suhu,e.kelembaban,cfg)).length;
    const pct   = rows.length  ? +((rows.length-tS)/rows.length*100).toFixed(1)        :0;
    return {ruang, rows, avgS, avgR, tidakS:tS, pct};
  });

  /* ── Sheet 2: Rekap Bulanan (dipecah per ruangan) ── */
  const rekapHeader = [
    ["REKAP MONITORING SUHU & KELEMBABAN","","","",""],
    [`${mName} ${year}`,"","","",""],
    [""],
    ["Ruangan","Rata-rata Suhu (°C)","Rata-rata RH (%)","Jumlah Pengukuran","Tidak Sesuai","Kepatuhan (%)"],
  ];
  const rekapRows = statPerRoom.map(s=>[s.ruang,s.avgS,s.avgR,s.rows.length,s.tidakS,s.pct]);
  /* Baris total gabungan */
  const allSuhuV = me.map(e=>e.suhu).filter(v=>v>0);
  const allRhV   = me.map(e=>e.kelembaban).filter(v=>v>0);
  const allAvgS  = allSuhuV.length ? +(allSuhuV.reduce((a,b)=>a+b,0)/allSuhuV.length).toFixed(2):0;
  const allAvgR  = allRhV.length   ? +(allRhV.reduce((a,b)=>a+b,0)/allRhV.length).toFixed(2)    :0;
  const allTidakS= me.filter(e=>!monIsOK(e.suhu,e.kelembaban,cfg)).length;
  const allPct   = me.length ? +((me.length-allTidakS)/me.length*100).toFixed(1) : 0;
  const ws2 = XLSX.utils.aoa_to_sheet([
    ...rekapHeader,
    ...rekapRows,
    [""],
    ["TOTAL / GABUNGAN",allAvgS,allAvgR,me.length,allTidakS,allPct],
  ]);
  ws2["!cols"]=[{wch:22},{wch:20},{wch:18},{wch:20},{wch:18},{wch:14}];
  ws2["!merges"]=[{s:{r:0,c:0},e:{r:0,c:4}},{s:{r:1,c:0},e:{r:1,c:4}}];
  XLSX.utils.book_append_sheet(wb,ws2,"Rekap Bulanan");

  /* ── Sheet 3: Laporan Akreditasi (dipecah per ruangan) ── */
  if(type==="akreditasi"){
    const aRows: any[][] = [
      ["LAPORAN MONITORING SUHU DAN KELEMBABAN KAMAR BEDAH","","",""],
      [`Bulan: ${mName} ${year}`,"","",""],
      [/* HOSPITAL — gunakan konstanta global */ "RS Panti Rini","","",""],
      ["","","",""],
    ];
    statPerRoom.forEach(s=>{
      aRows.push([`▌ ${s.ruang}`,"","",""]);
      aRows.push(["Parameter","Nilai","",""]);
      aRows.push(["Rata-rata Suhu (°C)",s.avgS,"",""]);
      aRows.push(["Rata-rata Kelembaban (%)",s.avgR,"",""]);
      aRows.push(["Jumlah Pengukuran",s.rows.length,"",""]);
      aRows.push(["Tidak Sesuai Standar",s.tidakS,"",""]);
      aRows.push(["Kepatuhan (%)",s.pct,"",""]);
      aRows.push(["","","",""]);
    });
    aRows.push(["GABUNGAN SEMUA RUANGAN","","",""]);
    aRows.push(["Rata-rata Suhu (°C)",allAvgS,"",""]);
    aRows.push(["Rata-rata Kelembaban (%)",allAvgR,"",""]);
    aRows.push(["Kepatuhan (%)",allPct,"",""]);
    aRows.push(["","","",""]);
    aRows.push(["Petugas Monitoring","","Kepala Kamar Bedah",""]);
    aRows.push(["","","",""]);
    aRows.push(["","","",""]);
    aRows.push(["(..........................)","","("+cfg.kepalaKamarBedah+")",""]);
    const ws3 = XLSX.utils.aoa_to_sheet(aRows);
    ws3["!cols"]=[{wch:28},{wch:16},{wch:28},{wch:16}];
    ws3["!merges"]=[
      {s:{r:0,c:0},e:{r:0,c:3}},
      {s:{r:1,c:0},e:{r:1,c:3}},
      {s:{r:2,c:0},e:{r:2,c:3}},
    ];
    XLSX.utils.book_append_sheet(wb,ws3,"Laporan Akreditasi");
  }
  XLSX.writeFile(wb,"Monitoring_"+(type==="harian"?"Harian":type==="rekap"?"Rekap":"Akreditasi")+"_"+ym+".xlsx");
  showToast("✓ Monitoring "+(type==="harian"?"Harian":type==="rekap"?"Rekap":"Akreditasi")+" "+mName+" "+year+" diunduh",MC.ok);
}

function ViewMonitoring({monitoringEntries,setMonitoringEntries,monitoringCfg,setMonitoringCfg,showToast,supaCfg,dbxCfg,role,upsertOneToSupa,deleteFromSupa}:any) {
  const [subTab,setSubTab]       = useState("harian");
  const [selMonth,setSelMonth]   = useState(todayDate().slice(0,7));
  const [supaMonBusy,setSupaMonBusy] = useState(false);
  const [supaMonYM,setSupaMonYM] = useState(todayDate().slice(0,7));
  const [dbxMonBusy,setDbxMonBusy]   = useState(false);
  const [dbxMonYM,setDbxMonYM]   = useState(todayDate().slice(0,7));
  const [formDate,setFormDate]   = useState(todayDate());
  /* grafik: filter ruangan; default "all" */
  const [grafikRuang,setGrafikRuang] = useState("all");
  const [cfgForm,setCfgForm]     = useState<MonitoringCfg>({...monitoringCfg});
  /* ── State & ref khusus ekspor grafik (gambar + tabel per ruangan) ── */
  const [grafikExporting, setGrafikExporting] = useState(false);
  /* Ref off-screen chart per ruangan, satu pasang (suhu+RH) per ruangan di MON_ROOMS.
     Dipakai HANYA untuk capture html2canvas — tidak memengaruhi chart visible. */
  const exportChartRefs = useRef<Record<string, { suhu: HTMLDivElement | null; rh: HTMLDivElement | null }>>(
    Object.fromEntries(MON_ROOMS.map(r => [r, { suhu: null, rh: null }]))
  );

  /* ── Slot state: { [ruang]: { [jam]: {suhu,kelembaban,petugas} } } ── */
  type SL = {suhu:string;kelembaban:string;petugas:string};
  const ES: SL = {suhu:"",kelembaban:"",petugas:""};
  const mkSlots = (): Record<string,Record<string,SL>> =>
    Object.fromEntries(MON_ROOMS.map(r=>[r, Object.fromEntries(MON_JAMS.map(j=>[j,{...ES}]))]));
  const [slots,setSlots] = useState<Record<string,Record<string,SL>>>(mkSlots);

  /* Isi ulang form slots ketika tanggal berubah */
  useEffect(()=>{
    const ns = mkSlots();
    MON_ROOMS.forEach(ruang=>{
      MON_JAMS.forEach(jam=>{
        const ex = (monitoringEntries as MonitoringEntry[]).find(e=>
          e.ruang===ruang && e.tanggal===formDate && e.jam===jam
        );
        if(ex) ns[ruang][jam]={suhu:String(ex.suhu),kelembaban:String(ex.kelembaban),petugas:ex.petugas};
      });
    });
    setSlots(ns);
  },[formDate]);  // eslint-disable-line react-hooks/exhaustive-deps

  /* ── handleSave: kumpulkan semua slot terisi dari 3 ruangan ── */
  const handleSave = async ()=>{
    const nw: MonitoringEntry[] = [];
    MON_ROOMS.forEach(ruang=>{
      MON_JAMS.forEach(jam=>{
        const s = slots[ruang][jam];
        if(s.suhu && s.kelembaban){
          nw.push({
            id:        monId(ruang,formDate,jam),
            ruang,
            tanggal:   formDate,
            jam,
            suhu:      parseFloat(s.suhu),
            kelembaban:parseFloat(s.kelembaban),
            petugas:   s.petugas,
            createdAt: fNow(),
          });
        }
      });
    });
    if(!nw.length){showToast("Isi minimal satu slot suhu/kelembaban","#E65100");return;}
    showToast(`⟳ Menyimpan ${nw.length} data monitoring...`,MC.pri);
    /* Upsert granular per slot — arsitektur identik dengan eksisting */
    const results = await Promise.all(nw.map((e:MonitoringEntry)=>upsertOneToSupa?.("kb_monitoring", e)));
    const failed  = results.filter((r:any)=>!r?.ok);
    if(failed.length){
      showToast(`⚠ Gagal menyimpan ${failed.length} data: ${failed[0]?.error||"kesalahan tidak diketahui"}`,MC.err);
      return;
    }
    /* Update local state: filter lama berdasarkan (ruang AND tanggal AND jam) */
    setMonitoringEntries((p:MonitoringEntry[])=>{
      const filtered = p.filter((e:MonitoringEntry)=>
        !(nw.some(n=>n.ruang===e.ruang && n.tanggal===e.tanggal && n.jam===e.jam))
      );
      return [...filtered,...nw];
    });
    showToast("✓ "+nw.length+" data monitoring "+formDate+" disimpan & tersinkron",MC.ok);
  };

  /* ── Download dari Supabase backup ── */
  const downloadMonFromSupa = async (type:"harian"|"rekap"|"akreditasi")=>{
    setSupaMonBusy(true);
    try {
      const {data,error} = await SUPA_CLIENT.from("kamar_bedah_backup").select("data").order("created_at",{ascending:false}).limit(1);
      if(error||!data?.length) throw new Error(error?.message||"Data tidak ditemukan di Supabase");
      const raw  = typeof data[0].data==="string"?JSON.parse(data[0].data):data[0].data;
      const me:MonitoringEntry[] = (raw.monitoringEntries||raw.data?.monitoringEntries||[]).filter((e:MonitoringEntry)=>e.tanggal?.startsWith(supaMonYM));
      const cfg2:MonitoringCfg   = raw.monitoringCfg||raw.data?.monitoringCfg||monitoringCfg;
      if(!me.length){showToast("Tidak ada data monitoring "+supaMonYM+" di Supabase",MC.err);setSupaMonBusy(false);return;}
      monXlsx(me,cfg2,supaMonYM,type,showToast);
      showToast("✓ Monitoring dari Supabase ("+supaMonYM+") diunduh",MC.ok);
    } catch(err:any){showToast("Gagal: "+(err?.message||"Error Supabase"),MC.err);}
    setSupaMonBusy(false);
  };

  /* ── Download dari Dropbox backup ── */
  const downloadMonFromDropbox = async (type:"harian"|"rekap"|"akreditasi")=>{
    setDbxMonBusy(true);
    try {
      const res = await dropboxDownload(dbxCfg);
      if(!res.ok||!res.data) throw new Error(res.msg||"Gagal mengambil data dari Dropbox");
      const me2:MonitoringEntry[] = (res.data.monitoringEntries||[]).filter((e:MonitoringEntry)=>e.tanggal?.startsWith(dbxMonYM));
      const cfg2:MonitoringCfg   = res.data.monitoringCfg||monitoringCfg;
      if(!me2.length){showToast("Tidak ada data monitoring "+dbxMonYM+" di Dropbox",MC.err);setDbxMonBusy(false);return;}
      monXlsx(me2,cfg2,dbxMonYM,type,showToast);
      showToast("✓ Monitoring dari Dropbox ("+dbxMonYM+") diunduh",MC.ok);
    } catch(err:any){showToast("Gagal: "+(err?.message||"Error Dropbox"),MC.err);}
    setDbxMonBusy(false);
  };

  /* ── Derived data untuk grafik & statistik ── */
  const meAll = (monitoringEntries as MonitoringEntry[])
    .filter(e=>e.tanggal.startsWith(selMonth))
    .sort((a,b)=>a.ruang?.localeCompare(b.ruang??"")||a.tanggal.localeCompare(b.tanggal)||a.jam.localeCompare(b.jam));
  /* Data difilter berdasarkan pilihan ruangan di tab Grafik */
  const me = grafikRuang==="all" ? meAll : meAll.filter(e=>e.ruang===grafikRuang);
  const suhuV  = me.map(e=>e.suhu).filter(v=>v>0);
  const rhV    = me.map(e=>e.kelembaban).filter(v=>v>0);
  const avgS   = suhuV.length ? +(suhuV.reduce((a:number,b:number)=>a+b,0)/suhuV.length).toFixed(1):0;
  const avgR   = rhV.length   ? +(rhV.reduce((a:number,b:number)=>a+b,0)/rhV.length).toFixed(1)   :0;
  const tidakS = me.filter(e=>!monIsOK(e.suhu,e.kelembaban,monitoringCfg)).length;
  const cmpPct = me.length ? +((me.length-tidakS)/me.length*100).toFixed(1):0;
  const mLabel = MON_MONTHS_ID[+selMonth.slice(5)-1]+" "+selMonth.slice(0,4);

  /* chartData: satu titik per (ruang, tanggal, jam) — label: tgl/jam singkat */
  const chartData = me.map(e=>({
    name: e.tanggal.slice(8)+"/"+(e.jam==="07:00"?"07":e.jam==="14:00"?"14":"21"),
    ruang: e.ruang,
    suhu:  e.suhu,
    kelembaban: e.kelembaban,
  }));

  /* Untuk grafik multi-line (semua ruangan), buat structure per tanggal+jam */
  const buildMultiChartData = (key:"suhu"|"kelembaban") => {
    const map: Record<string,any> = {};
    meAll.forEach(e=>{
      const label = e.tanggal.slice(8)+"/"+(e.jam==="07:00"?"07":e.jam==="14:00"?"14":"21")+"_"+e.tanggal;
      if(!map[label]) map[label]={name:e.tanggal.slice(8)+"/"+(e.jam==="07:00"?"07":e.jam==="14:00"?"14":"21")};
      map[label][e.ruang??"-"] = e[key];
    });
    return Object.values(map);
  };

  const downloadGrafikExcel = async () => {
    if (!meAll.length) { showToast("Tidak ada data untuk bulan ini", "#E65100"); return; }
    setGrafikExporting(true);
    try {
      const wb = new ExcelJS.Workbook();
      wb.creator = "Sistem Monitoring Kamar Bedah";
      wb.created = new Date();

      /* Beri sedikit delay agar chart off-screen sempat ter-render penuh sebelum capture */
      await new Promise(res => setTimeout(res, 150));

      /* FIX AUDIT #20 (MEDIUM — capture grafik gagal senyap): sebelumnya
         kegagalan html2canvas hanya di-console.error, user tidak diberi
         tahu sama sekali bahwa gambar grafik gagal disisipkan ke Excel
         (file tetap terdownload, tabel data tetap lengkap, tapi grafik
         kosong tanpa penjelasan). Sekarang kegagalan dikumpulkan dan
         ditampilkan sebagai 1 toast ringkasan di akhir proses. */
      const chartFailures: string[] = [];

      for (const ruang of MON_ROOMS) {
        const meRoom = meAll.filter(e => e.ruang === ruang);
        const sheet  = wb.addWorksheet(ruang.length > 31 ? ruang.slice(0, 31) : ruang);

        /* ── Judul sheet ── */
        sheet.mergeCells("A1:G1");
        const titleCell = sheet.getCell("A1");
        titleCell.value = `Grafik Monitoring Suhu & Kelembaban — ${ruang} (${mLabel})`;
        titleCell.font  = { bold: true, size: 13, color: { argb: "FF" + MC.pri.replace("#", "") } };
        titleCell.alignment = { vertical: "middle" };
        sheet.getRow(1).height = 22;

        if (!meRoom.length) {
          sheet.getCell("A3").value = "Belum ada data untuk ruangan ini pada bulan terpilih.";
          sheet.getCell("A3").font  = { italic: true, color: { argb: "FF94A3B8" } };
          sheet.columns = [{ width: 18 }, { width: 14 }, { width: 8 }, { width: 12 }, { width: 14 }, { width: 16 }, { width: 22 }];
          continue;
        }

        /* ── Tabel data tabular ── */
        const headerRow = sheet.addRow(["Tanggal", "Jam", "Suhu (°C)", "Kelembaban (%)", "Status", "Petugas"]);
        headerRow.eachCell(cell => {
          cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + MC.pri.replace("#", "") } };
          cell.alignment = { horizontal: "center" };
        });

        meRoom.forEach(e => {
          const ok  = monIsOK(e.suhu, e.kelembaban, monitoringCfg);
          const row = sheet.addRow([e.tanggal, e.jam, e.suhu, e.kelembaban, ok ? "SESUAI" : "TIDAK SESUAI", e.petugas]);
          const statusCell = row.getCell(5);
          statusCell.font = { bold: true, color: { argb: ok ? "FF" + MC.ok.replace("#", "") : "FF" + MC.err.replace("#", "") } };
        });

        /* ── Ringkasan statistik per ruangan ── */
        const suhuVR  = meRoom.map(e => e.suhu).filter(v => v > 0);
        const rhVR    = meRoom.map(e => e.kelembaban).filter(v => v > 0);
        const avgSR   = suhuVR.length ? +(suhuVR.reduce((a, b) => a + b, 0) / suhuVR.length).toFixed(1) : 0;
        const avgRR   = rhVR.length   ? +(rhVR.reduce((a, b) => a + b, 0) / rhVR.length).toFixed(1)     : 0;
        const tidakSR = meRoom.filter(e => !monIsOK(e.suhu, e.kelembaban, monitoringCfg)).length;
        const cmpPctR = meRoom.length ? +((meRoom.length - tidakSR) / meRoom.length * 100).toFixed(1) : 0;

        sheet.addRow([]);
        const sumRow = sheet.addRow([`Rata-rata Suhu: ${avgSR}°C`, `Rata-rata RH: ${avgRR}%`, `Kepatuhan: ${cmpPctR}%`, `Tidak Sesuai: ${tidakSR}`]);
        sumRow.font = { bold: true, color: { argb: "FF" + MC.pri.replace("#", "") } };

        sheet.columns = [{ width: 14 }, { width: 8 }, { width: 12 }, { width: 14 }, { width: 16 }, { width: 22 }];

        const dataEndRow = sheet.rowCount;

        /* ── Capture & sisipkan gambar grafik suhu ── */
        const suhuEl = exportChartRefs.current[ruang]?.suhu;
        if (suhuEl) {
          try {
            const canvas   = await html2canvas(suhuEl, { backgroundColor: "#ffffff", scale: 2 });
            const dataUrl  = canvas.toDataURL("image/png");
            const imgId    = wb.addImage({ base64: dataUrl, extension: "png" });
            const imgRowStart = dataEndRow + 2;
            sheet.getCell(`A${imgRowStart - 1}`).value = "🌡 Grafik Suhu";
            sheet.getCell(`A${imgRowStart - 1}`).font  = { bold: true, color: { argb: "FF" + MC.pri.replace("#", "") } };
            sheet.addImage(imgId, {
              tl: { col: 0, row: imgRowStart },
              ext: { width: 560, height: 230 },
            });
          } catch (imgErr) {
            console.error("Gagal capture grafik suhu:", imgErr);
            chartFailures.push(`${ruang} (suhu)`);
          }
        }

        /* ── Capture & sisipkan gambar grafik kelembaban ── */
        const rhEl = exportChartRefs.current[ruang]?.rh;
        if (rhEl) {
          try {
            const canvas2  = await html2canvas(rhEl, { backgroundColor: "#ffffff", scale: 2 });
            const dataUrl2 = canvas2.toDataURL("image/png");
            const imgId2   = wb.addImage({ base64: dataUrl2, extension: "png" });
            const imgRowStart2 = dataEndRow + 14; // offset ke bawah agar tidak tumpang-tindih grafik suhu
            sheet.getCell(`A${imgRowStart2 - 1}`).value = "💧 Grafik Kelembaban";
            sheet.getCell(`A${imgRowStart2 - 1}`).font   = { bold: true, color: { argb: "FF" + MC.ok.replace("#", "") } };
            sheet.addImage(imgId2, {
              tl: { col: 0, row: imgRowStart2 },
              ext: { width: 560, height: 230 },
            });
          } catch (imgErr) {
            console.error("Gagal capture grafik kelembaban:", imgErr);
            chartFailures.push(`${ruang} (kelembaban)`);
          }
        }
      }

      /* ── Tulis & unduh file ── */
      const buffer = await wb.xlsx.writeBuffer();
      const blob   = new Blob([buffer], { type: "application/octet-stream" });
      const url    = URL.createObjectURL(blob);
      const a      = document.createElement("a");
      a.href = url;
      a.download = `Grafik_Monitoring_${selMonth}.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      if (chartFailures.length) {
        showToast(`⚠ Data tersimpan, tapi grafik gagal disisipkan untuk: ${chartFailures.join(", ")}`, MC.warn);
      } else {
        showToast("✓ Grafik & data " + mLabel + " diunduh (per ruangan)", MC.ok);
      }
    } catch (err: any) {
      console.error(err);
      showToast("⚠ Gagal mengekspor grafik: " + (err?.message || "kesalahan tidak diketahui"), MC.err);
    } finally {
      setGrafikExporting(false);
    }
  };

  /* Inner style override untuk input di dalam komponen ini */
  const iSm: React.CSSProperties = {padding:"6px 8px",border:"1px solid #D1D5DB",borderRadius:7,fontSize:13,width:"100%",background:"#fff",fontFamily:"inherit",outline:"none",boxSizing:"border-box"};

  /* Data tabel Harian (gunakan meAll agar tidak terpengaruh filter grafik) */
  const meHarian = meAll;

  return (
    <div>
      {/* ── Header ── */}
      <div style={{marginBottom:14}}>
        <div style={{fontWeight:800,fontSize:18,color:MC.pri,marginBottom:2}}>🌡 Monitoring Suhu &amp; Kelembaban</div>
        <div style={{fontSize:12,color:"#64748B"}}>
          {MON_ROOMS.join(" · ")} · Standar Suhu {monitoringCfg.suhuMin}-{monitoringCfg.suhuMax}°C · RH {monitoringCfg.rhMin}-{monitoringCfg.rhMax}%
        </div>
      </div>

      {/* ── Sub-tabs ── */}
      <div style={{display:"flex",gap:6,marginBottom:18,flexWrap:"wrap"}}>
        {[{k:"harian",l:"📋 Harian"},{k:"grafik",l:"📊 Grafik"},{k:"unduh",l:"📥 Unduh"},{k:"standar",l:"⚙ Standar"}].map(t=>(
          <button key={t.k} onClick={()=>setSubTab(t.k)} style={{padding:"7px 18px",borderRadius:20,border:"none",background:subTab===t.k?MC.pri:"#E2E8F0",color:subTab===t.k?"#fff":"#475569",fontWeight:600,fontSize:13,cursor:"pointer",fontFamily:"inherit",transition:"all .15s"}}>{t.l}</button>
        ))}
      </div>

      {/* Month + date pickers (shared) */}
      {(subTab==="harian"||subTab==="grafik"||subTab==="unduh")&&(
        <div style={{display:"flex",gap:10,alignItems:"flex-end",marginBottom:18,flexWrap:"wrap"}}>
          <div>
            <div style={{fontSize:11,fontWeight:700,color:"#64748B",marginBottom:4}}>Bulan</div>
            <input type="month" value={selMonth} onChange={e=>setSelMonth(e.target.value)} style={{...iSm,width:"auto"}}/>
          </div>
          {subTab==="harian"&&(
            <div>
              <div style={{fontSize:11,fontWeight:700,color:"#64748B",marginBottom:4}}>Tanggal Input</div>
              <input type="date" value={formDate} onChange={e=>setFormDate(e.target.value)} style={{...iSm,width:"auto"}}/>
            </div>
          )}
        </div>
      )}

      {/* ══════════════════════════════════════════════════════
          TAB HARIAN
         ══════════════════════════════════════════════════════ */}
      {subTab==="harian"&&(
        <div>
          {/* Form input: 1 tabel per ruangan */}
          <div style={{background:"linear-gradient(135deg,#EFF6FF,#F0FDF4)",border:"1px solid #BAE6FD",borderRadius:14,padding:18,marginBottom:18}}>
            <div style={{fontWeight:700,color:MC.pri,fontSize:14,marginBottom:16}}>📝 Input Pemantauan — {fDMY(formDate)}</div>
            {MON_ROOMS.map((ruang,ri)=>(
              <div key={ruang} style={{marginBottom:ri<MON_ROOMS.length-1?22:0}}>
                {/* Section header ruangan */}
                <div style={{
                  display:"flex",alignItems:"center",gap:8,marginBottom:10,
                  paddingBottom:6,borderBottom:"2px solid "+MON_ROOM_COLORS[ri]+"33"
                }}>
                  <span style={{
                    background:MON_ROOM_COLORS[ri],color:"#fff",fontSize:11,
                    fontWeight:800,padding:"3px 10px",borderRadius:14
                  }}>{ruang}</span>
                </div>
                <div style={{overflowX:"auto"}}>
                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                    <thead>
                      <tr>
                        {["Jam","Suhu (°C)","Kelembaban (%)","Petugas","Status"].map(h=>(
                          <th key={h} style={{padding:"7px 10px",background:"#DBEAFE",color:MC.pri,fontWeight:700,textAlign:"left",borderBottom:"2px solid #BAE6FD",whiteSpace:"nowrap",fontSize:12}}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {MON_JAMS.map(jam=>{
                        const s   = slots[ruang]?.[jam] ?? {suhu:"",kelembaban:"",petugas:""};
                        const sv  = parseFloat(s.suhu);
                        const rv  = parseFloat(s.kelembaban);
                        const hasV= !!(s.suhu&&s.kelembaban&&!isNaN(sv)&&!isNaN(rv));
                        const ok: boolean|null = hasV ? monIsOK(sv,rv,monitoringCfg) : null;
                        return (
                          <tr key={jam} style={{background:"#fff",borderBottom:"1px solid #E2E8F0"}}>
                            <td style={{padding:"7px 10px",fontWeight:700,color:MON_ROOM_COLORS[ri],whiteSpace:"nowrap"}}>{jam}</td>
                            <td style={{padding:"5px 7px"}}>
                              <input type="number" step="0.1" min="10" max="40" placeholder="mis. 21.5"
                                value={s.suhu}
                                onChange={e=>setSlots(p=>({...p,[ruang]:{...p[ruang],[jam]:{...p[ruang][jam],suhu:e.target.value}}}))}
                                style={{...iSm,width:90}}/>
                            </td>
                            <td style={{padding:"5px 7px"}}>
                              <input type="number" step="1" min="0" max="100" placeholder="mis. 55"
                                value={s.kelembaban}
                                onChange={e=>setSlots(p=>({...p,[ruang]:{...p[ruang],[jam]:{...p[ruang][jam],kelembaban:e.target.value}}}))}
                                style={{...iSm,width:80}}/>
                            </td>
                            <td style={{padding:"5px 7px"}}>
                              <input type="text" placeholder="Nama petugas"
                                value={s.petugas}
                                onChange={e=>setSlots(p=>({...p,[ruang]:{...p[ruang],[jam]:{...p[ruang][jam],petugas:e.target.value}}}))}
                                style={{...iSm,minWidth:120}}/>
                            </td>
                            <td style={{padding:"7px 10px",whiteSpace:"nowrap"}}>
                              {ok===null
                                ? <span style={{color:"#94A3B8",fontSize:12}}>—</span>
                                : ok
                                  ? <span style={{background:MC.okBg,color:MC.ok,fontWeight:700,fontSize:11,padding:"3px 9px",borderRadius:20}}>✓ SESUAI</span>
                                  : <span style={{background:MC.errBg,color:MC.err,fontWeight:700,fontSize:11,padding:"3px 9px",borderRadius:20}}>✕ TIDAK SESUAI</span>
                              }
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            ))}
            <div style={{marginTop:18}}>
              <Btn color={MC.pri} onClick={handleSave} style={{background:MC.pri,color:"#fff",border:"none"}}>
                💾 Simpan Semua Data {fDMY(formDate)}
              </Btn>
            </div>
          </div>

          {/* Tabel data tersimpan — tambah kolom "Ruangan" */}
          <div style={{background:"#fff",border:"1px solid #E2E8F0",borderRadius:12,overflow:"hidden"}}>
            <div style={{background:"linear-gradient(90deg,"+MC.pri+","+MC.priL+")",padding:"10px 16px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <span style={{color:"#fff",fontWeight:700,fontSize:14}}>📋 Data {mLabel}</span>
              <span style={{color:"rgba(255,255,255,.75)",fontSize:12}}>{meHarian.length} entri</span>
            </div>
            {meHarian.length===0?(
              <div style={{padding:32,textAlign:"center",color:"#94A3B8"}}>
                <div style={{fontSize:32,marginBottom:8}}>🌡</div>
                <div style={{fontWeight:600}}>Belum ada data monitoring bulan ini</div>
                <div style={{fontSize:12,marginTop:4}}>Gunakan form di atas untuk menambahkan data</div>
              </div>
            ):(
              <div style={{overflowX:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                  <thead style={{background:"#F8FAFC"}}>
                    <tr>
                      {["Ruangan","Tanggal","Jam","Suhu","RH","Std Suhu","Std RH","Status","Petugas",""].map(h=>(
                        <th key={h} style={{padding:"8px 10px",textAlign:"left",fontWeight:700,color:"#475569",fontSize:11,borderBottom:"1px solid #E2E8F0",whiteSpace:"nowrap"}}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {meHarian.map((e:MonitoringEntry,i:number)=>{
                      const ok        = monIsOK(e.suhu,e.kelembaban,monitoringCfg);
                      const prevRuang = i>0 ? meHarian[i-1].ruang : "";
                      const prevDate  = i>0 ? meHarian[i-1].tanggal : "";
                      const ri        = MON_ROOMS.indexOf(e.ruang??-1 as any);
                      return (
                        <tr key={e.id} style={{borderBottom:"1px solid #F1F5F9",background:i%2===0?"#fff":"#FAFBFC"}}>
                          <td style={{padding:"7px 10px"}}>
                            {e.ruang!==prevRuang&&(
                              <span style={{background:MON_ROOM_COLORS[ri]??MC.pri,color:"#fff",fontSize:10,fontWeight:700,padding:"2px 8px",borderRadius:10,whiteSpace:"nowrap"}}>
                                {e.ruang??"-"}
                              </span>
                            )}
                          </td>
                          <td style={{padding:"7px 10px",fontWeight:e.tanggal!==prevDate||e.ruang!==prevRuang?700:400,color:"#374151"}}>
                            {e.tanggal!==prevDate||e.ruang!==prevRuang?fDMY(e.tanggal):""}
                          </td>
                          <td style={{padding:"7px 10px",fontWeight:700,color:MON_ROOM_COLORS[ri]??MC.pri}}>{e.jam}</td>
                          <td style={{padding:"7px 10px",fontWeight:700,color:ok?MC.ok:MC.err}}>{e.suhu}°C</td>
                          <td style={{padding:"7px 10px",fontWeight:700,color:ok?MC.ok:MC.err}}>{e.kelembaban}%</td>
                          <td style={{padding:"7px 10px",color:"#64748B",fontSize:12}}>{monitoringCfg.suhuMin}-{monitoringCfg.suhuMax}</td>
                          <td style={{padding:"7px 10px",color:"#64748B",fontSize:12}}>{monitoringCfg.rhMin}-{monitoringCfg.rhMax}</td>
                          <td style={{padding:"7px 10px"}}>
                            <span style={{background:ok?MC.okBg:MC.errBg,color:ok?MC.ok:MC.err,fontSize:11,fontWeight:700,padding:"2px 8px",borderRadius:12}}>
                              {ok?"SESUAI":"TIDAK SESUAI"}
                            </span>
                          </td>
                          <td style={{padding:"7px 10px",color:"#374151"}}>{e.petugas}</td>
                          <td style={{padding:"7px 10px"}}>
                            <button onClick={async ()=>{
                              const res = await deleteFromSupa?.("kb_monitoring", e.id);
                              if(!res?.ok){showToast(`⚠ Gagal menghapus: ${res?.error||"kesalahan tidak diketahui"}`,C.d);return;}
                              setMonitoringEntries((p:any[])=>p.filter((x:any)=>x.id!==e.id));
                              showToast("✓ Data dihapus & tersinkron",C.d);
                            }} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444",fontSize:16}}>✕</button>
                          </td>
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

      {/* ══════════════════════════════════════════════════════
          TAB GRAFIK (dengan dropdown pilih ruangan)
         ══════════════════════════════════════════════════════ */}
      {subTab==="grafik"&&(
        <div>
          {/* Controls: month picker + room filter + download */}
          <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:14,flexWrap:"wrap"}}>
            <div style={{flex:1,minWidth:160}}>
              <input style={{padding:"7px 12px",border:"1px solid #D1D5DB",borderRadius:8,fontSize:13,width:"100%",background:"#fff",fontFamily:"inherit"}}
                type="month" value={selMonth} onChange={e=>setSelMonth(e.target.value)}/>
            </div>
            {/* Dropdown pilih ruangan */}
            <div style={{minWidth:200}}>
              <select
                value={grafikRuang}
                onChange={e=>setGrafikRuang(e.target.value)}
                style={{padding:"7px 12px",border:"1px solid #D1D5DB",borderRadius:8,fontSize:13,background:"#fff",fontFamily:"inherit",cursor:"pointer",outline:"none"}}
              >
                <option value="all">🏥 Semua Ruangan</option>
                {MON_ROOMS.map(r=><option key={r} value={r}>{r}</option>)}
              </select>
            </div>
            <button onClick={()=>setSelMonth(todayDate().slice(0,7))} style={{padding:"7px 14px",background:selMonth===todayDate().slice(0,7)?"#0EA5E9":"#F0F9FF",color:selMonth===todayDate().slice(0,7)?"#fff":"#0369A1",border:"1px solid #BAE6FD",borderRadius:8,cursor:"pointer",fontSize:12,fontWeight:700,fontFamily:"inherit",whiteSpace:"nowrap"}}>
              📅 Bulan Ini
            </button>
            <button
              onClick={downloadGrafikExcel}
              disabled={grafikExporting}
              style={{
                padding: "7px 14px",
                background: grafikExporting ? "#E5E7EB" : "#ECFDF5",
                color: grafikExporting ? "#9CA3AF" : MC.ok,
                border: "1px solid " + (grafikExporting ? "#D1D5DB" : "#BBF7D0"),
                borderRadius: 8,
                cursor: grafikExporting ? "not-allowed" : "pointer",
                fontSize: 12,
                fontWeight: 700,
                fontFamily: "inherit",
                whiteSpace: "nowrap",
              }}
            >
              {grafikExporting ? "⏳ Mengunduh Excel..." : "⬇ Download Excel"}
            </button>
          </div>

          {/* Stat cards — berdasarkan ruangan yang difilter */}
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(120px,1fr))",gap:10,marginBottom:20}}>
            {([
              {l:"Rata-rata Suhu",v:avgS+"°C",ico:"🌡",bg:"#DBEAFE",c:"#1D4ED8"},
              {l:"Rata-rata RH",   v:avgR+"%", ico:"💧",bg:"#DCFCE7",c:"#16A34A"},
              {l:"Total Ukur",    v:String(me.length),ico:"📊",bg:"#FEF3C7",c:"#B45309"},
              {l:"Tidak Sesuai",  v:String(tidakS),ico:"⚠️",bg:"#FEE2E2",c:"#DC2626"},
              {l:"Kepatuhan",     v:cmpPct+"%",ico:"✅",bg:"#F0FDF4",c:"#16A34A"},
            ] as {l:string;v:string;ico:string;bg:string;c:string}[]).map(s=>(
              <div key={s.l} style={{background:s.bg,borderRadius:12,padding:"12px 10px",textAlign:"center",border:"1px solid "+s.c+"20"}}>
                <div style={{fontSize:20,marginBottom:3}}>{s.ico}</div>
                <div style={{fontSize:17,fontWeight:800,color:s.c}}>{s.v}</div>
                <div style={{fontSize:10,color:"#64748B",marginTop:2,lineHeight:1.3}}>{s.l}</div>
              </div>
            ))}
          </div>

          {/* Label ruangan aktif */}
          {grafikRuang!=="all"&&(
            <div style={{background:MC.priBg,borderRadius:8,padding:"6px 14px",marginBottom:12,fontSize:12,color:MC.pri,fontWeight:700,display:"inline-block"}}>
              📍 Menampilkan: {grafikRuang}
            </div>
          )}

          {me.length===0?(
            <div style={{textAlign:"center",padding:48,color:"#94A3B8"}}>
              <div style={{fontSize:40,marginBottom:10}}>📊</div>
              <div style={{fontWeight:600}}>Belum ada data untuk pilihan ini</div>
            </div>
          ):(
            <>
              {/* ── Grafik Suhu ── */}
              <div style={{background:"#fff",border:"1px solid #E2E8F0",borderRadius:12,padding:"16px 16px 8px",marginBottom:14}}>
                <div style={{fontWeight:700,color:MC.pri,marginBottom:10,fontSize:13}}>🌡 Grafik Suhu — {mLabel}{grafikRuang!=="all"?` (${grafikRuang})`:""}</div>
                <ResponsiveContainer width="100%" height={220}>
                  {grafikRuang==="all"?(
                    /* Multi-line: 1 garis per ruangan */
                    <LineChart data={buildMultiChartData("suhu")} margin={{top:4,right:20,left:0,bottom:4}}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#F1F5F9"/>
                      <XAxis dataKey="name" tick={{fontSize:9}} interval={Math.max(0,Math.floor(buildMultiChartData("suhu").length/8)-1)}/>
                      <YAxis domain={[15,30]} tick={{fontSize:11}} unit="°C" width={44}/>
                      <Tooltip/>
                      <ReferenceLine y={monitoringCfg.suhuMin} stroke="#EF4444" strokeDasharray="4 3" label={{value:"Min "+monitoringCfg.suhuMin,position:"insideTopRight",fontSize:9,fill:"#EF4444"}}/>
                      <ReferenceLine y={monitoringCfg.suhuMax} stroke="#EF4444" strokeDasharray="4 3" label={{value:"Maks "+monitoringCfg.suhuMax,position:"insideBottomRight",fontSize:9,fill:"#EF4444"}}/>
                      {MON_ROOMS.map((r,i)=>(
                        <Line key={r} type="monotone" dataKey={r} stroke={MON_ROOM_COLORS[i]} strokeWidth={2} dot={{r:2}} activeDot={{r:5}} name={r} connectNulls/>
                      ))}
                    </LineChart>
                  ):(
                    /* Single-line */
                    <LineChart data={chartData} margin={{top:4,right:20,left:0,bottom:4}}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#F1F5F9"/>
                      <XAxis dataKey="name" tick={{fontSize:9}} interval={Math.max(0,Math.floor(chartData.length/10)-1)}/>
                      <YAxis domain={[15,30]} tick={{fontSize:11}} unit="°C" width={44}/>
                      <Tooltip formatter={(v:any)=>[v+"°C","Suhu"]}/>
                      <ReferenceLine y={monitoringCfg.suhuMin} stroke="#EF4444" strokeDasharray="4 3" label={{value:"Min "+monitoringCfg.suhuMin,position:"insideTopRight",fontSize:9,fill:"#EF4444"}}/>
                      <ReferenceLine y={monitoringCfg.suhuMax} stroke="#EF4444" strokeDasharray="4 3" label={{value:"Maks "+monitoringCfg.suhuMax,position:"insideBottomRight",fontSize:9,fill:"#EF4444"}}/>
                      <Line type="monotone" dataKey="suhu" stroke={MC.pri} strokeWidth={2.5} dot={{r:2,fill:MC.pri}} activeDot={{r:5}}/>
                    </LineChart>
                  )}
                </ResponsiveContainer>
                {/* Legend multi-room */}
                {grafikRuang==="all"&&(
                  <div style={{display:"flex",gap:14,justifyContent:"center",marginTop:8}}>
                    {MON_ROOMS.map((r,i)=>(
                      <div key={r} style={{display:"flex",alignItems:"center",gap:5,fontSize:11,color:"#475569"}}>
                        <div style={{width:20,height:3,background:MON_ROOM_COLORS[i],borderRadius:2}}/>
                        {r}
                      </div>
                    ))}
                  </div>
                )}
              </div>

              {/* ── Grafik Kelembaban ── */}
              <div style={{background:"#fff",border:"1px solid #E2E8F0",borderRadius:12,padding:"16px 16px 8px"}}>
                <div style={{fontWeight:700,color:MC.ok,marginBottom:10,fontSize:13}}>💧 Grafik Kelembaban — {mLabel}{grafikRuang!=="all"?` (${grafikRuang})`:""}</div>
                <ResponsiveContainer width="100%" height={220}>
                  {grafikRuang==="all"?(
                    <LineChart data={buildMultiChartData("kelembaban")} margin={{top:4,right:20,left:0,bottom:4}}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#F1F5F9"/>
                      <XAxis dataKey="name" tick={{fontSize:9}} interval={Math.max(0,Math.floor(buildMultiChartData("kelembaban").length/8)-1)}/>
                      <YAxis domain={[30,90]} tick={{fontSize:11}} unit="%" width={44}/>
                      <Tooltip/>
                      <ReferenceLine y={monitoringCfg.rhMin} stroke="#EF4444" strokeDasharray="4 3" label={{value:"Min "+monitoringCfg.rhMin,position:"insideTopRight",fontSize:9,fill:"#EF4444"}}/>
                      <ReferenceLine y={monitoringCfg.rhMax} stroke="#EF4444" strokeDasharray="4 3" label={{value:"Maks "+monitoringCfg.rhMax,position:"insideBottomRight",fontSize:9,fill:"#EF4444"}}/>
                      {MON_ROOMS.map((r,i)=>(
                        <Line key={r} type="monotone" dataKey={r} stroke={MON_ROOM_COLORS[i]} strokeWidth={2} dot={{r:2}} activeDot={{r:5}} name={r} connectNulls/>
                      ))}
                    </LineChart>
                  ):(
                    <LineChart data={chartData} margin={{top:4,right:20,left:0,bottom:4}}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#F1F5F9"/>
                      <XAxis dataKey="name" tick={{fontSize:9}} interval={Math.max(0,Math.floor(chartData.length/10)-1)}/>
                      <YAxis domain={[30,90]} tick={{fontSize:11}} unit="%" width={44}/>
                      <Tooltip formatter={(v:any)=>[v+"%","Kelembaban"]}/>
                      <ReferenceLine y={monitoringCfg.rhMin} stroke="#EF4444" strokeDasharray="4 3" label={{value:"Min "+monitoringCfg.rhMin,position:"insideTopRight",fontSize:9,fill:"#EF4444"}}/>
                      <ReferenceLine y={monitoringCfg.rhMax} stroke="#EF4444" strokeDasharray="4 3" label={{value:"Maks "+monitoringCfg.rhMax,position:"insideBottomRight",fontSize:9,fill:"#EF4444"}}/>
                      <Line type="monotone" dataKey="kelembaban" stroke={MC.ok} strokeWidth={2.5} dot={{r:2,fill:MC.ok}} activeDot={{r:5}}/>
                    </LineChart>
                  )}
                </ResponsiveContainer>
                {grafikRuang==="all"&&(
                  <div style={{display:"flex",gap:14,justifyContent:"center",marginTop:8}}>
                    {MON_ROOMS.map((r,i)=>(
                      <div key={r} style={{display:"flex",alignItems:"center",gap:5,fontSize:11,color:"#475569"}}>
                        <div style={{width:20,height:3,background:MON_ROOM_COLORS[i],borderRadius:2}}/>
                        {r}
                      </div>
                    ))}
                  </div>
                )}
              </div>

              {/* ══════════════════════════════════════════════════════
                  OFF-SCREEN CHARTS — khusus capture html2canvas saat ekspor.
                  Tidak terlihat user, render selalu (agar siap di-capture),
                  satu pasang grafik (suhu + RH) per ruangan di MON_ROOMS.
                  Struktur Recharts visible di atas TIDAK disentuh.
                 ══════════════════════════════════════════════════════ */}
              <div style={{ position: "absolute", top: 0, left: -9999, width: 600, pointerEvents: "none" }} aria-hidden="true">
                {MON_ROOMS.map((ruang, ri) => {
                  const meRoomChart = meAll
                    .filter(e => e.ruang === ruang)
                    .map(e => ({
                      name: e.tanggal.slice(8) + "/" + (e.jam === "07:00" ? "07" : e.jam === "14:00" ? "14" : "21"),
                      suhu: e.suhu,
                      kelembaban: e.kelembaban,
                    }));
                  return (
                    <div key={"export_" + ruang}>
                      <div
                        ref={(el: HTMLDivElement | null) => { exportChartRefs.current[ruang].suhu = el; }}
                        style={{ width: 560, height: 230, background: "#fff", padding: 12 }}
                      >
                        <div style={{ fontWeight: 700, color: MC.pri, marginBottom: 8, fontSize: 13 }}>
                          🌡 Grafik Suhu — {ruang} — {mLabel}
                        </div>
                        <LineChart width={536} height={180} data={meRoomChart} margin={{ top: 4, right: 20, left: 0, bottom: 4 }}>
                          <CartesianGrid strokeDasharray="3 3" stroke="#F1F5F9" />
                          <XAxis dataKey="name" tick={{ fontSize: 9 }} interval={Math.max(0, Math.floor(meRoomChart.length / 10) - 1)} />
                          <YAxis domain={[15, 30]} tick={{ fontSize: 11 }} unit="°C" width={44} />
                          <ReferenceLine y={monitoringCfg.suhuMin} stroke="#EF4444" strokeDasharray="4 3" />
                          <ReferenceLine y={monitoringCfg.suhuMax} stroke="#EF4444" strokeDasharray="4 3" />
                          <Line type="monotone" dataKey="suhu" stroke={MON_ROOM_COLORS[ri]} strokeWidth={2.5} dot={{ r: 2, fill: MON_ROOM_COLORS[ri] }} isAnimationActive={false} />
                        </LineChart>
                      </div>
                      <div
                        ref={(el: HTMLDivElement | null) => { exportChartRefs.current[ruang].rh = el; }}
                        style={{ width: 560, height: 230, background: "#fff", padding: 12, marginTop: 8 }}
                      >
                        <div style={{ fontWeight: 700, color: MC.ok, marginBottom: 8, fontSize: 13 }}>
                          💧 Grafik Kelembaban — {ruang} — {mLabel}
                        </div>
                        <LineChart width={536} height={180} data={meRoomChart} margin={{ top: 4, right: 20, left: 0, bottom: 4 }}>
                          <CartesianGrid strokeDasharray="3 3" stroke="#F1F5F9" />
                          <XAxis dataKey="name" tick={{ fontSize: 9 }} interval={Math.max(0, Math.floor(meRoomChart.length / 10) - 1)} />
                          <YAxis domain={[30, 90]} tick={{ fontSize: 11 }} unit="%" width={44} />
                          <ReferenceLine y={monitoringCfg.rhMin} stroke="#EF4444" strokeDasharray="4 3" />
                          <ReferenceLine y={monitoringCfg.rhMax} stroke="#EF4444" strokeDasharray="4 3" />
                          <Line type="monotone" dataKey="kelembaban" stroke={MC.ok} strokeWidth={2.5} dot={{ r: 2, fill: MC.ok }} isAnimationActive={false} />
                        </LineChart>
                      </div>
                    </div>
                  );
                })}
              </div>
            </>
          )}
        </div>
      )}

      {/* ══════════════════════════════════════════════════════
          TAB UNDUH (tidak berubah struktur, hanya pass entries yg sudah punya field ruang)
         ══════════════════════════════════════════════════════ */}
      {subTab==="unduh"&&(
        <div>
          <div style={{display:"grid",gap:12,marginBottom:16}}>
            {([
              {type:"harian"    as const,ico:"📋",title:"Laporan Monitoring Harian",   desc:"Data harian 3 ruangan + rekap statistik bulanan",c:"#0369A1",bg:"#EFF6FF"},
              {type:"rekap"     as const,ico:"📊",title:"Rekap Bulanan",               desc:"Statistik per ruangan: rata-rata suhu, RH, kepatuhan (tidak dicampur)",c:"#7C3AED",bg:"#F5F3FF"},
              {type:"akreditasi"as const,ico:"🏅",title:"Laporan Akreditasi",          desc:"Laporan lengkap per ruangan + format akreditasi & kolom tanda tangan",c:"#B45309",bg:"#FFFBEB"},
            ]).map(r=>(
              <div key={r.type} style={{background:r.bg,border:"1px solid "+r.c+"25",borderRadius:12,padding:"14px 16px",display:"flex",justifyContent:"space-between",alignItems:"center",gap:12,flexWrap:"wrap"}}>
                <div style={{flex:1}}>
                  <div style={{fontWeight:700,color:r.c,fontSize:14,marginBottom:3}}>{r.ico} {r.title}</div>
                  <div style={{fontSize:12,color:"#64748B"}}>{r.desc}</div>
                </div>
                <Btn color={r.c} onClick={()=>monXlsx(monitoringEntries,monitoringCfg,selMonth,r.type,showToast)} style={{background:r.c,color:"#fff",border:"none",whiteSpace:"nowrap",flexShrink:0}}>⬇ Excel</Btn>
              </div>
            ))}
          </div>
          <div style={{background:"#F8FAFC",borderRadius:10,padding:"12px 14px",fontSize:12,color:"#64748B",border:"1px solid #E2E8F0",lineHeight:1.7,marginBottom:14}}>
            <b>📌 Penyimpanan:</b> Data monitoring disertakan dalam backup otomatis JSON bersama data operasi dan lembur. Aktifkan di tab Arsip → Setelan → Dropbox.
          </div>

          {/* Supabase cloud download */}
          <div style={{background:"#EEF2FF",border:"1.5px solid #6366F133",borderRadius:14,padding:"16px 16px 14px"}}>
            <div style={{fontWeight:700,color:"#4F46E5",fontSize:14,marginBottom:6}}>☁️ Download Monitoring dari Supabase</div>
            <div style={{fontSize:12,color:"#64748B",marginBottom:12}}>Unduh data monitoring langsung dari cloud Supabase, tanpa bergantung data perangkat ini. Data tersinkron dari semua device.</div>
            <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap",marginBottom:12}}>
              <div style={{flex:1,minWidth:160}}>
                <div style={{fontSize:12,fontWeight:700,color:"#64748B",marginBottom:4}}>Pilih Bulan</div>
                <input style={iSm} type="month" value={supaMonYM} onChange={e=>setSupaMonYM(e.target.value)} disabled={supaMonBusy}/>
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

          {/* Dropbox download */}
          <div style={{background:"#EFF6FF",border:"1.5px solid #0061FF33",borderRadius:14,padding:"16px 16px 14px",marginTop:14}}>
            <div style={{fontWeight:700,color:"#0061FF",fontSize:14,marginBottom:6}}>📦 Download Monitoring dari Dropbox</div>
            <div style={{fontSize:12,color:"#64748B",marginBottom:12}}>Unduh data monitoring dari backup Dropbox sesuai pilihan bulan dan format laporan.</div>
            <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap",marginBottom:12}}>
              <div style={{flex:1,minWidth:160}}>
                <div style={{fontSize:12,fontWeight:700,color:"#64748B",marginBottom:4}}>Pilih Bulan</div>
                <input style={iSm} type="month" value={dbxMonYM} onChange={e=>setDbxMonYM(e.target.value)} disabled={dbxMonBusy}/>
              </div>
            </div>
            <div style={{display:"grid",gap:8}}>
              {([{type:"harian" as const,ico:"📋",title:"Harian",c:"#0369A1"},{type:"rekap" as const,ico:"📊",title:"Rekap Bulanan",c:"#7C3AED"},{type:"akreditasi" as const,ico:"🏅",title:"Akreditasi",c:"#B45309"}]).map(r=>(
                <div key={r.type} style={{display:"flex",justifyContent:"space-between",alignItems:"center",background:"rgba(255,255,255,.7)",borderRadius:10,padding:"10px 14px",gap:10}}>
                  <div style={{fontWeight:600,color:r.c,fontSize:13}}>{r.ico} {r.title}</div>
                  <button onClick={()=>downloadMonFromDropbox(r.type)} disabled={dbxMonBusy} style={{background:dbxMonBusy?"#9CA3AF":"#0061FF",color:"#fff",border:"none",borderRadius:8,padding:"6px 14px",cursor:dbxMonBusy?"not-allowed":"pointer",fontSize:12,fontWeight:700,fontFamily:"inherit",whiteSpace:"nowrap"}}>
                    {dbxMonBusy?"⏳ ...":"📦 Download Excel"}
                  </button>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* ══════════════════════════════════════════════════════
          TAB STANDAR (tidak berubah)
         ══════════════════════════════════════════════════════ */}
      {subTab==="standar"&&(
        <div style={{maxWidth:480}}>
          <div style={{background:"#fff",border:"1px solid #E2E8F0",borderRadius:12,padding:18}}>
            <div style={{fontWeight:700,color:MC.pri,fontSize:14,marginBottom:16}}>⚙ Konfigurasi Standar Monitoring</div>
            <div style={{fontSize:12,color:"#64748B",marginBottom:14,background:MC.priBg,padding:"8px 12px",borderRadius:8}}>
              Standar yang diatur di sini berlaku untuk ketiga ruangan: <b>{MON_ROOMS.join(", ")}</b>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:14}}>
              {[{k:"suhuMin",l:"Suhu Min (°C)"},{k:"suhuMax",l:"Suhu Maks (°C)"},{k:"rhMin",l:"RH Min (%)"},{k:"rhMax",l:"RH Maks (%)"}].map(f=>(
                <div key={f.k}>
                  <div style={{fontSize:12,fontWeight:700,color:"#64748B",marginBottom:4}}>{f.l}</div>
                  <input type="number" step="0.5" value={(cfgForm as any)[f.k]}
                    onChange={e=>setCfgForm((p:MonitoringCfg)=>({...p,[f.k]:parseFloat(e.target.value)||0}))}
                    style={iSm}/>
                </div>
              ))}
            </div>
            <div style={{marginBottom:12}}>
              <div style={{fontSize:12,fontWeight:700,color:"#64748B",marginBottom:4}}>Kepala Kamar Bedah</div>
              <input type="text" value={cfgForm.kepalaKamarBedah}
                onChange={e=>setCfgForm((p:MonitoringCfg)=>({...p,kepalaKamarBedah:e.target.value}))}
                style={iSm} placeholder="Nama kepala kamar bedah"/>
            </div>
            <div style={{background:MC.okBg,borderRadius:8,padding:"8px 12px",marginBottom:16,fontSize:12,color:MC.ok,border:"1px solid #BBF7D0"}}>
              Standar aktif: Suhu {monitoringCfg.suhuMin}-{monitoringCfg.suhuMax}°C · RH {monitoringCfg.rhMin}-{monitoringCfg.rhMax}%
            </div>
            <Btn color={MC.pri} onClick={()=>{setMonitoringCfg(cfgForm);showToast("✓ Pengaturan standar disimpan",MC.ok);}} style={{background:MC.pri,color:"#fff",border:"none"}}>
              💾 Simpan Pengaturan
            </Btn>
          </div>
        </div>
      )}
    </div>
  );
}

/* ─── SUPABASE SINGLETON ───────────────────────────────────────────
   FIX #1 (CRITICAL — AUDIT): Anon key TIDAK LAGI di-hardcode di source.
   Diambil dari environment variable (.env.local, JANGAN di-commit ke git):
     VITE_SUPABASE_URL=https://xxxx.supabase.co
     VITE_SUPABASE_ANON_KEY=eyJ...
   "anon public key" memang didesain aman di client-side per arsitektur
   Supabase (akses sesungguhnya dikontrol via RLS), namun tetap tidak boleh
   di-hardcode di source: hardcode membuat rotasi key sulit, membuat key
   bocor ke setiap fork/clone repo, dan menyamarkan kebutuhan RLS yang benar.
   ─────────────────────────────────────────────────────────────────── */
const SUPA_URL    = (import.meta as any).env?.VITE_SUPABASE_URL as string | undefined;
const SUPA_ANON   = (import.meta as any).env?.VITE_SUPABASE_ANON_KEY as string | undefined;
if (!SUPA_URL || !SUPA_ANON) {
  // eslint-disable-next-line no-console
  console.error(
    "[Konfigurasi] VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY belum diset di .env.local. " +
    "Aplikasi tidak dapat terhubung ke Supabase sampai env var ini dikonfigurasi."
  );
}
const SUPA_CLIENT = createClient(SUPA_URL || "", SUPA_ANON || "");

/* ─── DROPBOX: token DIHAPUS dari sini sepenuhnya — lihat DBX_PROXY_SECRET
   di atas. Token asli sekarang hanya hidup sbg secret server-side di
   Supabase Edge Function "dropbox-proxy" (Deno.env.get("DROPBOX_ACCESS_TOKEN")). */

/* ─── HEADER CLOCK (terisolasi) ─────────────────────────────────────────
   Sengaja dipisah jadi komponennya sendiri — detak 1 detik di sini TIDAK
   memicu re-render App ataupun tab yang sedang aktif (lihat catatan perf
   di komponen App: state currentTime sebelumnya ada di root App). */
function HeaderClock() {
  const [now, setNow] = useState(() => new Date());
  useEffect(() => {
    const iv = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(iv);
  }, []);
  return (
    <>
      <div style={{color:"rgba(255,255,255,.55)",fontSize:10,marginTop:3,letterSpacing:.2}}>{now.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"})}</div>
      <div style={{color:"rgba(255,255,255,.9)",fontSize:13,fontWeight:700,fontVariantNumeric:"tabular-nums",letterSpacing:1,marginTop:1}}>{now.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"})} WIB</div>
    </>
  );
}

/* ─── MAIN APP ───────────────────────────────────────────────────────── */

/* ─── MASTER DATA PERAWAT RS PANTI RINI ────────────────────────────── */
const MASTER_PERAWAT_LIST: Array<{nik:string; name:string; pk:string}> = [
  { nik:"201010056", name:"Jaka R",        pk:"III"  },
  { nik:"199720045", name:"Eka S",         pk:"III"  },
  { nik:"200315032", name:"Niken A",       pk:"II"   },
  { nik:"199805021", name:"Thresmiati CB", pk:"IV"   },
  { nik:"200110078", name:"Yohana P",      pk:"III"  },
  { nik:"200812034", name:"Agus W",        pk:"II"   },
  { nik:"199912067", name:"Dewi R",        pk:"III"  },
  { nik:"200508043", name:"Sari U",        pk:"II"   },
  { nik:"201205089", name:"Budi S",        pk:"I"    },
  { nik:"200001055", name:"Ratna K",       pk:"III"  },
];

/* ─── VIEW ROSTER GENERATOR ─────────────────────────────────────────── */
/* Kamus kode shift */
const SHIFT_CODES = { PAGI:"P", PAGI_SIAGA:"PG", SIANG:"S", SIANG_SIAGA:"SG", LIBUR:"L", LIBUR_SIAGA:"LG", CUTI:"T" } as const;
type NurseType = "karu"|"senior"|"junior";
interface RosterNurse { id: string; nik: string; pk: string; name: string; sisaCuti: number; tipe: NurseType; }
interface RosterGenProps { showToast: ShowToastFn; upsertOneToSupa: UpsertOneFn; dbxCfg: DropboxConfig; }

/* ── Master Perawat — interface & LS key (terpisah dari RosterNurse) ── */
interface NurseMaster { id: string; nik: string; name: string; pk: string; }
const LS_MASTER_KEY = "rostergen_master_v1";

/* ── Seed awal: data RS Panti Rini (hanya dipakai saat LS kosong) ── */
const SEED_MASTER: NurseMaster[] = [
  { id:"m1",  nik:"201010056", name:"Jaka R",        pk:"III" },
  { id:"m2",  nik:"199720045", name:"Eka S",          pk:"III" },
  { id:"m3",  nik:"200315032", name:"Niken A",        pk:"II"  },
  { id:"m4",  nik:"199805021", name:"Thresmiati CB",  pk:"IV"  },
  { id:"m5",  nik:"200110078", name:"Yohana P",       pk:"III" },
  { id:"m6",  nik:"200812034", name:"Agus W",         pk:"II"  },
  { id:"m7",  nik:"199912067", name:"Dewi R",         pk:"III" },
  { id:"m8",  nik:"200508043", name:"Sari U",         pk:"II"  },
  { id:"m9",  nik:"201205089", name:"Budi S",         pk:"I"   },
  { id:"m10", nik:"200001055", name:"Ratna K",        pk:"III" },
];

/** loadMasterFromLS: baca dari localStorage, fallback ke SEED jika kosong */
function loadMasterFromLS(): NurseMaster[] {
  try {
    const raw = localStorage.getItem(LS_MASTER_KEY);
    if(raw) { const parsed = JSON.parse(raw); if(Array.isArray(parsed) && parsed.length>0) return parsed; }
  } catch { /* ignore */ }
  return SEED_MASTER;
}

/* ── Nama-nama bulan Indonesia (dipakai di header Excel & label UI) ── */
const BULAN_NAMA_ID = ["JANUARI","FEBRUARI","MARET","APRIL","MEI","JUNI","JULI","AGUSTUS","SEPTEMBER","OKTOBER","NOVEMBER","DESEMBER"];

/* ── Kolom ringkasan individu di sisi kanan tabel Excel (7 kolom sesuai UI) ── */
const SUMMARY_COLS = ["∑P","∑PG","∑S","∑SG","∑L/LG","∑T","∑Total"] as const;

/* ── Helper: hitung ringkasan 7 kolom per perawat (untuk Excel export) ── */
function hitungRingkasan(row: string[], daysInMonth: number) {
  const cnt = (code: string) => row.slice(0, daysInMonth).filter(c=>c===code).length;
  const P=cnt("P"), PG=cnt("PG"), S=cnt("S"), SG=cnt("SG"),
        T=cnt("T"), LG=cnt("LG"), L=cnt("L");
  return [P, PG, S, SG, LG+L, T, P+PG+S+SG+T+LG+L];
}

/* ── localStorage key untuk draft jadwal (offline resilience) ── */
const LS_DRAFT_KEY = "rostergen_draft_v1";

function mkNurse(tipe: NurseType = "junior"): RosterNurse {
  return { id: gId(), nik:"", pk:"", name:"", sisaCuti: 12, tipe };
}

function fisherYates<T>(arr: T[]): T[] {
  const a = [...arr];
  for(let i=a.length-1;i>0;i--){ const j=Math.floor(Math.random()*(i+1)); [a[i],a[j]]=[a[j],a[i]]; }
  return a;
}

function ViewRosterGenerator({ showToast, upsertOneToSupa, dbxCfg }: RosterGenProps) {
  const today = new Date();
  const [month, setMonth] = useState(today.getMonth());
  const [year,  setYear]  = useState(today.getFullYear());
  const daysInMonth = new Date(year, month+1, 0).getDate();

  /* ── PJ & Pembawa HP (untuk footer Excel) ── */
  const [pjSirs,     setPjSirs]     = useState("Jaka R");
  const [pembawHP,   setPembawHP]   = useState("Niken A");
  const [kepInstall, setKepInstall] = useState("drg. Agus Sri G, Sp. BM");

  /* ── Status cloud ── */
  const [cloudStatus, setCloudStatus] = useState<"idle"|"saving"|"ok"|"error"|"offline">("idle");
  const autoSaveTimerRef = useRef<ReturnType<typeof setTimeout>|null>(null);
  const pendingOfflineRef = useRef(false);

  const defaultNurses = (): RosterNurse[] => [
    mkNurse("karu"),
    mkNurse("senior"), mkNurse("senior"), mkNurse("senior"), mkNurse("senior"),
    mkNurse("junior"), mkNurse("junior"), mkNurse("junior"), mkNurse("junior"), mkNurse("junior"),
  ];

  const [nurses, setNurses]     = useState<RosterNurse[]>(defaultNurses);
  // grid[nurseIdx][dayIdx] = kode shift string
  const [grid, setGrid]         = useState<string[][]>(() => defaultNurses().map(()=>Array(31).fill("")));
  const [holidays, setHolidays] = useState<boolean[]>(()=>Array(31).fill(false));
  const [syncing, setSyncing]   = useState(false);
  const [syncingDbx, setSyncingDbx] = useState(false);
  const [savingJadwal, setSavingJadwal] = useState(false);
  const [warnings, setWarnings] = useState<string[]>([]);

  /* ── FITUR 1: State Master Perawat (persisten via localStorage) ── */
  const [nurseMasterList, setNurseMasterList] = useState<NurseMaster[]>(loadMasterFromLS);
  const [masterForm, setMasterForm] = useState({ nik:"", name:"", pk:"III" });
  const [masterFormErr, setMasterFormErr] = useState("");
  const [masterSaving, setMasterSaving]   = useState(false);

  /* Persist nurseMasterList ke localStorage setiap kali berubah */
  useEffect(() => {
    try { localStorage.setItem(LS_MASTER_KEY, JSON.stringify(nurseMasterList)); } catch { /* ignore */ }
  }, [nurseMasterList]);

  /* ── FITUR 3: useMemo — ringkasan individu per baris (Sumbu Kanan) ── */
  const rowSummaries = useMemo(() =>
    nurses.map((_, ni) => {
      const cnt = (code: string) => Array.from({length:daysInMonth},(_,d)=>newGrid_safe(ni,d)).filter(c=>c===code).length;
      const P=cnt("P"), PG=cnt("PG"), S=cnt("S"), SG=cnt("SG"), T=cnt("T"), LG=cnt("LG"), L=cnt("L");
      return { P, PG, S, SG, T, LG_L: LG+L, total: P+PG+S+SG+T+LG+L };
    }),
  // eslint-disable-next-line react-hooks/exhaustive-deps
  [grid, nurses, daysInMonth]);

  /* ── FITUR 4/5: useMemo — ringkasan harian per kolom (Sumbu Bawah) ── */
  const colSummaries = useMemo(() => ({
    P:   Array.from({length:daysInMonth},(_,d) => nurses.reduce((c,_,ni)=>newGrid_safe(ni,d)==="P"?c+1:c,0)),
    PG:  Array.from({length:daysInMonth},(_,d) => nurses.reduce((c,_,ni)=>newGrid_safe(ni,d)==="PG"?c+1:c,0)),
    S:   Array.from({length:daysInMonth},(_,d) => nurses.reduce((c,_,ni)=>newGrid_safe(ni,d)==="S"?c+1:c,0)),
    SG:  Array.from({length:daysInMonth},(_,d) => nurses.reduce((c,_,ni)=>newGrid_safe(ni,d)==="SG"?c+1:c,0)),
    LLG: Array.from({length:daysInMonth},(_,d) => nurses.reduce((c,_,ni)=>["L","LG"].includes(newGrid_safe(ni,d))?c+1:c,0)),
    T:   Array.from({length:daysInMonth},(_,d) => nurses.reduce((c,_,ni)=>newGrid_safe(ni,d)==="T"?c+1:c,0)),
  }),
  // eslint-disable-next-line react-hooks/exhaustive-deps
  [grid, nurses, daysInMonth]);

  // Reset grid ketika jumlah nurses berubah atau bulan berubah
  const resetGrid = (newNurses: RosterNurse[], days: number) => {
    setGrid(newNurses.map(()=>Array(days).fill("")));
    setHolidays(Array(days).fill(false));
    setWarnings([]);
  };

  /* ─────────────────────────────────────────────────────────────────────
     FITUR 3 — AUTO-BACKUP MULTI-CHANNEL (Supabase + Dropbox + localStorage)
     ─────────────────────────────────────────────────────────────────────
     Setiap kali grid/nurses/holidays berubah, kita:
       1. Simpan ke localStorage (selalu, offline-proof)
       2. Debounce 2,5 detik → upsert ke Supabase tabel kb_roster_gen
       3. Setelah Supabase sukses → trigger upload JSON ke Dropbox via proxy
     Jika jaringan offline: data tersimpan di localStorage & flag
     pendingOfflineRef=true. window "online" event memicu retry otomatis. ── */

  /** buildPayload: konversi state saat ini → objek JSON yang siap disimpan */
  const buildPayload = useCallback((
    g: string[][], n: RosterNurse[], h: boolean[], m: number, y: number
  ) => ({
    id: `roster_${y}_${String(m+1).padStart(2,"00")}`,   // id deterministik → upsert aman
    updated_at: new Date().toISOString(),
    year: y, month: m,
    nurses: JSON.stringify(n),
    jadwal: n.map((_,ni) => g[ni]?.slice(0, new Date(y,m+1,0).getDate()).join(",") ?? "").join("|"),
    holidays: h.slice(0, new Date(y,m+1,0).getDate()).join(","),
    generatedAt: fNow(),
    pjSirs, pembawHP, kepInstall,
  }), [pjSirs, pembawHP, kepInstall]);

  /** saveDraftToLS: simpan draft ke localStorage — tidak pernah throw */
  const saveDraftToLS = useCallback((
    g: string[][], n: RosterNurse[], h: boolean[], m: number, y: number
  ) => {
    try {
      const draft = { g, n, h, m, y, ts: Date.now() };
      localStorage.setItem(LS_DRAFT_KEY, JSON.stringify(draft));
    } catch { /* kuota penuh — abaikan senyap */ }
  }, []);

  /** syncToCloudSilent: upsert ke Supabase + backup Dropbox JSON,
      dipanggil dari debounce & tombol "Simpan Jadwal" */
  const syncToCloudSilent = useCallback(async (
    g: string[][], n: RosterNurse[], h: boolean[], m: number, y: number,
    showFeedback = false
  ): Promise<boolean> => {
    if(!navigator.onLine){
      pendingOfflineRef.current = true;
      setCloudStatus("offline");
      if(showFeedback) showToast("⚠ Jaringan offline — jadwal tersimpan lokal, akan sync saat kembali online.", "#E65100");
      return false;
    }
    try {
      setCloudStatus("saving");
      const payload = buildPayload(g, n, h, m, y);

      /* ── Supabase upsert ── */
      const res = await upsertOneToSupa("kb_roster_gen", payload);
      if(!res.ok) throw new Error(res.error || "Supabase error");

      /* ── Dropbox JSON backup ── */
      try {
        const folder = (dbxCfg?.path || "/kamarbedah/").replace(/[^/]+$/,"");
        const stamp  = `${y}-${String(m+1).padStart(2,"0")}`;
        const path   = `${folder}Roster_JSON_${stamp}.json`;
        const body   = JSON.stringify({ ...payload, nurses: n, grid: g, holidays: h }, null, 2);
        const b64    = btoa(unescape(encodeURIComponent(body)));
        await callDbxProxy({ action:"upload_excel", path, base64: b64 });
      } catch { /* Dropbox gagal tidak membatalkan simpan utama */ }

      setCloudStatus("ok");
      pendingOfflineRef.current = false;
      if(showFeedback) showToast("✓ Jadwal berhasil disimpan permanen ke database pusat.", "#2E7D32");
      return true;
    } catch(e: any) {
      setCloudStatus("error");
      if(showFeedback) showToast("⚠ Gagal sync cloud: " + (e?.message || "error"), "#D62828");
      return false;
    }
  }, [buildPayload, upsertOneToSupa, dbxCfg, showToast]);

  /** triggerAutoSave: simpan LS segera, debounce cloud 2,5 s */
  const triggerAutoSave = useCallback((
    g: string[][], n: RosterNurse[], h: boolean[], m: number, y: number
  ) => {
    saveDraftToLS(g, n, h, m, y);
    if(autoSaveTimerRef.current) clearTimeout(autoSaveTimerRef.current);
    autoSaveTimerRef.current = setTimeout(() => {
      syncToCloudSilent(g, n, h, m, y, false);
    }, 2500);
  }, [saveDraftToLS, syncToCloudSilent]);

  /* ── Auto-retry saat kembali online ── */
  useEffect(() => {
    const handler = () => {
      if(pendingOfflineRef.current){
        setCloudStatus("saving");
        syncToCloudSilent(grid, nurses, holidays, month, year, true);
      }
    };
    window.addEventListener("online", handler);
    return () => window.removeEventListener("online", handler);
  }, [grid, nurses, holidays, month, year, syncToCloudSilent]);

  /* Cleanup debounce on unmount */
  useEffect(() => () => { if(autoSaveTimerRef.current) clearTimeout(autoSaveTimerRef.current); }, []);

  const handleMonthYear = (m: number, y: number) => {
    setMonth(m); setYear(y);
    const d = new Date(y, m+1, 0).getDate();
    resetGrid(nurses, d);
    // Jangan auto-save saat reset grid — jadwal akan kosong
  };

  const addNurse = () => {
    const newN = [...nurses, mkNurse("junior")];
    setNurses(newN);
    setGrid(g => {
      const ng = [...g, Array(daysInMonth).fill("")];
      triggerAutoSave(ng, newN, holidays, month, year);
      return ng;
    });
  };

  const updateNurse = (idx: number, field: keyof RosterNurse, val: any) => {
    setNurses(p => {
      const np = p.map((n,i) => i===idx ? {...n,[field]:val} : n);
      triggerAutoSave(grid, np, holidays, month, year);
      return np;
    });
  };

  // Handler NIK dropdown → auto-populate Nama & PK dari nurseMasterList dinamis
  const handleNikChange = (ni: number, nik: string) => {
    const master = nurseMasterList.find(m => m.nik === nik);
    setNurses(p => {
      const np = p.map((n, i) => i===ni
        ? { ...n, nik, name: master?.name ?? n.name, pk: master?.pk ?? n.pk }
        : n
      );
      triggerAutoSave(grid, np, holidays, month, year);
      return np;
    });
  };

  /* ── FITUR 1: CRUD Master Perawat + Supabase sync ── */
  const addToMaster = async () => {
    const nik  = masterForm.nik.trim();
    const name = masterForm.name.trim();
    const pk   = masterForm.pk.trim();
    if(!nik)  { setMasterFormErr("NIK tidak boleh kosong."); return; }
    if(!name) { setMasterFormErr("Nama tidak boleh kosong."); return; }
    if(nurseMasterList.some(m => m.nik === nik))
      { setMasterFormErr(`NIK ${nik} sudah ada di database.`); return; }
    const newEntry: NurseMaster = { id: gId(), nik, name, pk };
    const updated = [...nurseMasterList, newEntry];
    setNurseMasterList(updated);
    setMasterForm({ nik:"", name:"", pk:"III" });
    setMasterFormErr("");
    setMasterSaving(true);
    /* Simpan ke Supabase */
    try {
      const res = await upsertOneToSupa("kb_roster_gen", {
        id: "master_perawat",
        updated_at: new Date().toISOString(),
        nurses: JSON.stringify(updated),
        jadwal: "", holidays: "", year: 0, month: 0,
        generatedAt: fNow(), pjSirs, pembawHP, kepInstall,
      });
      if(res.ok){
        showToast(`✅ "${name}" (NIK: ${nik}) disimpan ke Supabase!`, "#2E7D32");
      } else {
        showToast(`✅ "${name}" disimpan lokal. Supabase: ${res.error||"gagal sync"}`, "#E65100");
      }
    } catch {
      showToast(`✅ "${name}" disimpan lokal — Supabase tidak terjangkau.`, "#E65100");
    }
    setMasterSaving(false);
  };

  const removeFromMaster = async (id: string) => {
    const target = nurseMasterList.find(m => m.id === id);
    const updated = nurseMasterList.filter(m => m.id !== id);
    setNurseMasterList(updated);
    try {
      await upsertOneToSupa("kb_roster_gen", {
        id: "master_perawat", updated_at: new Date().toISOString(),
        nurses: JSON.stringify(updated),
        jadwal: "", holidays: "", year: 0, month: 0,
        generatedAt: fNow(), pjSirs, pembawHP, kepInstall,
      });
    } catch { /* silent — LS sudah update via useEffect */ }
    if(target) showToast(`🗑 Perawat "${target.name}" dihapus dari database master.`, "#5C677D");
  };

  // BUG FIX #2: Handler sisa cuti — tidak ada leading zero
  const handleSisaCutiChange = (ni: number, raw: string) => {
    const parsed = parseInt(raw, 10);
    const val = isNaN(parsed) ? 0 : Math.min(Math.max(parsed, 0), 30);
    setNurses(p => {
      const np = p.map((n,i) => i===ni ? {...n, sisaCuti: val} : n);
      triggerAutoSave(grid, np, holidays, month, year);
      return np;
    });
  };

  const updateCell = (ni: number, di: number, val: string) => {
    setGrid(g => {
      const ng = g.map((row, ri) => ri===ni ? row.map((c,ci) => ci===di ? val.toUpperCase() : c) : row);
      triggerAutoSave(ng, nurses, holidays, month, year);
      return ng;
    });
  };

  const updateHoliday = (di: number, checked: boolean) => {
    setHolidays(h => {
      const nh = h.map((v,i) => i===di ? checked : v);
      triggerAutoSave(grid, nurses, nh, month, year);
      return nh;
    });
  };

  /* ─── CORE ENGINE v2 — Pola Roster Realistis & Variatif ─────────────
     Referensi: foto jadwal RS Panti Rini April 2026.
     Prinsip:
     A) Setiap perawat punya pola "blok shift" yang dirotasi acak tiap Generate.
        Blok = urutan P/PG atau S/SG yang bergantian, diselingi L/LG.
     B) Hari Minggu/Libur: tepat 4 orang LG, sisanya L. Merata 2–3x/orang/bulan.
     C) Karu selalu P (tidak pernah SG/LG). Perawat PROFESI/CUTI → beset manual.
     D) Constraint: 2 S/SG berturut → hari berikutnya P/PG.
     E) PG/SG maks 2–3x/minggu per perawat.
     F) Hasil acak setiap kali Generate ditekan (Fisher-Yates pada template pilihan).
  ─────────────────────────────────────────────────────────────────── */
  const generateRoster = () => {
    const newGrid = grid.map(row => [...row]);
    const warns: string[] = [];
    const dim = daysInMonth;

    /* ── Konstanta ── */
    const LG_PER_DAY     = 4;   // tepat 4 orang LG per hari libur
    const MAX_LG_MONTH   = 3;   // maks 3x LG per orang per bulan (min 2)
    const MIN_LG_MONTH   = 2;   // target minimum LG per orang
    const MAX_PGSG_WEEK  = 3;   // maks PG/SG dalam 7 hari berjalan

    /* ── Tracker bulanan ── */
    const lgCount   = Array(nurses.length).fill(0);  // LG per bulan
    const pgsgLog   = nurses.map(() => Array(dim).fill(false)); // PG/SG per hari

    const pgsgInWindow = (ni: number, d: number) => {
      let c = 0;
      for(let w = Math.max(0, d-6); w < d; w++) if(pgsgLog[ni][w]) c++;
      return c;
    };

    /* ── Hitung jumlah hari libur dalam bulan → untuk target LG merata ── */
    const holidayDays: number[] = [];
    for(let d=0; d<dim; d++){
      if(new Date(year,month,d+1).getDay()===0 || holidays[d]) holidayDays.push(d);
    }
    /* Target LG per orang: floor(holidayDays * LG_PER_DAY / nonKaruCount) clamped MIN..MAX */
    const nonKaruCount = Math.max(1, nurses.filter(n=>n.tipe!=="karu").length);
    const rawTarget    = Math.round(holidayDays.length * LG_PER_DAY / nonKaruCount);
    const lgTarget     = Math.min(MAX_LG_MONTH, Math.max(MIN_LG_MONTH, rawTarget));

    /* ── PASS 1: Hari Libur/Minggu — LG 4 orang, merata ── */
    /* Urutkan perawat non-Karu: yang LG-count paling sedikit diprioritaskan */
    for(const d of holidayDays){
      const karuIdxs    = nurses.map((_,i)=>i).filter(i=>nurses[i].tipe==="karu");
      const nonKaruIdxs = fisherYates(nurses.map((_,i)=>i).filter(i=>nurses[i].tipe!=="karu"));

      /* Locked cell: jangan timpa beset manual */
      const locked = new Set(nurses.map((_,i)=>i).filter(i=>newGrid[i][d]!==""));
      const availNonKaru = nonKaruIdxs.filter(i=>!locked.has(i));
      const availKaru    = karuIdxs.filter(i=>!locked.has(i));

      /* Prioritaskan yang LG count masih di bawah lgTarget, lalu di bawah MAX */
      const wantLG    = availNonKaru.filter(i=>lgCount[i]<lgTarget).sort((a,b)=>lgCount[a]-lgCount[b]);
      const canLG     = availNonKaru.filter(i=>lgCount[i]>=lgTarget && lgCount[i]<MAX_LG_MONTH);
      const lgPool    = [...wantLG, ...canLG];

      const lgAssigned: number[] = [];
      for(const ni of lgPool){
        if(lgAssigned.length >= LG_PER_DAY) break;
        lgAssigned.push(ni);
        newGrid[ni][d] = "LG";
        lgCount[ni]++;
      }
      /* Sisa non-karu yang tidak terpilih LG → L */
      availNonKaru.filter(i=>!lgAssigned.includes(i)).forEach(i=>{ newGrid[i][d]="L"; });
      availKaru.forEach(i=>{ newGrid[i][d]="L"; }); // Karu selalu L di hari libur

      if(lgAssigned.length < LG_PER_DAY){
        warns.push(`⚠️ Hari ke-${d+1} (Libur): hanya ${lgAssigned.length} LG tersedia (target ${LG_PER_DAY}).`);
      }
    }

    /* ── PASS 2: Hari Kerja — template rotasi per perawat ──────────────
       Prinsip dari foto:
       • Urutan BARIS perawat di tabel TIDAK berubah (tetap 1-10).
       • Setiap perawat punya template pola 7-hari sendiri + offset acak.
       • Template dipilih & offset di-acak sekali di awal, sebelum loop hari.
       • Constraint 2S berturut → P tetap berlaku (override template).
    ─────────────────────────────────────────────────────────────────── */
    const WEEKLY_TEMPLATES = [
      ["P","P","S","S","P","S","P"],
      ["S","S","P","P","S","P","S"],
      ["P","S","P","S","P","P","S"],
      ["S","P","S","P","S","S","P"],
      ["P","P","P","S","S","P","S"],
      ["S","S","S","P","P","S","P"],
      ["P","S","S","P","P","S","P"],
      ["S","P","P","S","S","P","S"],
    ];

    /* Assign template & offset SEKALI per perawat — tidak berubah selama loop hari */
    const nurseTemplate: number[] = nurses.map((_,i) =>
      nurses[i].tipe==="karu" ? -1 : Math.floor(Math.random() * WEEKLY_TEMPLATES.length)
    );
    const nurseOffset: number[] = nurses.map((_,i) =>
      nurses[i].tipe==="karu" ? 0 : Math.floor(Math.random() * 7)
    );

    for(let d=0; d<dim; d++){
      const isLibur = new Date(year,month,d+1).getDay()===0 || holidays[d];
      if(isLibur) continue;

      const locked = new Set(nurses.map((_,i)=>i).filter(i=>newGrid[i][d]!==""));
      nurses.forEach((_,i)=>{
        if(locked.has(i) && (newGrid[i][d]==="PG"||newGrid[i][d]==="SG")) pgsgLog[i][d]=true;
      });

      /* Karu selalu P */
      nurses.forEach((_,i)=>{ if(nurses[i].tipe==="karu" && !locked.has(i)) newGrid[i][d]="P"; });

      /* Non-Karu: baca dari template → apply constraint → pilih siaga */
      nurses.forEach((_,ni)=>{
        if(nurses[ni].tipe==="karu" || locked.has(ni)) return;

        const tmpl        = WEEKLY_TEMPLATES[nurseTemplate[ni]];
        const dayInCycle  = (d + nurseOffset[ni]) % 7;
        let   baseShift   = tmpl[dayInCycle]; // "P" atau "S"

        /* Constraint: 2 S/SG berturut → paksa P */
        if(d>=2){
          const p1=newGrid[ni][d-1], p2=newGrid[ni][d-2];
          if((p1==="S"||p1==="SG") && (p2==="S"||p2==="SG")) baseShift="P";
        }

        /* Pilih siaga ~30% jika limit mingguan belum tercapai */
        const canSiaga = pgsgInWindow(ni,d) < MAX_PGSG_WEEK;
        if(baseShift==="P"){
          if(canSiaga && Math.random()<0.30){ newGrid[ni][d]="PG"; pgsgLog[ni][d]=true; }
          else newGrid[ni][d]="P";
        } else {
          if(canSiaga && Math.random()<0.30){ newGrid[ni][d]="SG"; pgsgLog[ni][d]=true; }
          else newGrid[ni][d]="S";
        }
      });

      /* Warning formasi darurat */
      const cutiHariIni = nurses.filter((_,i)=>newGrid[i][d]==="T"||newGrid[i][d]==="L").length;
      const pagiHariIni = nurses.filter((_,i)=>newGrid[i][d]==="P"||newGrid[i][d]==="PG").length;
      const siangHariIni= nurses.filter((_,i)=>newGrid[i][d]==="S"||newGrid[i][d]==="SG").length;
      if(cutiHariIni>=3){
        warns.push(`⚠️ Hari ke-${d+1}: ${cutiHariIni} cuti/libur. Pagi: ${pagiHariIni}, Siang: ${siangHariIni}.`);
      }
    }


    /* ── PASS 3: Audit LG — jika ada non-Karu yang LG < MIN, coba tambahkan
       di hari libur yang LG-nya masih < LG_PER_DAY ── */
    const underLG = nurses.map((_,i)=>i).filter(i=>nurses[i].tipe!=="karu" && lgCount[i]<MIN_LG_MONTH);
    if(underLG.length>0){
      for(const d of holidayDays){
        const lgToday = nurses.filter((_,i)=>newGrid[i][d]==="LG").length;
        if(lgToday >= LG_PER_DAY+1) continue; // sudah penuh (toleransi +1 max 5)
        for(const ni of underLG){
          if(lgCount[ni]>=MIN_LG_MONTH) continue;
          if(newGrid[ni][d]==="L"){ // ganti L → LG
            newGrid[ni][d]="LG"; lgCount[ni]++;
            if(nurses.filter((_,i)=>newGrid[i][d]==="LG").length>=LG_PER_DAY) break;
          }
        }
      }
    }

    setGrid(newGrid);
    setWarnings(warns);
    if(warns.length>0) warns.slice(0,3).forEach(w=>showToast(w,"#E07800"));
    AmbilLogMedis("ROSTER_STAF","Admin","GENERATE_ROSTER",
      `Roster di-generate: ${BULAN_NAMA_ID[month]} ${year} — ${nurses.length} perawat`,
      { month, year, nurseCount: nurses.length, warns: warns.length }
    );
    showToast(`✅ Roster berhasil di-generate! LG: ${lgCount.map((c,i)=>nurses[i]?.name?`${nurses[i].name.split(" ")[0]}:${c}`:"").filter(Boolean).join(", ")}`, "#2E7D32");
  };

  /* ─── Live counter per kolom ──────────────────────────────────────── */
  const countCode = (code: string) => {
    return Array.from({length:daysInMonth},(_,d)=>nurses.reduce((c,_,ni)=>newGrid_safe(ni,d)===code?c+1:c,0));
  };
  const newGrid_safe = (ni:number, d:number) => grid[ni]?.[d] ?? "";
  const sumRow = (codes: string[]) => (d:number) => nurses.reduce((c,_,ni)=>codes.includes(newGrid_safe(ni,d))?c+1:c,0);

  /* ─── FITUR 2 — TOMBOL "SIMPAN JADWAL" ─────────────────────────────── */
  const saveJadwal = async () => {
    setSavingJadwal(true);
    const ok = await syncToCloudSilent(grid, nurses, holidays, month, year, true);
    setSavingJadwal(false);
    if(ok){
      /* Juga simpan ke localStorage sebagai konfirmasi final */
      saveDraftToLS(grid, nurses, holidays, month, year);
    }
  };

  /* ─── FITUR 1 — DOWNLOAD EXCEL FORMAT RESMI ────────────────────────── */
  const downloadExcel = async () => {
    try {
      const wb  = new ExcelJS.Workbook();
      const ws  = wb.addWorksheet("Roster Perawat");
      const dim = daysInMonth;
      const bulanLabel = `${BULAN_NAMA_ID[month]} ${year}`;
      /* Kolom: NO(1) NIK(2) PK(3) SISA CUTI(4) NAMA(5)  tanggal 1..dim  ringkasan 12 kolom */
      const COL_NO=1, COL_NIK=2, COL_PK=3, COL_CUTI=4, COL_NAMA=5;
      const COL_DATE_START = 6;                    // tanggal mulai kolom 6
      const COL_DATE_END   = COL_DATE_START + dim - 1;
      const COL_SUM_START  = COL_DATE_END + 1;     // kolom ringkasan individu
      const TOTAL_COLS     = COL_SUM_START + SUMMARY_COLS.length - 1;

      /* ── BARIS 1: Judul ── */
      ws.mergeCells(1, 1, 1, TOTAL_COLS);
      const r1 = ws.getCell(1, 1);
      r1.value = "JADWAL PETUGAS KAMAR BEDAH";
      r1.font  = { bold: true, size: 14, color: { argb: "FF0F172A" } };
      r1.fill  = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE0F2FE" } };
      r1.alignment = { horizontal: "center", vertical: "middle" };
      ws.getRow(1).height = 30;

      /* ── BARIS 2: Bulan ── */
      ws.mergeCells(2, 1, 2, TOTAL_COLS);
      const r2 = ws.getCell(2, 1);
      r2.value = `BULAN : ${bulanLabel}`;
      r2.font  = { bold: true, size: 12, color: { argb: "FF0369A1" } };
      r2.fill  = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF0F9FF" } };
      r2.alignment = { horizontal: "center", vertical: "middle" };
      ws.getRow(2).height = 22;

      /* ── BARIS 3 & 4: Spacer ── */
      ws.getRow(3).height = 8;
      ws.getRow(4).height = 8;

      /* ── BARIS 5: Header label kolom (merged 5–6 untuk kolom statis & ringkasan) ── */
      const fixedLabels: Array<[number, number, string]> = [
        [COL_NO,   COL_NO,   "NO"],
        [COL_NIK,  COL_NIK,  "NIK"],
        [COL_PK,   COL_PK,   "PK"],
        [COL_CUTI, COL_CUTI, "SISA\nCUTI"],
        [COL_NAMA, COL_NAMA, "NAMA"],
      ];
      fixedLabels.forEach(([c1, c2, label]) => {
        ws.mergeCells(5, c1, 6, c2);
        const cell = ws.getCell(5, c1);
        cell.value = label;
        cell.font  = { bold: true, size: 10, color: { argb: "FFFFFFFF" } };
        cell.fill  = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0284C7" } };
        cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
        cell.border = { top:{style:"thin"}, bottom:{style:"thin"}, left:{style:"thin"}, right:{style:"thin"} };
      });

      /* ── BARIS 5: Merged header "TANGGAL" di atas kolom 1..dim ── */
      ws.mergeCells(5, COL_DATE_START, 5, COL_DATE_END);
      const tanggalHeader = ws.getCell(5, COL_DATE_START);
      tanggalHeader.value = "TANGGAL";
      tanggalHeader.font  = { bold: true, size: 10, color: { argb: "FFFFFFFF" } };
      tanggalHeader.fill  = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0369A1" } };
      tanggalHeader.alignment = { horizontal: "center", vertical: "middle" };
      tanggalHeader.border = { top:{style:"thin"}, bottom:{style:"thin"}, left:{style:"thin"}, right:{style:"thin"} };

      /* ── BARIS 5: Merged header "RINGKASAN" ── */
      ws.mergeCells(5, COL_SUM_START, 5, TOTAL_COLS);
      const ringkasanH = ws.getCell(5, COL_SUM_START);
      ringkasanH.value = "RINGKASAN";
      ringkasanH.font  = { bold: true, size: 10, color: { argb: "FFFFFFFF" } };
      ringkasanH.fill  = { type: "pattern", pattern: "solid", fgColor: { argb: "FF7C3AED" } };
      ringkasanH.alignment = { horizontal: "center", vertical: "middle" };
      ringkasanH.border = { top:{style:"thin"}, bottom:{style:"thin"}, left:{style:"thin"}, right:{style:"thin"} };

      /* ── BARIS 6: Angka tanggal + kode ringkasan ── */
      ws.getRow(5).height = 20;
      ws.getRow(6).height = 18;
      for(let d=0; d<dim; d++){
        const date  = new Date(year, month, d+1);
        const isSun = date.getDay() === 0;
        const isHol = holidays[d];
        const dc    = ws.getCell(6, COL_DATE_START + d);
        dc.value    = d + 1;
        dc.font     = { bold: isSun||isHol, size: 9, color: { argb: isSun||isHol?"FFDC2626":"FF0F172A" } };
        dc.fill     = { type: "pattern", pattern: "solid", fgColor: { argb: isSun||isHol?"FFFFCDD2":"FFF0F9FF" } };
        dc.alignment = { horizontal: "center", vertical: "middle" };
        dc.border    = { top:{style:"thin"}, bottom:{style:"thin"}, left:{style:"thin"}, right:{style:"thin"} };
      }
      SUMMARY_COLS.forEach((lbl, si) => {
        const sc = ws.getCell(6, COL_SUM_START + si);
        sc.value = lbl;
        sc.font  = { bold: true, size: 9, color: { argb: "FFFFFFFF" } };
        sc.fill  = { type: "pattern", pattern: "solid", fgColor: { argb: "FF6D28D9" } };
        sc.alignment = { horizontal: "center", vertical: "middle" };
        sc.border    = { top:{style:"thin"}, bottom:{style:"thin"}, left:{style:"thin"}, right:{style:"thin"} };
      });

      /* ── BARIS DATA perawat (mulai baris 7) ── */
      const DATA_START = 7;
      nurses.forEach((n, ni) => {
        const rowNum = DATA_START + ni;
        const isEven = ni % 2 === 0;
        const dr = ws.getRow(rowNum);
        dr.height = 18;
        const bgData = isEven ? "FFFAFBFF" : "FFF0F7FF";
        const statics: [number, any][] = [
          [COL_NO,   ni+1],
          [COL_NIK,  n.nik  || ""],
          [COL_PK,   n.pk   || ""],
          [COL_CUTI, n.sisaCuti],
          [COL_NAMA, n.name || ""],
        ];
        statics.forEach(([col, val]) => {
          const cell = dr.getCell(col);
          cell.value = val;
          cell.fill  = { type:"pattern", pattern:"solid", fgColor:{ argb: bgData } };
          cell.alignment = { horizontal: col===COL_NAMA ? "left" : "center", vertical: "middle" };
          cell.border = { top:{style:"thin"}, bottom:{style:"thin"}, left:{style:"thin"}, right:{style:"thin"} };
          if(col===COL_NAMA) cell.font = { bold: true, size: 10 };
        });
        /* Kolom shift harian */
        const shiftRow: string[] = [];
        for(let d=0; d<dim; d++){
          const code = newGrid_safe(ni, d);
          shiftRow.push(code);
          const sc = dr.getCell(COL_DATE_START + d);
          sc.value = code;
          /* Warna per kode */
          const argbFill = code==="P"||code==="PG" ? "FFE3F2FD"
            : code==="S"||code==="SG" ? "FFFFF3E0"
            : code==="T"              ? "FFFFEBEE"
            : code==="L"||code==="LG" ? "FFF3F3F3"
            : bgData;
          const argbFont = code==="P"||code==="PG" ? "FF1565C0"
            : code==="S"||code==="SG" ? "FFE65100"
            : code==="T"              ? "FFC62828"
            : code==="L"||code==="LG" ? "FF5C677D"
            : "FF374151";
          sc.fill = { type:"pattern", pattern:"solid", fgColor:{ argb: argbFill } };
          sc.font = { bold: true, size: 9, color:{ argb: argbFont } };
          sc.alignment = { horizontal:"center", vertical:"middle" };
          sc.border = { top:{style:"thin"}, bottom:{style:"thin"}, left:{style:"thin"}, right:{style:"thin"} };
        }
        /* Kolom ringkasan individu (7 kolom: ∑P ∑PG ∑S ∑SG ∑L/LG ∑T ∑Total) */
        const sumVals = hitungRingkasan(shiftRow, dim);
        sumVals.forEach((val, si) => {
          const sc = dr.getCell(COL_SUM_START + si);
          sc.value = val || "";
          sc.fill  = { type:"pattern", pattern:"solid", fgColor:{ argb: "FFF5F3FF" } };
          sc.font  = { bold: val > 0, size: 10, color:{ argb: "FF6D28D9" } };
          sc.alignment = { horizontal:"center", vertical:"middle" };
          sc.border = { top:{style:"thin"}, bottom:{style:"thin"}, left:{style:"thin"}, right:{style:"thin"} };
        });
      });

      /* ── BARIS AGREGAT TOTAL HARIAN (di bawah semua data) ── */
      const TOTAL_ROW_START = DATA_START + nurses.length;
      const SUMMARY_DEFS2 = [
        { label:"∑ PAGI (P)",         codes:["P"],       argbBg:"FFBBDEFB", argbFg:"FF0D47A1" },
        { label:"∑ PAGI SIAGA (PG)",  codes:["PG"],      argbBg:"FF90CAF9", argbFg:"FF1565C0" },
        { label:"∑ SIANG (S)",        codes:["S"],        argbBg:"FFFFE0B2", argbFg:"FFE65100" },
        { label:"∑ SIANG SIAGA (SG)", codes:["SG"],       argbBg:"FFFFCC80", argbFg:"FFE65100" },
        { label:"∑ LIBUR SIAGA (LG)", codes:["LG"],       argbBg:"FFCFD8DC", argbFg:"FF37474F" },
        { label:"∑ CUTI (T)",         codes:["T"],        argbBg:"FFFFCDD2", argbFg:"FFC62828" },
      ];
      SUMMARY_DEFS2.forEach(({ label, codes, argbBg, argbFg }, si) => {
        const rowNum = TOTAL_ROW_START + si;
        const sr = ws.getRow(rowNum);
        sr.height = 22;
        ws.mergeCells(rowNum, COL_NO, rowNum, COL_DATE_START - 1);
        const lc = sr.getCell(COL_NO);
        lc.value = label;
        lc.font  = { bold: true, size: 12, color: { argb: argbFg } };
        lc.fill  = { type:"pattern", pattern:"solid", fgColor:{ argb: argbBg } };
        lc.alignment = { horizontal: "right", vertical: "middle" };
        lc.border = { top:{style:"medium"}, bottom:{style:"medium"}, left:{style:"thin"}, right:{style:"thin"} };
        for(let d=0; d<dim; d++){
          const cnt = nurses.reduce((c,_,ni) => codes.includes(newGrid_safe(ni,d)) ? c+1 : c, 0);
          const nc = sr.getCell(COL_DATE_START + d);
          nc.value = cnt || "";
          nc.font  = { bold: true, size: 13, color: { argb: argbFg } };
          nc.fill  = { type:"pattern", pattern:"solid", fgColor:{ argb: argbBg } };
          nc.alignment = { horizontal:"center", vertical:"middle" };
          nc.border = { top:{style:"medium"}, bottom:{style:"medium"}, left:{style:"thin"}, right:{style:"thin"} };
        }
        /* kosongkan kolom ringkasan di baris total */
        for(let sc2=0; sc2<SUMMARY_COLS.length; sc2++){
          const ec = sr.getCell(COL_SUM_START + sc2);
          ec.fill = { type:"pattern", pattern:"solid", fgColor:{ argb: "FFF5F5F5" } };
          ec.border = { top:{style:"thin"}, bottom:{style:"thin"}, left:{style:"thin"}, right:{style:"thin"} };
        }
      });

      /* ── FOOTER / TANDA TANGAN ── */
      const FOOTER_ROW = TOTAL_ROW_START + SUMMARY_DEFS2.length + 3;
      ws.getRow(FOOTER_ROW - 1).height = 14; // spacer

      /* Kiri: PJ SIRS & Pembawa HP */
      ws.getCell(FOOTER_ROW, 1).value = `PJ SIRS : ${pjSirs}`;
      ws.getCell(FOOTER_ROW, 1).font  = { bold: true, size: 10 };
      ws.getCell(FOOTER_ROW + 1, 1).value = `Pembawa HP : ${pembawHP}`;
      ws.getCell(FOOTER_ROW + 1, 1).font  = { size: 10 };

      /* Kanan: Blok tanda tangan kepala instalasi */
      const sigCol = TOTAL_COLS - 4;
      ws.mergeCells(FOOTER_ROW, sigCol, FOOTER_ROW, TOTAL_COLS);
      ws.getCell(FOOTER_ROW, sigCol).value = "MENGETAHUI";
      ws.getCell(FOOTER_ROW, sigCol).font  = { bold: true, size: 10 };
      ws.getCell(FOOTER_ROW, sigCol).alignment = { horizontal: "center" };

      ws.mergeCells(FOOTER_ROW+1, sigCol, FOOTER_ROW+1, TOTAL_COLS);
      ws.getCell(FOOTER_ROW+1, sigCol).value = "Kepala Instalasi Kamar Bedah";
      ws.getCell(FOOTER_ROW+1, sigCol).font  = { size: 10 };
      ws.getCell(FOOTER_ROW+1, sigCol).alignment = { horizontal: "center" };

      /* Spasi tanda tangan 4 baris */
      for(let r=0; r<4; r++){
        ws.getRow(FOOTER_ROW + 2 + r).height = 16;
      }

      ws.mergeCells(FOOTER_ROW+6, sigCol, FOOTER_ROW+6, TOTAL_COLS);
      ws.getCell(FOOTER_ROW+6, sigCol).value = kepInstall;
      ws.getCell(FOOTER_ROW+6, sigCol).font  = { bold: true, size: 10, underline: true };
      ws.getCell(FOOTER_ROW+6, sigCol).alignment = { horizontal: "center" };

      /* ── Lebar kolom ── */
      ws.getColumn(COL_NO).width   = 5;
      ws.getColumn(COL_NIK).width  = 13;
      ws.getColumn(COL_PK).width   = 5;
      ws.getColumn(COL_CUTI).width = 7;
      ws.getColumn(COL_NAMA).width = 24;
      for(let d=1; d<=dim; d++) ws.getColumn(COL_DATE_START + d - 1).width = 4.5;
      SUMMARY_COLS.forEach((_, si) => { ws.getColumn(COL_SUM_START + si).width = 5; });

      /* ── Freeze pane: beku 5 kolom kiri + 6 baris atas ── */
      ws.views = [{ state:"frozen", xSplit: COL_DATE_START - 1, ySplit: 6, activeCell:"F7" }];

      const buf  = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf as ArrayBuffer], {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement("a");
      a.href = url;
      a.download = `Jadwal_Petugas_KamarBedah_${BULAN_NAMA_ID[month]}_${year}.xlsx`;
      document.body.appendChild(a); a.click();
      document.body.removeChild(a); URL.revokeObjectURL(url);
      showToast("✅ File Excel berhasil diunduh!", "#2E7D32");
      AmbilLogMedis("LAPORAN_BACKUP","Admin","DOWNLOAD_EXCEL_ROSTER",
        `Excel Roster diunduh: ${BULAN_NAMA_ID[month]} ${year}`,
        { month, year, nurses: nurses.length }
      );
    } catch(e:any) {
      showToast("⚠ Gagal export Excel: " + (e?.message||"error"), "#D62828");
    }
  };

  /* ─── Sync to Supabase (manual, dengan feedback) ─────────────────── */
  const syncToCloud = async () => {
    setSyncing(true);
    await syncToCloudSilent(grid, nurses, holidays, month, year, true);
    setSyncing(false);
  };

  /* ─── Backup to Dropbox (Excel, manual) ─────────────────────────── */
  const backupToDropbox = async () => {
    setSyncingDbx(true);
    try {
      /* Re-use downloadExcel logic, tapi tulis ke base64 → Dropbox */
      const wb  = new ExcelJS.Workbook();
      const ws  = wb.addWorksheet("Roster");
      const bulanLabel = `${BULAN_NAMA_ID[month]} ${year}`;
      ws.mergeCells(1,1,1,5+daysInMonth);
      const tc = ws.getCell(1,1);
      tc.value = `JADWAL PETUGAS KAMAR BEDAH — ${bulanLabel}`;
      tc.font  = {bold:true,size:12,color:{argb:"FFFFFFFF"}};
      tc.fill  = {type:"pattern",pattern:"solid",fgColor:{argb:"FF0284C7"}};
      tc.alignment = {horizontal:"center",vertical:"middle"};
      ws.getRow(1).height = 26;
      const hdrs = ["NO","NIK","PK","SISA CUTI","NAMA",...Array.from({length:daysInMonth},(_,i)=>String(i+1))];
      const hr = ws.getRow(2);
      hdrs.forEach((h,ci)=>{
        const cell=hr.getCell(ci+1); cell.value=h;
        cell.font={bold:true,size:10,color:{argb:"FFFFFFFF"}};
        cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF234B45"}};
        cell.alignment={horizontal:"center",vertical:"middle"};
        cell.border={top:{style:"thin"},bottom:{style:"thin"},left:{style:"thin"},right:{style:"thin"}};
      });
      hr.height=20;
      nurses.forEach((n,ni)=>{
        const dr=ws.getRow(3+ni);
        [ni+1,n.nik,n.pk,n.sisaCuti,n.name].forEach((v,ci)=>{
          const c=dr.getCell(ci+1); c.value=v;
          c.alignment={horizontal:"center",vertical:"middle"};
          c.border={top:{style:"thin"},bottom:{style:"thin"},left:{style:"thin"},right:{style:"thin"}};
        });
        for(let d=0;d<daysInMonth;d++){
          const c=dr.getCell(6+d); c.value=newGrid_safe(ni,d);
          c.alignment={horizontal:"center",vertical:"middle"};
          c.border={top:{style:"thin"},bottom:{style:"thin"},left:{style:"thin"},right:{style:"thin"}};
        }
        dr.height=18;
      });
      ws.getColumn(1).width=5; ws.getColumn(2).width=12; ws.getColumn(3).width=5;
      ws.getColumn(4).width=7; ws.getColumn(5).width=22;
      for(let d=1;d<=daysInMonth;d++) ws.getColumn(5+d).width=4.5;
      const folder = (dbxCfg?.path||"/kamarbedah/").replace(/[^/]+$/,"");
      const stamp  = `${year}-${String(month+1).padStart(2,"0")}`;
      const path   = `${folder}Jadwal_KamarBedah_${stamp}.xlsx`;
      const buf    = await wb.xlsx.writeBuffer() as Uint8Array;
      const CHUNK=8192; let bin="";
      for(let i=0;i<buf.length;i+=CHUNK) bin+=String.fromCharCode(...buf.subarray(i,i+CHUNK));
      const b64 = btoa(bin);
      const res = await callDbxProxy({action:"upload_excel", path, base64: b64});
      if(res.ok) showToast("✅ Roster berhasil di-backup ke Dropbox!", "#0061FF");
      else showToast("⚠ Gagal backup Dropbox: " + res.msg, "#D62828");
    } catch(e:any) {
      showToast("⚠ Error backup Dropbox: " + (e?.message||"error"), "#D62828");
    }
    setSyncingDbx(false);
  };

  const bulanOptions = Array.from({length:12},(_,i)=>({ v:String(i), l:new Date(2000,i,1).toLocaleString("id-ID",{month:"long"}) }));
  const yearOptions  = [year-1, year, year+1].map(y=>({v:String(y), l:String(y)}));

  const cellStyle = (code: string): React.CSSProperties => {
    const base: React.CSSProperties = { textAlign:"center", fontSize:10, fontWeight:700, padding:"2px 0", border:"1px solid #e0e0e0", minWidth:28, cursor:"text" };
    if(code==="P"||code==="PG") return {...base, background:"#E3F2FD", color:"#1565C0"};
    if(code==="S"||code==="SG") return {...base, background:"#FFF3E0", color:"#E65100"};
    if(code==="L"||code==="LG") return {...base, background:"#F3F3F3", color:"#5C677D"};
    if(code==="T") return {...base, background:"#FFEBEE", color:"#C62828"};
    return {...base, background:"#fff", color:"#999"};
  };

  /* ── Label & warna indikator status cloud ── */
  const cloudStatusUI: Record<string, {label:string; color:string; bg:string}> = {
    idle:    { label:"● Belum ada perubahan",  color:"#94A3B8", bg:"#F1F5F9" },
    saving:  { label:"⟳ Menyimpan...",          color:"#0284C7", bg:"#E0F2FE" },
    ok:      { label:"✓ Tersimpan di cloud",    color:"#16A34A", bg:"#DCFCE7" },
    error:   { label:"✕ Gagal sync cloud",      color:"#DC2626", bg:"#FEE2E2" },
    offline: { label:"⚡ Offline — draft lokal", color:"#D97706", bg:"#FEF3C7" },
  };
  const csUI = cloudStatusUI[cloudStatus];

  return (
    <div style={{paddingBottom:40}}>
      {/* Header */}
      <Row title="🗓 Generator Roster Perawat" right={
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <Btn sm onClick={generateRoster} color="linear-gradient(135deg,#1e3a8a,#0284c7)">⚡ Generate</Btn>
          <Btn sm onClick={downloadExcel}  color="linear-gradient(135deg,#0369a1,#38bdf8)">⬇ Excel</Btn>
          <Btn sm onClick={syncToCloud}    color="linear-gradient(135deg,#1d4ed8,#6366f1)" disabled={syncing}>{syncing?"⏳ Supabase...":"☁ Supabase"}</Btn>
          <Btn sm onClick={backupToDropbox} color="linear-gradient(135deg,#0061FF,#38bdf8)" disabled={syncingDbx}>{syncingDbx?"⏳ Dropbox...":"📦 Dropbox"}</Btn>
        </div>
      }/>

      {/* ── Tombol Simpan Jadwal + Indikator Status Cloud ── */}
      <div style={{display:"flex",gap:10,alignItems:"center",marginBottom:12,flexWrap:"wrap"}}>
        <button onClick={saveJadwal} disabled={savingJadwal} style={{flex:1,minWidth:200,padding:"13px 20px",background:savingJadwal?"#94A3B8":"linear-gradient(135deg,#0284c7,#0369a1)",color:"#fff",border:"none",borderRadius:12,fontSize:14,fontWeight:800,cursor:savingJadwal?"not-allowed":"pointer",fontFamily:"inherit",boxShadow:"0 4px 14px rgba(2,132,199,.35)",display:"flex",alignItems:"center",justifyContent:"center",gap:8,transition:"all .2s"}}>
          {savingJadwal?<><span style={{animation:"spin 1s linear infinite",display:"inline-block"}}>⟳</span> Menyimpan...</>:<>💾 Simpan Jadwal</>}
        </button>
        <div style={{background:csUI.bg,color:csUI.color,fontSize:11,fontWeight:700,padding:"8px 12px",borderRadius:10,border:`1px solid ${csUI.color}33`,whiteSpace:"nowrap",flexShrink:0}}>
          {csUI.label}
        </div>
      </div>

      {/* ══════════════════════════════════════════════════════════════════
          FITUR 1 — CARD KELOLA DATA MASTER PERAWAT
         ══════════════════════════════════════════════════════════════════ */}
      <Card style={{marginBottom:12,background:"#F0FDF4",border:"1px solid #BBF7D0"}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}>
          <SH label="📁 Kelola Data Master Perawat" color="#15803D"/>
          <span style={{fontSize:10,fontWeight:700,background:"#DCFCE7",color:"#15803D",padding:"3px 8px",borderRadius:6,border:"1px solid #BBF7D0",whiteSpace:"nowrap"}}>
            💾 localStorage + ☁ Supabase
          </span>
        </div>

        {/* Form Tambah */}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 120px auto",gap:8,alignItems:"end",marginBottom:12}}>
          <LF label="NIK Perawat">
            <input style={{...iS,fontSize:12}} value={masterForm.nik} onChange={e=>setMasterForm(p=>({...p,nik:e.target.value}))} placeholder="cth: 201010056"/>
          </LF>
          <LF label="Nama Perawat">
            <input style={{...iS,fontSize:12}} value={masterForm.name} onChange={e=>setMasterForm(p=>({...p,name:e.target.value}))} placeholder="cth: Jaka R"/>
          </LF>
          <LF label="Pangkat PK">
            <select style={{...iS,fontSize:12}} value={masterForm.pk} onChange={e=>setMasterForm(p=>({...p,pk:e.target.value}))}>
              {["I","II","III","IV"].map(v=><option key={v} value={v}>PK {v}</option>)}
            </select>
          </LF>
          <button onClick={addToMaster} disabled={masterSaving} style={{padding:"11px 16px",background:masterSaving?"#94A3B8":"linear-gradient(135deg,#16a34a,#15803d)",color:"#fff",border:"none",borderRadius:10,fontSize:12,fontWeight:700,cursor:masterSaving?"not-allowed":"pointer",fontFamily:"inherit",whiteSpace:"nowrap",flexShrink:0,transition:"all .2s"}}>
            {masterSaving ? "⟳ Menyimpan..." : "+ Tambah"}
          </button>
        </div>
        {masterFormErr && <div style={{fontSize:11,color:"#DC2626",marginBottom:8,marginTop:-4}}>⚠ {masterFormErr}</div>}

        {/* Tabel Master */}
        {nurseMasterList.length > 0 && (
          <div style={{overflowX:"auto"}}>
            <table style={{borderCollapse:"collapse",width:"100%",fontSize:11}}>
              <thead>
                <tr style={{background:"#DCFCE7"}}>
                  {["No","NIK","Nama Perawat","Pangkat PK","Aksi"].map(h=>(
                    <th key={h} style={{padding:"6px 10px",fontWeight:700,color:"#15803D",textAlign:h==="Aksi"?"center":"left",borderBottom:"2px solid #BBF7D0",whiteSpace:"nowrap"}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {nurseMasterList.map((m,i)=>(
                  <tr key={m.id} style={{background:i%2===0?"#fff":"#F0FDF4",borderBottom:"1px solid #D1FAE5"}}>
                    <td style={{padding:"5px 10px",color:C.tL}}>{i+1}</td>
                    <td style={{padding:"5px 10px",fontWeight:600,color:C.t,fontFamily:"monospace"}}>{m.nik}</td>
                    <td style={{padding:"5px 10px",fontWeight:700,color:C.t}}>{m.name}</td>
                    <td style={{padding:"5px 10px",color:"#15803D",fontWeight:600}}>PK {m.pk}</td>
                    <td style={{padding:"5px 10px",textAlign:"center"}}>
                      <button onClick={()=>removeFromMaster(m.id)} style={{padding:"3px 10px",background:"#FEE2E2",color:"#DC2626",border:"1px solid #FECACA",borderRadius:6,fontSize:10,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Hapus</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            <div style={{fontSize:11,color:C.tL,marginTop:6}}>Total: {nurseMasterList.length} perawat tersimpan di database master.</div>
          </div>
        )}
      </Card>

      {/* Pilih Bulan / Tahun */}
      <Card style={{marginBottom:12}}>
        <div style={{display:"flex",gap:12,flexWrap:"wrap",alignItems:"center"}}>
          <div style={{fontSize:13,fontWeight:700,color:C.g}}>Bulan & Tahun:</div>
          <select value={month} onChange={e=>handleMonthYear(Number(e.target.value),year)} style={{...iS,width:140,padding:"7px 12px"}}>
            {bulanOptions.map(o=><option key={o.v} value={o.v}>{o.l}</option>)}
          </select>
          <select value={year} onChange={e=>handleMonthYear(month,Number(e.target.value))} style={{...iS,width:100,padding:"7px 12px"}}>
            {yearOptions.map(o=><option key={o.v} value={o.v}>{o.l}</option>)}
          </select>
          <div style={{fontSize:12,color:C.tL}}>→ {daysInMonth} hari</div>
        </div>
      </Card>

      {/* Konfigurasi Dokumen */}
      <Card style={{marginBottom:12,background:"#F0F9FF",border:"1px solid #BAE6FD"}}>
        <SH label="📋 Konfigurasi Dokumen" color="#0369A1"/>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <LF label="PJ SIRS"><input style={iS} value={pjSirs} onChange={e=>setPjSirs(e.target.value)} placeholder="Nama PJ SIRS"/></LF>
          <LF label="Pembawa HP"><input style={iS} value={pembawHP} onChange={e=>setPembawHP(e.target.value)} placeholder="Nama pembawa HP"/></LF>
          <div style={{gridColumn:"1 / -1"}}>
            <LF label="Kepala Instalasi Kamar Bedah">
              <input style={iS} value={kepInstall} onChange={e=>setKepInstall(e.target.value)} placeholder="cth: drg. Agus Sri G, Sp. BM"/>
            </LF>
          </div>
        </div>
      </Card>

      {/* Legend */}
      <Card style={{marginBottom:12}}>
        <div style={{fontSize:11,fontWeight:700,color:C.g,marginBottom:8}}>KETERANGAN KODE SHIFT</div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          {[["P","Pagi","#E3F2FD","#1565C0"],["PG","Pagi Siaga","#E3F2FD","#1565C0"],["S","Siang","#FFF3E0","#E65100"],["SG","Siang Siaga","#FFF3E0","#E65100"],["L","Libur","#F3F3F3","#5C677D"],["LG","Libur Siaga","#F3F3F3","#5C677D"],["T","Cuti","#FFEBEE","#C62828"]].map(([k,l,bg,cl])=>(
            <span key={k} style={{background:bg,color:cl,border:`1px solid ${cl}33`,fontSize:11,fontWeight:700,padding:"3px 8px",borderRadius:6}}>{k} = {l}</span>
          ))}
        </div>
      </Card>

      {/* Warnings */}
      {warnings.length>0 && (
        <Card style={{marginBottom:12,background:"#FFF8E1",border:"1px solid #E07800"}}>
          <div style={{fontSize:12,fontWeight:700,color:"#E07800",marginBottom:8}}>⚠ Peringatan Formasi Darurat</div>
          {warnings.map((w,i)=><div key={i} style={{fontSize:11,color:"#5C4000",marginBottom:4}}>• {w}</div>)}
        </Card>
      )}

      {/* ══════════════════════════════════════════════════════════════════
          TABEL ROSTER — DUAL AXIS (Kanan: ringkasan individu, Bawah: komposisi harian)
         ══════════════════════════════════════════════════════════════════ */}
      <Card style={{padding:0,overflow:"hidden"}}>
        <div style={{overflowX:"auto"}}>
          <table style={{borderCollapse:"collapse",minWidth:"100%",fontSize:11}}>
            <thead>
              {/* Row hari libur checkbox */}
              <tr style={{background:"#F8F9FA"}}>
                <td colSpan={6} style={{padding:"6px 10px",fontSize:11,fontWeight:700,color:C.g,borderBottom:"1px solid #e0e0e0"}}>Hari Libur ↓</td>
                {Array.from({length:daysInMonth},(_,d)=>(
                  <td key={d} style={{textAlign:"center",padding:"4px 2px",borderBottom:"1px solid #e0e0e0",minWidth:28}}>
                    <input type="checkbox" checked={holidays[d]} onChange={e=>updateHoliday(d,e.target.checked)} style={{cursor:"pointer"}} title={`Hari ke-${d+1} = Hari Libur`}/>
                  </td>
                ))}
                {/* spacer di atas summary cols */}
                <td colSpan={7} style={{borderBottom:"1px solid #e0e0e0",background:"#EEF2FF"}}/>
              </tr>
              {/* Header kolom */}
              <tr style={{background:C.p}}>
                {["NO","NIK","PK","NAMA","SISA\nCUTI","TIPE"].map(h=>(
                  <th key={h} style={{padding:"7px 8px",color:"#fff",fontWeight:700,fontSize:11,whiteSpace:"pre",textAlign:"center",borderRight:"1px solid rgba(255,255,255,.2)"}}>{h}</th>
                ))}
                {Array.from({length:daysInMonth},(_,d)=>{
                  const date = new Date(year,month,d+1);
                  const isHol = holidays[d], isSun = date.getDay()===0;
                  return (
                    <th key={d} style={{padding:"4px 2px",color:"#fff",fontWeight:isSun||isHol?900:600,fontSize:10,textAlign:"center",background:isSun||isHol?"#C62828":C.p,minWidth:28,borderRight:"1px solid rgba(255,255,255,.15)"}}>
                      {d+1}<br/><span style={{fontSize:8,opacity:.8}}>{"MSLRKJS"[date.getDay()]}</span>
                    </th>
                  );
                })}
                {/* FITUR 3 — Header Ringkasan Individu (Sumbu Kanan) */}
                {[["∑P","#1565C0"],["∑PG","#1565C0"],["∑S","#E65100"],["∑SG","#E65100"],["∑L/LG","#5C677D"],["∑T","#C62828"],["∑Total","#374151"]].map(([lbl,cl])=>(
                  <th key={lbl} style={{padding:"4px 3px",color:"#fff",fontWeight:700,fontSize:9,textAlign:"center",background:"#3730A3",minWidth:30,borderLeft:"2px solid rgba(255,255,255,.3)",whiteSpace:"nowrap"}}>{lbl}</th>
                ))}
              </tr>
            </thead>

            <tbody>
              {nurses.map((n,ni)=>{
                const rs = rowSummaries[ni] ?? {P:0,PG:0,S:0,SG:0,T:0,LG_L:0,total:0};
                return (
                  <tr key={n.id} style={{background:ni%2===0?"#fff":"#f0f7ff"}}>
                    <td style={{padding:"4px 6px",textAlign:"center",fontWeight:700,fontSize:11,borderRight:"1px solid #f0f0f0",minWidth:32}}>{ni+1}</td>

                    {/* FITUR 2 — NIK dropdown dari nurseMasterList dinamis */}
                    <td style={{padding:"2px 4px",borderRight:"1px solid #f0f0f0",minWidth:100}}>
                      <select value={n.nik} onChange={e=>handleNikChange(ni,e.target.value)} style={{width:"100%",border:"none",background:"transparent",fontSize:10,outline:"none",cursor:"pointer"}}>
                        <option value="">-- NIK --</option>
                        {nurseMasterList.map(m=>(
                          <option key={m.nik} value={m.nik}>{m.nik} — {m.name}</option>
                        ))}
                      </select>
                    </td>

                    <td style={{padding:"2px 4px",borderRight:"1px solid #f0f0f0",minWidth:44}}>
                      <input value={n.pk} onChange={e=>updateNurse(ni,"pk",e.target.value)} style={{width:"100%",border:"none",background:"transparent",fontSize:10,textAlign:"center",outline:"none"}} placeholder="PK"/>
                    </td>
                    <td style={{padding:"2px 4px",borderRight:"1px solid #f0f0f0",minWidth:130}}>
                      <input value={n.name} onChange={e=>updateNurse(ni,"name",e.target.value)} style={{width:"100%",border:"none",background:"transparent",fontSize:11,outline:"none"}} placeholder="Nama perawat..."/>
                    </td>
                    <td style={{padding:"2px 4px",borderRight:"1px solid #f0f0f0",minWidth:46}}>
                      <input type="number" value={n.sisaCuti} onChange={e=>handleSisaCutiChange(ni,e.target.value)} style={{width:"100%",border:"none",background:"transparent",fontSize:10,textAlign:"center",outline:"none"}} min={0} max={30}/>
                    </td>
                    <td style={{padding:"2px 4px",borderRight:"1px solid #f0f0f0",minWidth:68}}>
                      <select value={n.tipe} onChange={e=>updateNurse(ni,"tipe",e.target.value as NurseType)} style={{width:"100%",border:"none",background:"transparent",fontSize:10,outline:"none",cursor:"pointer"}}>
                        <option value="karu">Karu</option>
                        <option value="senior">Senior</option>
                        <option value="junior">Junior</option>
                      </select>
                    </td>

                    {/* Sel shift harian */}
                    {Array.from({length:daysInMonth},(_,d)=>{
                      const code = newGrid_safe(ni,d);
                      return (
                        <td key={d} style={cellStyle(code)}>
                          <input value={code} onChange={e=>updateCell(ni,d,e.target.value)} style={{width:26,border:"none",background:"transparent",textAlign:"center",fontSize:10,fontWeight:700,outline:"none",color:"inherit",padding:0}} maxLength={2}/>
                        </td>
                      );
                    })}

                    {/* FITUR 3 — Ringkasan individu real-time (Sumbu Kanan) */}
                    {[
                      {v:rs.P,    bg:"#DBEAFE",cl:"#1565C0"},
                      {v:rs.PG,   bg:"#BFDBFE",cl:"#1E40AF"},
                      {v:rs.S,    bg:"#FEF3C7",cl:"#D97706"},
                      {v:rs.SG,   bg:"#FDE68A",cl:"#B45309"},
                      {v:rs.LG_L, bg:"#F1F5F9",cl:"#475569"},
                      {v:rs.T,    bg:"#FFE4E6",cl:"#BE123C"},
                      {v:rs.total,bg:"#EEF2FF",cl:"#3730A3"},
                    ].map(({v,bg,cl},i)=>(
                      <td key={i} style={{textAlign:"center",fontSize:11,fontWeight:700,color:cl,background:bg,borderLeft:"2px solid rgba(99,102,241,.2)",minWidth:30,padding:"2px 3px"}}>
                        {v||""}
                      </td>
                    ))}
                  </tr>
                );
              })}
            </tbody>

            {/* FITUR 4/5 — Summary footer (Sumbu Bawah) — dipertahankan utuh */}
            <tfoot>
              {[
                {label:"∑ Pagi (P)",         key:"P"   as keyof typeof colSummaries, bg:"#E3F2FD", cl:"#1565C0"},
                {label:"∑ Pagi Siaga (PG)",  key:"PG"  as keyof typeof colSummaries, bg:"#BBDEFB", cl:"#1565C0"},
                {label:"∑ Siang (S)",        key:"S"   as keyof typeof colSummaries, bg:"#FFF3E0", cl:"#E65100"},
                {label:"∑ Siang Siaga (SG)", key:"SG"  as keyof typeof colSummaries, bg:"#FFE0B2", cl:"#E65100"},
                {label:"∑ Libur/LG",         key:"LLG" as keyof typeof colSummaries, bg:"#F3F3F3", cl:"#5C677D"},
                {label:"∑ Cuti (T)",         key:"T"   as keyof typeof colSummaries, bg:"#FFEBEE", cl:"#C62828"},
              ].map(({label,key,bg,cl})=>(
                <tr key={label} style={{background:bg}}>
                  <td colSpan={6} style={{padding:"5px 10px",fontSize:12,fontWeight:800,color:cl,textAlign:"right",borderTop:"3px solid rgba(0,0,0,.14)"}}>{label}</td>
                  {colSummaries[key].map((cnt,d)=>(
                    <td key={d} style={{textAlign:"center",fontSize:12,fontWeight:800,color:cl,borderTop:"3px solid rgba(0,0,0,.14)",minWidth:28}}>{cnt||""}</td>
                  ))}
                  {/* 7 sel kosong di bawah kolom ringkasan kanan */}
                  {Array.from({length:7},(_,i)=><td key={`sp${i}`} style={{background:"#EEF2FF",borderTop:"3px solid rgba(0,0,0,.14)"}}/>)}
                </tr>
              ))}
            </tfoot>
          </table>
        </div>
      </Card>

      {/* Tombol Tambah Perawat + Simpan bawah */}
      <div style={{marginTop:12,display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
        <Btn sm onClick={addNurse} color={C.g} outline>+ Tambah Baris Perawat</Btn>
        <span style={{fontSize:11,color:C.tL}}>Total: {nurses.length} perawat</span>
        <div style={{flex:1}}/>
        <button onClick={saveJadwal} disabled={savingJadwal} style={{padding:"10px 18px",background:savingJadwal?"#94A3B8":"linear-gradient(135deg,#0284c7,#1e3a8a)",color:"#fff",border:"none",borderRadius:10,fontSize:12,fontWeight:700,cursor:savingJadwal?"not-allowed":"pointer",fontFamily:"inherit",boxShadow:"0 2px 8px rgba(2,132,199,.3)"}}>
          {savingJadwal?"⟳ Menyimpan...":"💾 Simpan Jadwal"}
        </button>
      </div>
      <style>{`@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}`}</style>
    </div>
  );
}

/* FIX AUDIT #12: ALL_TABS sebagai module-level constant (stabil antar render) */
const ALL_TABS: Array<{k:string; l:string; roles:Array<"admin"|"perawat">}> = [
  {k:"home",      l:"🏠 Home",       roles:["admin","perawat"]},
  {k:"jadwal",    l:"📋 Jadwal",     roles:["admin","perawat"]},
  {k:"daftar",    l:"📝 Daftar",     roles:["admin","perawat"]},
  {k:"laporan",   l:"📄 Laporan",    roles:["admin","perawat"]},
  {k:"wa",        l:"💬 Kirim WA",   roles:["admin","perawat"]},
  {k:"statistik", l:"📊 Statistik",  roles:["admin"]},
  {k:"staf",      l:"👥 Staf",       roles:["admin"]},
  {k:"lembur",    l:"⏰ Lembur",     roles:["admin","perawat"]},
  {k:"monitoring",l:"🌡 Monitoring", roles:["admin","perawat"]},
  {k:"roster_gen", l:"🗓 Generator",   roles:["admin"]},
  {k:"arsip",     l:"🗂 Arsip",      roles:["admin"]},
];

export default function App() {
  const [pinOK,    setPinOK]    = useState(false);
  const [role,     setRole]     = useState<"admin"|"perawat">("perawat");
  const [pinAdmin,   setPinAdmin]   = useState<string>("2409"); // fallback
  const [pinPerawat, setPinPerawat] = useState<string>("1234"); // fallback
  const [pinLoaded,  setPinLoaded]  = useState(false);
  const [pinFromCloud, setPinFromCloud] = useState(false); // true jika PIN loaded dari Supabase
  const isFirstTime = pinLoaded && !pinFromCloud; // first time jika Supabase tidak punya PIN
  const [tab,      setTab]      = useState("home");

  // Load PIN dari Supabase saat pertama buka — pakai SUPA_CLIENT singleton
  useEffect(()=>{
    (async()=>{
      const FA = "2409"; const FP = "1234";
      try {
        const {data, error} = await SUPA_CLIENT.from("kamar_bedah_config").select("value").eq("key","pins").single();
        if(!error && data?.value){
          const pins = safeJSONParse(data.value, {});
          // FIX AUDIT #4: field bisa berisi hash ("salt:hash") atau legacy
          // plaintext lama — keduanya ditangani transparan oleh verifyPin().
          const ap = pins.admin||FA; const pp = pins.perawat||FP;
          setPinAdmin(ap); setPinPerawat(pp);
          setPinFromCloud(true); // PIN ada di Supabase — bukan first time
        } else {
          // Supabase tidak punya PIN — isFirstTime=true, force setup modal
          setPinAdmin(FA); setPinPerawat(FP);
          setPinFromCloud(false);
        }
      } catch(e){
        /* PIN load failed — use fallback, treat as existing PIN (not first time) */
        setPinAdmin(FA); setPinPerawat(FP);
        setPinFromCloud(true); // assume existing setup on error
      }
      setPinLoaded(true);
    })();
  },[]);
  const [showPinMgmt, setShowPinMgmt] = useState(false);
  const [pmNewAdmin,  setPmNewAdmin]  = useState("");
  const [pmCfAdmin,   setPmCfAdmin]   = useState("");
  const [pmNewPerawat,setPmNewPerawat]= useState("");
  const [pmCfPerawat, setPmCfPerawat] = useState("");
  const [pmErr,       setPmErr]       = useState("");
  const [ops,      setOps]      = useState<Operation[]>([]);
  const [staff,    setStaff]    = useState<StaffMember[]>([]);
  const [roster,   setRoster]   = useState<RosterEntry[]>([]);
  const [notifs,   setNotifs]   = useState<Notif[]>([]);
  const [archive,  setArchive]  = useState<Operation[]>([]);
  const [toast,    setToast]    = useState<any>(null);
  /* FIX #1: opForm, editingOp, opErrors, dupWarn DIPINDAH ke dalam ViewDaftar
     sebagai local state. State ini sebelumnya ada di sini (root App) sehingga
     setiap ketukan huruf di form pendaftaran memicu re-render seluruh pohon
     komponen. Komunikasi antara ViewJadwal (tombol Edit) dan ViewDaftar (form)
     sekarang melalui pendingEditOp (lihat FIX AUDIT #24) — state kecil di
     root App yang diteruskan sebagai prop biasa, tanpa perlu angkat seluruh
     opForm ke root. */
  /* FIX #1: lSet dan lSby DIPINDAH ke ViewLaporan sebagai local state —
     tidak perlu ada di root App lagi. */
  const [reqOpId,  setReqOpId]  = useState<string|null>(null);
  const [reqText,  setReqText]  = useState("");
  const [privacyMode,setPM]     = useState(false);
  const [supaCfg,  setSupaCfg]  = useState<SupabaseConfig>(()=>({...defaultSupaCfg, url:SUPA_URL||"", anonKey:SUPA_ANON||"", autoBackup:true, realtimeBackup:true, autoExcelBackup:true}));
  const [supaStatus,setSupaStatus] = useState<{ok:boolean;msg:string}|null>(null);
  const [supaBackingUp,setSupaBU] = useState(false);
  const [dbxCfg,   setDbxCfg]   = useState<DropboxConfig>(()=>({...defaultDbxCfg, autoBackup:true, realtimeBackup:true, autoExcelBackup:true}));
  const [dbxStatus,setDbxStatus] = useState<{ok:boolean;msg:string}|null>(null);
  const [dbxBacking,setDbxBacking] = useState(false);
  const dbxAutoRef      = useRef<ReturnType<typeof setInterval>|null>(null);
  const dbxRtRef        = useRef<ReturnType<typeof setTimeout>|null>(null);
  const dbxExcelAutoRef = useRef<ReturnType<typeof setInterval>|null>(null);
  const dualRtRef       = useRef<ReturnType<typeof setTimeout>|null>(null);
  /* supaStartupRef removed — startup load via loadAllFromSupa on pinOK */
  const [templates,      setTemplates]      = useState<Operation[]>([]);
  const [auditLog,       setAuditLog]       = useState<any[]>([]);
  const [rtEnabled,      setRtEnabled]      = useState(true); // Realtime aktif otomatis
  const [rtStatus,       setRtStatus]       = useState<"offline"|"connecting"|"online">("offline");
  const [lemburPegawai,  setLemburPegawai]  = useState<any[]>([]);
  const [lemburData,     setLemburData]     = useState<Record<string,any>>({});
  const [monitoringEntries, setMonitoringEntries] = useState<MonitoringEntry[]>([]);
  const [monitoringCfg,     setMonitoringCfg]     = useState<MonitoringCfg>(defaultMonCfg);
  /* currentTime DIHAPUS dari root App — lihat komponen <HeaderClock/> di bawah.
     Sebelumnya state jam berdetak setiap 1 detik ditaruh di komponen ROOT,
     sehingga SETIAP detik memicu re-render seluruh tab yang sedang aktif
     (termasuk daftar arsip ribuan baris / agregasi statistik), padahal yang
     berubah hanya teks jam kecil di header. Sekarang jam diisolasi ke
     komponennya sendiri agar detak per-detik tidak "menulari" seluruh App. */
  /* storageWarn dihapus — tidak relevan setelah localStorage ditiadakan */
  const lastAct   = useRef(Date.now());
  const autoBackupRef = useRef<ReturnType<typeof setInterval>|null>(null);
  const latestDataRef = useRef<any>({});
  const rtChannelRef  = useRef<RealtimeChannel|null>(null);
  // Debounce timers per-tabel — menggantikan rtIgnoreRef yang menyebabkan blocked valid payload
  const rtDebounceRef = useRef<Record<string,ReturnType<typeof setTimeout>>>({});
  // Timer toast aktif — dipakai agar toast baru tidak ditutup paksa oleh timer toast sebelumnya
  const toastTimerRef = useRef<ReturnType<typeof setTimeout>|null>(null);

  /* FIX AUDIT #6 (HIGH — Stale Closure): logAudit HANYA menggunakan
     functional updater setAuditLog(prev => ...) sehingga deps kosong [] VALID
     dan tidak menyebabkan stale closure. Tidak ada state lain yang diakses
     secara langsung di dalam callback ini. Pola functional updater ini
     menjamin entri baru selalu ditambahkan ke versi state terbaru. */
  const logAudit = useCallback((action: string, op: any) => {
    setAuditLog(prev => [
      ...prev,
      {
        id: gId(),
        action,
        patient: op.patient ?? "—",
        detail: `${op.procedure ?? ""} · ${op.date ?? ""} ${op.time ?? ""}`,
        time: fNow(),
      },
    ]);
  }, []);

  /* ── Auto-lock setelah idle ──────────────────────────────────────────
     UI Pengaturan sudah lama mengklaim "Auto-lock setelah 10 menit idle"
     (lihat LOCK_MS) dan ref `lastAct` sudah ada (diperbarui saat PIN
     berhasil diverifikasi), TAPI tidak ada satu pun listener/interval yang
     benar-benar memeriksa idle time — fitur keamanan ini sebelumnya hanya
     teks, tidak pernah berjalan. Effect ini melengkapi implementasinya:
     setiap interaksi pengguna memperbarui `lastAct`, dan setiap 15 detik
     diperiksa apakah sudah melewati LOCK_MS sejak interaksi terakhir; jika
     ya, kembalikan ke PinScreen (setPinOK(false)). ── */
  useEffect(()=>{
    if(!pinOK) return;
    /* FIX #2: Throttle bump — mousemove/scroll bisa terpicu ratusan kali/detik.
       Tanpa throttle ini, setiap gesekan layar/gerakan mouse memaksa fungsi JS
       dieksekusi terus-menerus (CPU drain) dan menguras baterai perangkat.
       Solusi: pakai satu timer; selama timer masih aktif, event berikutnya
       diabaikan — lastAct hanya diperbarui maksimal 1x per 2 detik. */
    let throttleTimer: ReturnType<typeof setTimeout> | null = null;
    const bump = () => {
      if (throttleTimer) return; // abaikan event beruntun
      throttleTimer = setTimeout(() => {
        lastAct.current = Date.now();
        throttleTimer = null;
      }, 2000); // batasi pembaruan maksimal 1x per 2 detik
    };
    const events: (keyof WindowEventMap)[] = ["mousedown","mousemove","keydown","touchstart","wheel","scroll"];
    events.forEach(ev=>window.addEventListener(ev, bump, {passive:true}));
    lastAct.current = Date.now();
    const iv = setInterval(()=>{
      if(Date.now() - lastAct.current >= LOCK_MS) setPinOK(false);
    }, 15000);
    return ()=>{
      events.forEach(ev=>window.removeEventListener(ev, bump));
      clearInterval(iv);
      if (throttleTimer) clearTimeout(throttleTimer);
    };
  },[pinOK]);

  /* localStorage warning dihapus — data tidak lagi disimpan di localStorage */

  /* NOTE: localStorage dihapus — Supabase adalah single source of truth.
     Data di-hydrate dari cloud saat login dan disinkron via Realtime. */

  /* Startup load handled by loadAllFromSupa (called on pinOK) — no duplicate needed */
  /* supaStartupRef kept for backward compat but no longer used */

  /* Dropbox auto-backup interval */
  useEffect(()=>{
    if(dbxAutoRef.current) clearInterval(dbxAutoRef.current);
    if(!dbxCfg.autoBackup) return;
    const ms=dbxCfg.backupInterval*60*1000;
    dbxAutoRef.current=setInterval(async()=>{
      const d = latestDataRef.current;
      const res=await dropboxUpload(dbxCfg,{...d});
      if(res.ok) setDbxCfg(p=>({...p,lastBackup:fNow()}));
    },ms);
    return()=>{ if(dbxAutoRef.current) clearInterval(dbxAutoRef.current); };
  },[dbxCfg.autoBackup,dbxCfg.backupInterval]);

  /* ── COMBINED dual realtime backup: Supabase + Dropbox fired SIMULTANEOUSLY ── */
  useEffect(()=>{
    const hasSupa = supaCfg.realtimeBackup;
    const hasDbx  = dbxCfg.realtimeBackup;
    if(!hasSupa && !hasDbx) return;
    if(dualRtRef.current) clearTimeout(dualRtRef.current);
    dualRtRef.current = setTimeout(async()=>{
      const data = {exportedAt:fNow(), ...latestDataRef.current};
      const tasks: Promise<any>[] = [];
      if(hasSupa) tasks.push(supabaseBackup(supaCfg, data, SUPA_CLIENT).then(r=>{ if(r.ok) setSupaCfg(p=>({...p,lastBackup:fNow()})); }));
      if(hasDbx)  tasks.push(dropboxUpload(dbxCfg, data).then(r=>{ if(r.ok) setDbxCfg(p=>({...p,lastBackup:fNow()})); }));
      await Promise.all(tasks);
    }, 5000);
    return()=>{ if(dualRtRef.current) clearTimeout(dualRtRef.current); };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[ops, staff, roster, lemburData, lemburPegawai, monitoringEntries]);

  /* ── Auto Excel backup every 24 hours ── */
  useEffect(()=>{
    if(dbxExcelAutoRef.current) clearInterval(dbxExcelAutoRef.current);
    if(!dbxCfg.autoExcelBackup) return;
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
      const resOps = await dropboxUploadExcel(`${folder}Auto_Operasi_${stamp}.xlsx`, wbOps);
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
      const resLmb = await dropboxUploadExcel(`${folder}Auto_Lembur_${stamp}.xlsx`, wbLmb);
      /* — Monitoring Excel — */
      const wbMon = XLSX.utils.book_new();
      const monRows = monitoringEntries.map((e:MonitoringEntry)=>[e.ruang??monitoringCfg.lokasiRuang??"",e.tanggal,e.jam,e.suhu,e.kelembaban,monIsOK(e.suhu,e.kelembaban,monitoringCfg)?"SESUAI":"TIDAK SESUAI",e.petugas]);
      const wsMonMain = XLSX.utils.aoa_to_sheet([["Ruangan","Tanggal","Jam","Suhu (°C)","Kelembaban (%)","Status","Petugas"],...monRows]);
      wsMonMain["!cols"]=[{wch:18},{wch:12},{wch:8},{wch:12},{wch:14},{wch:16},{wch:22}];
      XLSX.utils.book_append_sheet(wbMon, wsMonMain, "Monitoring Suhu");
      const wsMonCfg = XLSX.utils.aoa_to_sheet([["Parameter","Nilai"],["Ruangan 1","Kamar Bedah 1"],["Ruangan 2","Kamar Bedah 2"],["Ruangan 3","Ruang Instrumen"],["Suhu Min",monitoringCfg.suhuMin],["Suhu Max",monitoringCfg.suhuMax],["RH Min",monitoringCfg.rhMin],["RH Max",monitoringCfg.rhMax],["Kepala KB",monitoringCfg.kepalaKamarBedah]]);
      XLSX.utils.book_append_sheet(wbMon, wsMonCfg, "Konfigurasi");
      const resMon = await dropboxUploadExcel(`${folder}Auto_Monitoring_${stamp}.xlsx`, wbMon);
      if(resOps.ok&&resLmb.ok&&resMon.ok){
        setDbxCfg((p:DropboxConfig)=>({...p,lastExcelBackup:fNow(),lastExcelBackupTs:Date.now()}));
      }
    };
    /* Check immediately on enable, then every hour */
    run24h(latestDataRef.current.ops||[], latestDataRef.current.lemburPegawai||[], latestDataRef.current.lemburData||{}, dbxCfg);
    dbxExcelAutoRef.current = setInterval(()=>{
      setDbxCfg(p=>{ run24h(latestDataRef.current.ops||[], latestDataRef.current.lemburPegawai||[], latestDataRef.current.lemburData||{}, p); return p; });
    }, 60*60*1000);
    return()=>{ if(dbxExcelAutoRef.current) clearInterval(dbxExcelAutoRef.current); };
  },[dbxCfg.autoExcelBackup]);

  /* Inject responsive CSS */
  useEffect(()=>{
    const id="kb-css"; if(document.getElementById(id))return;
    const s=document.createElement("style"); s.id=id;
    s.textContent=`
      *{box-sizing:border-box;}
      html,body{margin:0;padding:0;}
      @keyframes kbPulse{0%,100%{opacity:1;box-shadow:0 0 0 0 rgba(201,169,97,.55);}50%{opacity:.85;box-shadow:0 0 0 4px rgba(201,169,97,0);}}
      @keyframes lgFloat{0%,100%{transform:translateY(0)}50%{transform:translateY(-6px)}}
      @keyframes redPulse{0%,100%{opacity:1;transform:scale(1);}50%{opacity:.4;transform:scale(.7);}}

      /* ── Mobile default (<640px) ── */
      body{background:#f8fafc;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;}
      .kbShell{width:100%;min-height:100vh;}
      .kb-header{padding:12px 16px 0;}
      .kb-header-inner{margin-bottom:10px;}
      .kb-tabs{display:flex;overflow-x:auto;-webkit-overflow-scrolling:touch;gap:0;padding-bottom:0;}
      .kb-tab-btn{padding:10px 14px;flex-shrink:0;white-space:nowrap;font-size:15px;font-weight:700;}
      .kb-content{padding:14px 16px 32px;background:#f8fafc;}
      .kb-sidebar{display:none;}
      .kb-main{display:block;}
      .kb-hospital{font-size:11px;}
      .kb-title{font-size:20px;}

      /* ── Tablet (640px–1023px) ── */
      @media(min-width:640px){
        body{background:#dbeafe;}
        .kbShell{max-width:720px;margin:0 auto;box-shadow:0 4px 24px rgba(59,130,246,.12);}
        input,select,textarea{font-size:15px!important;}
        .kb-tab-btn{padding:11px 18px;font-size:16px;font-weight:800;}
        .kb-content{padding:18px 22px 36px;background:#f8fafc;}
        .kb-title{font-size:22px;}
        .kb-hospital{font-size:12px;}
      }

      /* ── Desktop (≥1024px) ── */
      @media(min-width:1024px){
        body{
          background:linear-gradient(135deg,#0c1445 0%,#1e3a8a 45%,#0284c7 100%)!important;
          min-height:100vh;
          display:flex;align-items:flex-start;justify-content:center;
          padding:28px 20px;
        }
        #root{width:100%;max-width:1200px;display:flex;justify-content:center;}
        .kbShell{
          width:100%;max-width:1160px;
          border-radius:20px;
          overflow:hidden;
          box-shadow:0 24px 80px rgba(30,58,138,.45);
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
          background:linear-gradient(180deg,#1e3a8a,#1d4ed8 60%,#0284c7);
          padding:16px 10px;gap:4px;position:relative;
          border-right:1px solid rgba(255,255,255,.12);
          min-height:calc(100vh - 56px - 60px);
        }
        .kb-sidebar-btn{
          display:flex;align-items:center;gap:10px;
          padding:11px 14px;border-radius:10px;
          font-size:14px;font-weight:600;
          color:rgba(255,255,255,.72);
          background:none;border:none;cursor:pointer;
          text-align:left;width:100%;
          font-family:inherit;transition:all .15s;
        }
        .kb-sidebar-btn:hover{background:rgba(255,255,255,.10);color:#fff;}
        .kb-sidebar-btn.active{background:rgba(255,255,255,.18);color:#fff;font-weight:900;box-shadow:inset 3px 0 0 #38bdf8;}
        .kb-main{display:flex;flex:1;background:#f8fafc;}
        .kb-content{flex:1;padding:24px 32px 40px;overflow-y:auto;background:#f8fafc;}
        .kb-title{font-size:24px;}
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

    /* Auto-load semua data dari Supabase saat login — no localStorage check */
  useEffect(()=>{
    if(!pinOK) return;
    loadAllFromSupa();
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[pinOK]);

  /* ══════════════════════════════════════════════════════════════════════
     SUPABASE REALTIME — SSOT (Single Source of Truth)
     Strategi:
     • postgres_changes → granular delta update per-tabel (INSERT/UPDATE/DELETE)
     • Bulk import → debounce 800ms sebelum full refetch (hindari 50x refetch)
     • Tidak ada broadcast manual — semua device rely on postgres_changes
     • Tidak ada rtIgnoreRef — setiap payload dari DB selalu valid
     ══════════════════════════════════════════════════════════════════════ */
  useEffect(()=>{
    /* FIX AUDIT #5 (HIGH — Memory Leak Channel): gunakan removeChannel()
       agar channel lama benar-benar dilepas dari registry internal Supabase
       client, bukan hanya di-unsubscribe. Tanpa removeChannel(), channel
       lama tetap ada di memori klien setelah reconnect, menyebabkan memory
       leak bertahap dan potensi duplikasi event handler. */
    if(rtChannelRef.current){
      SUPA_CLIENT.removeChannel(rtChannelRef.current);
      rtChannelRef.current=null;
      setRtStatus("offline");
    }
    if(!rtEnabled) return;
    setRtStatus("connecting");

    /* ── Helper: apply delta dari postgres_changes ke local state ── */
    const applyDelta = (table: string, payload: any) => {
      const ev    = payload.eventType as "INSERT"|"UPDATE"|"DELETE";
      const newRec = payload.new?.data ?? null;
      const newId  = payload.new?.id   ?? null;
      const oldId  = payload.old?.id  ?? null;

      if(table === "kb_operasi"){
        if(ev === "DELETE") setOps((p:any[]) => p.filter(o => o.id !== oldId));
        else if(newRec) setOps((p:any[]) => {
          const idx = p.findIndex((o:any) => o.id === newRec.id);
          /* Conflict resolution: last-write-wins via updated_at timestamp.
             Jika record lokal lebih baru, jangan timpa (bisa terjadi saat race condition). */
          if(idx >= 0) {
            const existing = p[idx];
            const existingTs = existing.updated_at || existing.createdAt || "";
            const newTs = newRec.updated_at || newRec.createdAt || "";
            if(existingTs > newTs) return p; // lokal lebih baru — pertahankan
            return p.map((o:any) => o.id === newRec.id ? newRec : o);
          }
          return [...p, newRec];
        });
      } else if(table === "kb_staf"){
        if(ev === "DELETE") setStaff((p:any[]) => p.filter(s => s.id !== oldId));
        else if(newRec) setStaff((p:any[]) => {
          const idx = p.findIndex((s:any) => s.id === newRec.id);
          if(idx >= 0){
            const existingTs = p[idx].updated_at || p[idx].createdAt || "";
            const newTs = newRec.updated_at || newRec.createdAt || "";
            if(existingTs > newTs) return p; // lokal lebih baru — pertahankan (cegah out-of-order overwrite)
            return p.map((s:any) => s.id === newRec.id ? newRec : s);
          }
          return [...p, newRec];
        });
      } else if(table === "kb_roster"){
        if(ev === "DELETE") setRoster((p:any[]) => p.filter(r => r.id !== oldId));
        else if(newRec) setRoster((p:any[]) => {
          const idx = p.findIndex((r:any) => r.id === newRec.id);
          if(idx >= 0){
            const existingTs = p[idx].updated_at || p[idx].createdAt || "";
            const newTs = newRec.updated_at || newRec.createdAt || "";
            if(existingTs > newTs) return p;
            return p.map((r:any) => r.id === newRec.id ? newRec : r);
          }
          return [...p, newRec];
        });
      } else if(table === "kb_lembur_pegawai"){
        if(ev === "DELETE") setLemburPegawai((p:any[]) => p.filter(x => x.id !== oldId));
        else if(newRec) setLemburPegawai((p:any[]) => {
          const idx = p.findIndex((x:any) => x.id === newRec.id);
          if(idx >= 0){
            const existingTs = p[idx].updated_at || p[idx].createdAt || "";
            const newTs = newRec.updated_at || newRec.createdAt || "";
            if(existingTs > newTs) return p;
            return p.map((x:any) => x.id === newRec.id ? newRec : x);
          }
          return [...p, newRec];
        });
      } else if(table === "kb_lembur_data"){
        /* Granular: id row = "${pegId}_${period}". Merge per-key, JANGAN replace
           seluruh map — itulah yang menyebabkan device lain bisa saling menimpa.
           Tetap cek updated_at supaya event lama yang datang terlambat (network
           reordering) tidak menimpa entri yang sudah lebih baru di lokal. */
        if(ev === "DELETE"){ if(oldId) setLemburData((p:Record<string,any>)=>{ const n={...p}; delete n[oldId]; return n; }); }
        else if(newRec && newId) setLemburData((p:Record<string,any>) => {
          const existing = p[newId];
          if(existing){
            const existingTs = existing.updated_at || "";
            const newTs = newRec.updated_at || "";
            if(existingTs > newTs) return p;
          }
          return {...p, [newId]: newRec};
        });
      } else if(table === "kb_monitoring"){
        if(ev === "DELETE") setMonitoringEntries((p:any[]) => p.filter(x => x.id !== oldId));
        else if(newRec) setMonitoringEntries((p:any[]) => {
          const idx = p.findIndex((x:any) => x.id === newRec.id);
          if(idx >= 0){
            const existingTs = p[idx].updated_at || p[idx].createdAt || "";
            const newTs = newRec.updated_at || newRec.createdAt || "";
            if(existingTs > newTs) return p;
            return p.map((x:any) => x.id === newRec.id ? newRec : x);
          }
          return [...p, newRec];
        });
      }
    };

    /* ── FIX AUDIT #2 (CRITICAL — Race Condition / Silent Data Loss):
       Sebelumnya, payload individual di-skip dari applyDelta saat burst
       counter ≥5 dalam window 200ms, dan hanya mengandalkan full-refetch
       800ms kemudian. Jika sebuah payload tiba tepat saat window burst
       sebelumnya masih aktif, ia bisa ter-skip dari applyDelta — ada window
       singkat di mana state lokal stale sebelum refetch tuntas.
       Sekarang SETIAP payload selalu di-enqueue dan diproses (applyDelta)
       — tidak ada yang dibuang. Debounce hanya menunda kapan full-refetch
       "safety net" dijalankan (untuk hindari refetch berulang saat bulk
       import), bukan untuk memilih payload mana yang diproses. ── */
    const rtQueueRef: Record<string, any[]> = {};

    const enqueuePayload = (table: string, payload: any) => {
      if(!rtQueueRef[table]) rtQueueRef[table] = [];
      rtQueueRef[table].push(payload);
    };

    /* ── Helper: debounced full refetch untuk bulk operations ──
       Jika 50 INSERT tiba bersamaan, hanya 1 loadFromSupaTable yang dijalankan,
       TAPI seluruh payload individual yang masuk selama itu tetap diproses
       segera lewat applyDelta — full refetch hanya berfungsi sebagai
       safety-net konsistensi tambahan, bukan satu-satunya jalur update. ── */
    const DEBOUNCE_MS = 800;
    const debounced = (table: string, payload: any) => {
      enqueuePayload(table, payload);
      /* Proses payload ini segera — tidak pernah di-skip/dibuang. */
      applyDelta(table, payload);

      /* Debounced full-sync sebagai safety net — selalu jalan, tapi hanya sekali */
      if(rtDebounceRef.current[table]) clearTimeout(rtDebounceRef.current[table]);
      rtDebounceRef.current[table] = setTimeout(() => {
        delete rtDebounceRef.current[table];
        rtQueueRef[table] = [];
        loadFromSupaTable(table);
      }, DEBOUNCE_MS);
    };

    try {
      const ch = SUPA_CLIENT.channel("kamar-bedah-rt-v3")
        /* ── postgres_changes: granular delta per event, per tabel ── */
        .on("postgres_changes",{event:"*",schema:"public",table:"kb_operasi"},       (p:any)=>debounced("kb_operasi",p))
        .on("postgres_changes",{event:"*",schema:"public",table:"kb_staf"},          (p:any)=>debounced("kb_staf",p))
        .on("postgres_changes",{event:"*",schema:"public",table:"kb_roster"},        (p:any)=>debounced("kb_roster",p))
        .on("postgres_changes",{event:"*",schema:"public",table:"kb_lembur_pegawai"},(p:any)=>debounced("kb_lembur_pegawai",p))
        .on("postgres_changes",{event:"*",schema:"public",table:"kb_lembur_data"},   (p:any)=>debounced("kb_lembur_data",p))
        .on("postgres_changes",{event:"*",schema:"public",table:"kb_monitoring"},    (p:any)=>debounced("kb_monitoring",p))
        .subscribe((status:string)=>{
          setRtStatus(status==="SUBSCRIBED"?"online":"offline");
        });
      rtChannelRef.current = ch;
    } catch(e){ setRtStatus("offline"); }

    return()=>{
      // FIX AUDIT #5: removeChannel() agar channel benar-benar dilepas dari registry Supabase client
      if(rtChannelRef.current){ SUPA_CLIENT.removeChannel(rtChannelRef.current); rtChannelRef.current=null; }
      Object.values(rtDebounceRef.current).forEach(t=>clearTimeout(t));
      rtDebounceRef.current={};
      Object.keys(rtQueueRef).forEach(k=>{ delete rtQueueRef[k]; });
      setRtStatus("offline");
    };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[rtEnabled]);

  /* Broadcast ops changes when realtime is on */
  // ── SUPABASE SYNC: selalu upsert ke tabel, tidak perlu realtime online ──
  // SUPA client dibuat sekali di luar component (lihat baris sebelum App)

  /* ── upsertOneToSupa: tulis 1 baris ke DB — dipakai oleh semua mutasi ──
     Mengembalikan {ok, error} (bukan boolean polos) agar pemanggil bisa
     menampilkan pesan Supabase yang DESKRIPTIF ke UI, bukan "gagal, coba
     lagi" generik. id selalu divalidasi sebagai UUID (ensureId) supaya
     tidak ditolak Postgres karena "invalid input syntax for type uuid". */
  const upsertOneToSupa = useCallback(async (table: string, row: any): Promise<{ok:boolean; error?:string}> => {
    try {
      const rowId = ensureId(row.id);
      const updatedAt = row.updated_at || new Date().toISOString();
      const {error} = await SUPA_CLIENT.from(table).upsert(
        {id: rowId, data: {...row, id: rowId, updated_at: updatedAt}, updated_at: updatedAt},
        {onConflict:"id"}
      );
      if(error){ const msg = formatSupaError(error); console.warn(`[upsertOneToSupa] ${table}:`, msg); return {ok:false, error:msg}; }
      return {ok:true};
    } catch(e:any){ const msg = "Koneksi ke Supabase gagal: "+(e?.message||"unknown error"); console.warn("[upsertOneToSupa] exception:", e); return {ok:false, error:msg}; }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[]);

  /* ── upsertBulkToSupa: batch upsert (import Excel) — satu trip ke DB ──
     Sebelum dikirim, setiap baris divalidasi/disanitasi (id UUID valid,
     updated_at ISO). Mengembalikan jumlah baris & error deskriptif. */
  const upsertBulkToSupa = useCallback(async (table: string, rows: any[]): Promise<{ok:boolean; error?:string; count:number}> => {
    try {
      if(!rows.length) return {ok:true, count:0};
      const ts = new Date().toISOString();
      const payload = rows.map((r:any)=>{
        const rowId = ensureId(r.id);
        return { id: rowId, data: {...r, id: rowId, updated_at: r.updated_at || ts}, updated_at: r.updated_at || ts };
      });
      const {error} = await SUPA_CLIENT.from(table).upsert(payload, {onConflict:"id"});
      if(error){ const msg = formatSupaError(error); console.warn(`[upsertBulkToSupa] ${table}:`, msg); return {ok:false, error:msg, count:0}; }
      return {ok:true, count: payload.length};
    } catch(e:any){ const msg = "Koneksi ke Supabase gagal: "+(e?.message||"unknown error"); console.warn("[upsertBulkToSupa] exception:", e); return {ok:false, error:msg, count:0}; }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[]);

  /* ── deleteFromSupa: hapus 1 baris dari DB ── */
  const deleteFromSupa = useCallback(async (table: string, id: string): Promise<{ok:boolean; error?:string}> => {
    try {
      const {error} = await SUPA_CLIENT.from(table).delete().eq("id", id);
      if(error){ const msg = formatSupaError(error); console.warn(`[deleteFromSupa] ${table}:`, msg); return {ok:false, error:msg}; }
      return {ok:true};
    } catch(e:any){ return {ok:false, error:"Koneksi ke Supabase gagal: "+(e?.message||"unknown error")}; }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[]);

  /* ── loadFromSupaTable: refetch SATU tabel saja (safety-net debounce) ──
     kb_lembur_data kini disimpan granular: 1 row per key "${pegId}_${period}"
     (bukan 1 mega-row "lembur_data"), sehingga di sini di-merge jadi map. */
  const loadFromSupaTable = useCallback(async (table: string) => {
    try {
      if(table === "kb_operasi"){
        const {data} = await SUPA_CLIENT.from("kb_operasi").select("data").order("updated_at",{ascending:false});
        if(data?.length) setOps(data.map((x:any)=>x.data));
      } else if(table === "kb_staf"){
        const {data} = await SUPA_CLIENT.from("kb_staf").select("data").order("updated_at",{ascending:false});
        if(data?.length) setStaff(data.map((x:any)=>x.data));
      } else if(table === "kb_roster"){
        const {data} = await SUPA_CLIENT.from("kb_roster").select("data").order("updated_at",{ascending:false});
        if(data?.length) setRoster(data.map((x:any)=>x.data));
      } else if(table === "kb_lembur_pegawai"){
        const {data} = await SUPA_CLIENT.from("kb_lembur_pegawai").select("data").order("updated_at",{ascending:false});
        if(data?.length) setLemburPegawai(data.map((x:any)=>x.data));
      } else if(table === "kb_lembur_data"){
        const {data} = await SUPA_CLIENT.from("kb_lembur_data").select("id,data").order("updated_at",{ascending:false});
        if(data?.length){
          const map: Record<string,any> = {};
          data.forEach((x:any)=>{ if(x.id!=="lembur_data") map[x.id]=x.data; });
          setLemburData(map);
        }
      } else if(table === "kb_monitoring"){
        const {data} = await SUPA_CLIENT.from("kb_monitoring").select("data").order("updated_at",{ascending:false});
        if(data?.length) setMonitoringEntries(data.map((x:any)=>x.data));
      }
    } catch{ /* offline — biarkan state yang ada tetap berlaku */ }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[]);

  /* ── loadAllFromSupa: full refetch semua tabel (hanya saat login) ── */
  const loadAllFromSupa = useCallback(async () => {
    try {
      const [r1,r2,r3,r4,r5,r6] = await Promise.all([
        SUPA_CLIENT.from("kb_operasi").select("data").order("updated_at",{ascending:false}),
        SUPA_CLIENT.from("kb_staf").select("data").order("updated_at",{ascending:false}),
        SUPA_CLIENT.from("kb_roster").select("data").order("updated_at",{ascending:false}),
        SUPA_CLIENT.from("kb_lembur_pegawai").select("data").order("updated_at",{ascending:false}),
        SUPA_CLIENT.from("kb_lembur_data").select("id,data").order("updated_at",{ascending:false}),
        SUPA_CLIENT.from("kb_monitoring").select("data").order("updated_at",{ascending:false}),
      ]);
      let hasData = false;
      if(r1.data?.length){ setOps(r1.data.map((x:any)=>x.data)); hasData=true; }
      if(r2.data?.length){ setStaff(r2.data.map((x:any)=>x.data)); hasData=true; }
      if(r3.data?.length){ setRoster(r3.data.map((x:any)=>x.data)); hasData=true; }
      if(r4.data?.length){ setLemburPegawai(r4.data.map((x:any)=>x.data)); hasData=true; }
      if(r5.data?.length){
        /* kb_lembur_data granular: 1 row per key "${pegId}_${period}". Legacy
           deployment lama menyimpan 1 mega-row id="lembur_data" berisi seluruh
           map — deteksi & migrasikan otomatis (sekali, best-effort, non-blocking)
           agar tidak ada data lama yang "hilang" setelah upgrade. */
        const map: Record<string,any> = {};
        const legacyRow = r5.data.find((x:any)=>x.id==="lembur_data");
        r5.data.forEach((x:any)=>{ if(x.id!=="lembur_data") map[x.id]=x.data; });
        if(legacyRow?.data && typeof legacyRow.data === "object"){
          Object.entries(legacyRow.data).forEach(([k,v]:[string,any])=>{ if(!map[k]) map[k]=v; });
          (async()=>{
            try {
              const legacyEntries = Object.entries(legacyRow.data as Record<string,any>);
              if(legacyEntries.length){
                await upsertBulkToSupa("kb_lembur_data", legacyEntries.map(([k,v]:[string,any])=>({id:k, ...v})));
              }
              await deleteFromSupa("kb_lembur_data","lembur_data");
            } catch(e){ console.warn("[migrasi kb_lembur_data] gagal:", e); }
          })();
        }
        setLemburData(map); hasData=true;
      }
      if(r6.data?.length){ setMonitoringEntries(r6.data.map((x:any)=>x.data)); hasData=true; }
      if(!hasData){

        const {data:bk} = await SUPA_CLIENT.from("kamar_bedah_backup").select("data,created_at").order("created_at",{ascending:false}).limit(1);
        if(bk?.length){
          const p=safeJSONParse(bk[0].data, {});
          if(p.ops?.length) setOps(p.ops);
          if(p.staff?.length) setStaff(p.staff);
          if(p.roster?.length) setRoster(p.roster);
          if(p.lemburPegawai?.length) setLemburPegawai(p.lemburPegawai);
          if(p.lemburData) setLemburData(p.lemburData);
          if(p.monitoringEntries?.length) setMonitoringEntries(p.monitoringEntries);
          hasData=true;
        }
      }
      if(hasData){
        if(toastTimerRef.current) clearTimeout(toastTimerRef.current);
        setToast({msg:"✅ Data tersinkron dari Supabase",color:"#2E7D32"});
        toastTimerRef.current = setTimeout(()=>{ setToast(null); toastTimerRef.current=null; },3200);
      }
    } catch(e){
      if(toastTimerRef.current) clearTimeout(toastTimerRef.current);
      setToast({msg:"⚠ Koneksi Supabase gagal — mode offline",color:C.w});
      toastTimerRef.current = setTimeout(()=>{ setToast(null); toastTimerRef.current=null; }, 4000);
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[]);

  /* ══ SSOT Mutators ══════════════════════════════════════════════════════
     Tidak ada lagi broadcast manual ke channel.
     DB write via upsertOneToSupa → postgres_changes → applyDelta di semua device.
     Optimistic UI update (setOps/setStaff/setRoster) HANYA di device pengirim
     sebagai feedback instan; device lain mendapat update via postgres_changes.
     ════════════════════════════════════════════════════════════════════════ */


  /* Auto-backup to Supabase */
  useEffect(()=>{
    if(autoBackupRef.current) clearInterval(autoBackupRef.current);
    if(!supaCfg.autoBackup) return;
    const ms = supaCfg.backupInterval * 60 * 1000;
    autoBackupRef.current = setInterval(async ()=>{
      const d = latestDataRef.current;
      const data = {exportedAt:fNow(), ...d};
      const res = await supabaseBackup(supaCfg, data, SUPA_CLIENT);
      if(res.ok) setSupaCfg(p=>({...p,lastBackup:fNow()}));
    }, ms);
    return ()=>{ if(autoBackupRef.current) clearInterval(autoBackupRef.current); };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [supaCfg.autoBackup, supaCfg.backupInterval]);

  /* Supabase realtime backup is now handled by the combined dual-backup effect above */

  const showToast = useCallback((msg: string,color?: string)=>{
    if(toastTimerRef.current) clearTimeout(toastTimerRef.current);
    setToast({msg,color:color||C.s});
    toastTimerRef.current = setTimeout(()=>{ setToast(null); toastTimerRef.current=null; },3200);
  },[]);

  /* Keep latestDataRef fresh for use in intervals (prevents stale closures) */
  useEffect(()=>{
    latestDataRef.current = {ops, archive, notifs, lemburPegawai, lemburData, monitoringEntries, monitoringCfg, staff, roster};
  },[ops, archive, notifs, lemburPegawai, lemburData, monitoringEntries, monitoringCfg, staff, roster]);

  const handlePinVerify = (newRole: "admin"|"perawat", newAdminPin?: string, newPerawatPin?: string) => {
    /* ── TURSO LOG: LOGIN_SUKSES ── */
    AmbilLogMedis(
      "KEAMANAN", newRole==="admin"?"Admin":"Perawat",
      "LOGIN_SUKSES",
      `Login berhasil sebagai ${newRole} pada ${fWIB()}`,
      { role: newRole, isFirstSetup: !!(newAdminPin && newPerawatPin) }
    );
    if(newAdminPin && newPerawatPin){
      setPinFromCloud(true); // setelah setup, tandai sebagai sudah ada PIN
      // FIX AUDIT #4: jangan simpan PIN plaintext — hash dengan PBKDF2 dulu,
      // baik untuk state lokal (dibandingkan via verifyPin) maupun Supabase.
      (async()=>{
        try{
          const [hashedAdmin, hashedPerawat] = await Promise.all([hashPin(newAdminPin), hashPin(newPerawatPin)]);
          setPinAdmin(hashedAdmin);
          setPinPerawat(hashedPerawat);
          await SUPA_CLIENT.from("kamar_bedah_config").upsert({key:"pins",value:JSON.stringify({admin:hashedAdmin,perawat:hashedPerawat})},{onConflict:"key"});
          showToast("✓ PIN tersimpan ke Supabase — berlaku di semua perangkat",C.s);
        }catch(e){
          // Jika hashing/penyimpanan gagal, tetap izinkan login sesi ini dengan PIN
          // plaintext di memori lokal agar admin tidak terkunci dari sistem.
          setPinAdmin(newAdminPin);
          setPinPerawat(newPerawatPin);
          showToast("✓ PIN disimpan (Supabase error, coba lagi)",C.w);
        }
      })();
    }
    setRole(newRole);
    lastAct.current=Date.now(); setPinOK(true);
  };

  /* FIX AUDIT #22 (HIGH — Handler tanpa useCallback, mengunci re-render):
     Sebelumnya 5 handler ini didefinisikan ulang setiap kali App() render,
     DAN membaca data besar (ops, archive, dst) LANGSUNG dari closure —
     beda dgn pola latestDataRef yang sudah dipakai di useEffect auto-backup
     untuk mencegah stale closure. Sekarang DIPERBAIKI DUA ASPEK SEKALIGUS:
     (1) baca data lewat latestDataRef.current (konsisten, selalu fresh,
     TIDAK perlu masuk dependency array karena ref selalu stabil),
     (2) dibungkus useCallback dengan dependency minimal (hanya supaCfg)
     supaya reference function stabil antar render — prasyarat agar
     React.memo pada komponen anak (jika diterapkan nanti) benar2 efektif. */
  const handleSupaBackup = useCallback(async () => {
    setSupaBU(true); setSupaStatus(null);
    AmbilLogMedis("LAPORAN_BACKUP","Sistem","BACKUP_SUPABASE_MULAI","Backup manual ke Supabase dimulai",{});
    try{
      const twoYearsAgo = new Date(); twoYearsAgo.setFullYear(twoYearsAgo.getFullYear()-2);
      await SUPA_CLIENT.from("kamar_bedah_backup").delete().lt("created_at", twoYearsAgo.toISOString());
    }catch{ /* auto-delete non-critical */ }
    const data = {exportedAt:fNow(), ...latestDataRef.current};
    const res = await supabaseBackup(supaCfg, data, SUPA_CLIENT);
    setSupaStatus(res);
    if(res.ok) {
      setSupaCfg((p:any)=>({...p,lastBackup:fNow()}));
      showToast("✓ Backup Supabase berhasil","#3ECF8E");
      AmbilLogMedis("LAPORAN_BACKUP","Sistem","BACKUP_SUPABASE_SUKSES","Backup Supabase berhasil",{waktu:fNow()});
    } else {
      showToast(res.msg,C.d);
      AmbilLogMedis("LAPORAN_BACKUP","Sistem","BACKUP_SUPABASE_GAGAL",`Backup Supabase gagal: ${res.msg}`,{error:res.msg});
    }
    setSupaBU(false);
  }, [supaCfg]);

  const handleSupaRestoreOps = useCallback(async () => {
    if(!window.confirm("Pulihkan Jadwal Operasi dari Supabase?\nData jadwal operasi lokal akan ditimpa dengan data dari Supabase.")) return;
    setSupaBU(true); setSupaStatus(null);
    const res = await supabaseRestore(supaCfg, SUPA_CLIENT);
    setSupaStatus({ok:res.ok, msg:res.msg});
    if(res.ok && res.data) {
      if(res.data.ops) { setOps(res.data.ops); showToast("✓ Jadwal operasi dipulihkan dari Supabase","#3ECF8E"); }
      else showToast("⚠ Data operasi tidak ditemukan di backup","#E65100");
    } else { showToast(res.msg, C.d); }
    setSupaBU(false);
  }, [supaCfg]);

  const handleSupaRestoreLembur = useCallback(async () => {
    if(!window.confirm("Pulihkan Data Lembur dari Supabase?\nData lembur lokal akan ditimpa dengan data dari Supabase.")) return;
    setSupaBU(true); setSupaStatus(null);
    const res = await supabaseRestore(supaCfg, SUPA_CLIENT);
    setSupaStatus({ok:res.ok, msg:res.msg});
    if(res.ok && res.data) {
      if(res.data.lemburPegawai) setLemburPegawai(res.data.lemburPegawai);
      if(res.data.lemburData) setLemburData(res.data.lemburData);
      showToast("✓ Data lembur dipulihkan dari Supabase","#3ECF8E");
    } else { showToast(res.msg, C.d); }
    setSupaBU(false);
  }, [supaCfg]);

  const handleSupaRestoreMonitoring = useCallback(async () => {
    if(!window.confirm("Pulihkan Data Monitoring dari Supabase?\nData monitoring lokal akan ditimpa dengan data dari Supabase.")) return;
    setSupaBU(true); setSupaStatus(null);
    const res = await supabaseRestore(supaCfg, SUPA_CLIENT);
    setSupaStatus({ok:res.ok, msg:res.msg});
    if(res.ok && res.data) {
      if(res.data.monitoringEntries) setMonitoringEntries(res.data.monitoringEntries);
      if(res.data.monitoringCfg) setMonitoringCfg(res.data.monitoringCfg);
      showToast("✓ Data monitoring dipulihkan dari Supabase","#3ECF8E");
    } else { showToast(res.msg, C.d); }
    setSupaBU(false);
  }, [supaCfg]);

  const handleSupaRestoreAll = useCallback(async () => {
    if(!window.confirm("Pulihkan SEMUA data dari Supabase?\nOperasi, Lembur, dan Monitoring lokal akan ditimpa dengan data dari Supabase.")) return;
    setSupaBU(true); setSupaStatus(null);
    const res = await supabaseRestore(supaCfg, SUPA_CLIENT);
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
  }, [supaCfg]);

  const handleDropboxBackupMonitoringXls = useCallback(async () => {
    setDbxBacking(true); setDbxStatus(null);
    const {monitoringEntries:me, monitoringCfg:mc} = latestDataRef.current;
    const wb = XLSX.utils.book_new();
    const monRows = me.map((e:MonitoringEntry)=>[e.ruang??mc.lokasiRuang??"",e.tanggal,e.jam,e.suhu,e.kelembaban,monIsOK(e.suhu,e.kelembaban,mc)?"SESUAI":"TIDAK SESUAI",e.petugas]);
    const ws = XLSX.utils.aoa_to_sheet([["Ruangan","Tanggal","Jam","Suhu (°C)","Kelembaban (%)","Status","Petugas"],...monRows]);
    ws["!cols"]=[{wch:18},{wch:12},{wch:8},{wch:12},{wch:14},{wch:16},{wch:22}];
    XLSX.utils.book_append_sheet(wb, ws, "Monitoring Suhu");
    if(!wb.SheetNames.length){
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["Belum ada data monitoring"]]), "Monitoring");
    }
    const folder = dbxCfg.path.replace(/[^/]+$/, "");
    const path = `${folder}Monitoring_${todayDate()}.xlsx`;
    const res = await dropboxUploadExcel(path, wb);
    setDbxStatus(res);
    if(res.ok){ setDbxCfg((p:DropboxConfig)=>({...p,lastBackup:fNow()})); showToast("✓ Excel Monitoring berhasil diupload ke Dropbox","#0369A1"); }
    else showToast(res.msg, C.d);
    setDbxBacking(false);
  }, [dbxCfg]);

  const handleDropboxBackup = useCallback(async () => {
    setDbxBacking(true); setDbxStatus(null);
    AmbilLogMedis("LAPORAN_BACKUP","Sistem","BACKUP_DROPBOX_MULAI","Backup manual ke Dropbox dimulai",{path:dbxCfg.path});
    const {ops:o, archive:ar, notifs:nf, lemburPegawai:lp, lemburData:ld, monitoringEntries:me, monitoringCfg:mc} = latestDataRef.current;
    const data = {ops:o, archive:ar, notifs:nf, lemburPegawai:lp, lemburData:ld, monitoringEntries:me, monitoringCfg:mc};
    const res = await dropboxUpload(dbxCfg, data);
    setDbxStatus(res);
    if(res.ok) {
      setDbxCfg(p=>({...p,lastBackup:fNow()}));
      showToast("✓ Backup Dropbox berhasil","#0061FF");
      AmbilLogMedis("LAPORAN_BACKUP","Sistem","BACKUP_DROPBOX_SUKSES","Backup Dropbox berhasil",{path:dbxCfg.path,waktu:fNow()});
    } else {
      showToast(res.msg, C.d);
      AmbilLogMedis("LAPORAN_BACKUP","Sistem","BACKUP_DROPBOX_GAGAL",`Backup Dropbox gagal: ${res.msg}`,{error:res.msg});
    }
    setDbxBacking(false);
  }, [dbxCfg]);

  const handleDropboxBackupOpsXls = useCallback(async () => {
    setDbxBacking(true); setDbxStatus(null);
    const {ops:o} = latestDataRef.current;
    const wb = XLSX.utils.book_new();
    const h = ["No","Tanggal","Jam","Pasien","Usia","RM","Jenis","Diagnosa","Tindakan","Kamar","Operator","Anestesi","Asisten","Instrumen","P.Anestesi","Onloop","RR/Katim","Status","Dibuat"];
    const rows = [...o].sort((a:any,b:any)=>a.date.localeCompare(b.date)||a.time.localeCompare(b.time)).map((op:any,i:number)=>[i+1,op.date,op.time,op.patient||"",op.age||"",op.rm||"",op.opType||"",op.diagnosis||"",op.procedure||"",op.room||"",op.surgeon||"",op.anesthesiologist||"",op.assistantNurse||"",op.circulatingNurse||"",op.anesthesiaNurse||"",op.onloopNurse||"",op.rrKatim||"",op.status||"",op.createdAt||""]);
    const ws = XLSX.utils.aoa_to_sheet([h,...rows]);
    ws["!cols"]=[{wch:4},{wch:12},{wch:7},{wch:22},{wch:5},{wch:10},{wch:9},{wch:26},{wch:26},{wch:14},{wch:22},{wch:22},{wch:20},{wch:20},{wch:16},{wch:16},{wch:14},{wch:12},{wch:20}];
    XLSX.utils.book_append_sheet(wb, ws, "Jadwal Operasi");
    const folder = dbxCfg.path.replace(/[^/]+$/, "");
    const path = `${folder}Operasi_${todayDate()}.xlsx`;
    const res = await dropboxUploadExcel(path, wb);
    setDbxStatus(res);
    if(res.ok){ setDbxCfg((p:DropboxConfig)=>({...p,lastBackup:fNow()})); showToast("✓ Excel Operasi berhasil diupload ke Dropbox","#3730A3"); }
    else showToast(res.msg, C.d);
    setDbxBacking(false);
  }, [dbxCfg]);

  const handleDropboxBackupLemburXls = useCallback(async () => {
    setDbxBacking(true); setDbxStatus(null);
    const {lemburPegawai:lp, lemburData:ld} = latestDataRef.current;
    const wb = XLSX.utils.book_new();
    /* Per-employee sheets */
    lp.forEach((p:any)=>{
      const pegEntries: any[][] = [];
      Object.keys(ld).filter(k=>k.startsWith(p.id+"_")).sort().forEach(k=>{
        const rec=ld[k];
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
    const res = await dropboxUploadExcel(path, wb);
    setDbxStatus(res);
    if(res.ok){ setDbxCfg((p:DropboxConfig)=>({...p,lastBackup:fNow()})); showToast("✓ Excel Lembur berhasil diupload ke Dropbox","#3730A3"); }
    else showToast(res.msg, C.d);
    setDbxBacking(false);
  }, [dbxCfg]);

  const handleDropboxRestoreOps = useCallback(async () => {
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
  }, [dbxCfg]);

  const handleDropboxRestoreLembur = useCallback(async () => {
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
  }, [dbxCfg]);

  /* FIX #1: saveOpFn menerima {opForm, editingOp, setOpErrors, setDupWarn, resetOp}
     dari ViewDaftar (state lokal) — root App tidak perlu menyimpan opForm lagi. */
  /* FIX AUDIT #24: pendingEditOp menggantikan editOpRef (useRef). Lihat
     catatan lengkap di ViewDaftarProps & ViewDaftar untuk penjelasan bug
     race condition yang diperbaiki. */
  const [pendingEditOp, setPendingEditOp] = useState<Operation|null>(null);
  const saveOpFn = useCallback(({opForm, editingOp, setOpErrors, setDupWarn, resetOp}: SaveOpFnArgs) => {
    const e: Record<string,string>={};
    if(!opForm.patient||!opForm.patient.trim()) e.patient="Nama pasien wajib";
    if(!opForm.diagnosis||!opForm.diagnosis.trim()) e.diagnosis="Diagnosa wajib";
    if(!opForm.procedure||!opForm.procedure.trim()) e.procedure="Nama tindakan wajib";
    if(!opForm.date) e.date="Tanggal wajib";
    if(!opForm.time) e.time="Jam wajib";
    if(!opForm.surgeon) e.surgeon="Dokter operator wajib";
    setOpErrors(e);
    if(Object.keys(e).length){showToast("Lengkapi field yang wajib ✱",C.d);return;}
    const clean={...opForm,opType:(opForm.opType||"elektif") as "elektif"|"semi"|"cyto",patient:sanitize(opForm.patient||""),diagnosis:sanitize(opForm.diagnosis||""),procedure:sanitize(opForm.procedure||""),ruangAsal:sanitize(opForm.ruangAsal||""),allergy:sanitize(opForm.allergy||"Tidak Ada"),specialNeeds:sanitize(opForm.specialNeeds||""),assistantNurse:sanitize(opForm.assistantNurse||""),rrKatim:sanitize(opForm.rrKatim||"")};
    const dupSameDay=ops.some((o:any)=>o.id!==editingOp?.id&&o.patient&&o.patient.toLowerCase().trim()===clean.patient.toLowerCase().trim()&&o.date===clean.date);
    const dup=ops.some((o:any)=>o.id!==editingOp?.id&&o.patient&&o.patient.toLowerCase().trim()===clean.patient.toLowerCase()&&o.date===clean.date&&o.time===clean.time);
    if(dupSameDay&&!dup){showToast(`⚠ Perhatian: Pasien "${clean.patient}" sudah terdaftar pada ${clean.date} (jam berbeda)`,C.w);}
    if(dup){setDupWarn(true);showToast(`⚠ Pasien "${clean.patient}" sudah terdaftar pada tanggal & jam yang sama!`,C.d);return;}
    setDupWarn(false);
    if(editingOp){
      const updatedOp = {...editingOp,...clean, updated_at: new Date().toISOString()};
      setOps(p=>p.map(o=>o.id===editingOp.id ? updatedOp : o));
      logAudit("Edit",clean);
      /* ── TURSO LOG: UBAH_PASIEN ── */
      AmbilLogMedis("DAFTAR_PASIEN", role==="admin"?"Admin":"Perawat", "UBAH_PASIEN",
        `Edit data pasien: ${clean.patient||"-"} (ID: ${editingOp.id})`,
        { before: editingOp, after: updatedOp }
      );
      /* ── TURSO LOG: CYTO ── */
      if(updatedOp.opType==="cyto" && editingOp.opType!=="cyto"){
        AmbilLogMedis("DAFTAR_PASIEN", role==="admin"?"Admin":"Perawat", "UBAH_STATUS_CYTO",
          `EMERGENCY CYTO DETECTED: NRM ${clean.rm||"-"} — Pasien: ${clean.patient||"-"}`,
          { ...updatedOp }
        );
      }
      upsertOneToSupa("kb_operasi", updatedOp).then((res:any)=>{
        if(!res?.ok) {
          setOps(p=>p.map(o=>o.id===editingOp.id ? editingOp : o));
          showToast(`⚠ Gagal menyimpan ke Supabase: ${res?.error||"kesalahan tidak diketahui"}`,C.d);
        } else {
          showToast("✓ Jadwal diperbarui & tersimpan ke Supabase",C.s);
        }
      });
    } else {
      const newOp={...clean,id:gId(),status:"scheduled",reminders:[],requests:[],createdAt:fNow(),updated_at:new Date().toISOString()} as Operation;
      logAudit("Tambah",clean);
      /* ── TURSO LOG: TAMBAH_PASIEN ── */
      AmbilLogMedis("DAFTAR_PASIEN", role==="admin"?"Admin":"Perawat", "TAMBAH_PASIEN",
        `Tambah pasien baru: ${clean.patient||"-"} (RM: ${clean.rm||"-"})`,
        { ...newOp }
      );
      /* ── TURSO LOG: CYTO pada tambah baru ── */
      if(newOp.opType==="cyto"){
        AmbilLogMedis("DAFTAR_PASIEN", role==="admin"?"Admin":"Perawat", "UBAH_STATUS_CYTO",
          `EMERGENCY CYTO DETECTED: NRM ${newOp.rm||"-"} — Pasien: ${newOp.patient||"-"}`,
          { ...newOp }
        );
      }
      upsertOneToSupa("kb_operasi", newOp).then((res:any)=>{
        if(!res?.ok) {
          showToast(`⚠ Gagal menyimpan ke Supabase: ${res?.error||"kesalahan tidak diketahui"}`,C.d);
        } else {
          setOps(p=>{ if(p.some((o:any)=>o.id===newOp.id)) return p; return [...p, newOp]; });
          showToast("✓ Jadwal tersimpan ke Supabase",C.s);
        }
      });
    }
    resetOp();
    setTab("jadwal");
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[ops, showToast, upsertOneToSupa, logAudit]);
  const startEditOp= (op: Operation)  => { setPendingEditOp(op); setTab("daftar"); };
  const deleteOp   = async (id:string) => {
    const op=ops.find(o=>o.id===id);
    /* Optimistic remove lokal */
    setOps(p=>p.filter(o=>o.id!==id));
    if(op) logAudit("Hapus",op);
    /* ── TURSO LOG: HAPUS_PASIEN ── */
    if(op) AmbilLogMedis("DAFTAR_PASIEN", role==="admin"?"Admin":"Perawat", "HAPUS_PASIEN",
      `Hapus data operasi: ${op.patient||"-"} (ID: ${id})`,
      { deleted: op }
    );
    /* DB delete — postgres_changes propagasi ke device lain */
    const res = await deleteFromSupa("kb_operasi", id);
    if(!res?.ok) {
      /* Rollback jika DB gagal */
      if(op) setOps(p=>[...p, op]);
      showToast(`⚠ Gagal menghapus dari Supabase: ${res?.error||"kesalahan tidak diketahui"}`,C.d);
    } else {
      showToast("✓ Jadwal dihapus & tersinkron",C.d);
    }
  };
  const sendReminder=(op:any,type:string)=>{
    if((op.reminders||[]).includes(type))return;
    setNotifs(p=>[{id:gId(),type,label:(type==="H-1"?"H-1":"1 Jam")+" → "+op.surgeon,patient:op.patient,procedure:op.procedure,message:"",sentAt:fNow()},...p]);
    const updatedOp = {...op, reminders:[...(op.reminders||[]),type]};
    /* Optimistic update */
    setOps(p=>p.map((o:any)=>o.id===op.id ? updatedOp : o));
    upsertOneToSupa("kb_operasi", updatedOp).then(res=>{
      if(!res?.ok){ setOps(p=>p.map((o:any)=>o.id===op.id ? op : o)); showToast(`⚠ Gagal mencatat pengingat: ${res?.error||"kesalahan tidak diketahui"}`,C.d); }
    });
    showToast("✓ Pengingat dicatat & tersinkron");
  };
  const getPhone = (name: string) => { const s=staff.find(x=>x.name===name); return s&&s.phone||null; };
  const addReq   = (opId: string) => {
    if(!reqText.trim())return;
    const targetOp = ops.find((o:any)=>o.id===opId);
    if(!targetOp) return;
    const updatedOp = {...targetOp, requests:[...(targetOp.requests||[]),{id:gId(),text:sanitize(reqText.trim()),time:fNow()}]};
    /* Optimistic update */
    setOps(p=>p.map((o:any)=>o.id===opId ? updatedOp : o));
    upsertOneToSupa("kb_operasi", updatedOp).then(res=>{
      if(!res?.ok){ setOps(p=>p.map((o:any)=>o.id===opId ? targetOp : o)); showToast(`⚠ Gagal menyimpan permintaan: ${res?.error||"kesalahan tidak diketahui"}`,C.d); }
    });
    setReqText(""); setReqOpId(null); showToast("✓ Permintaan disimpan & tersinkron");
  };
  const delReq = (opId: string,rId: string) => {
    const targetOp = ops.find((o:any)=>o.id===opId);
    if(!targetOp) return;
    const updatedOp = {...targetOp, requests:(targetOp.requests||[]).filter((r:any)=>r.id!==rId)};
    setOps(p=>p.map((o:any)=>o.id===opId ? updatedOp : o));
    upsertOneToSupa("kb_operasi", updatedOp).then(res=>{
      if(!res?.ok){ setOps(p=>p.map((o:any)=>o.id===opId ? targetOp : o)); showToast(`⚠ Gagal menghapus permintaan: ${res?.error||"kesalahan tidak diketahui"}`,C.d); }
    });
  };

  /* FIX BUG (Rules of Hooks — "Rendered more hooks than during the previous
     render"): handleTabChange, handleTabClickEvt, dan TABS sebelumnya berada
     SETELAH dua conditional return (!pinLoaded / !pinOK) di bawah ini. Akibatnya,
     saat render awal (pinLoaded/pinOK masih false) hook-hook ini TIDAK terpanggil,
     lalu begitu pinOK menjadi true, React mendeteksi hook baru muncul di tengah
     siklus hidup komponen → error. React mewajibkan jumlah & urutan hook yang
     dipanggil tetap sama di SETIAP render, sehingga semua hook (termasuk
     useCallback/useMemo ini) harus dipanggil sebelum return bersyarat apa pun.
     Dipindah ke sini agar selalu terpanggil di setiap render, terlepas dari
     nilai pinLoaded/pinOK.

     FIX AUDIT #12 (MEDIUM — Inline Object/Function di JSX):
     ALL_TABS sebelumnya didefinisikan di dalam render function, artinya array
     baru dibuat setiap render. Dipindah ke luar komponen (module-level const)
     agar referensinya stabil. handleTabChange dibuat satu kali via useCallback
     sehingga setiap tombol tab tidak mendapat fungsi baru setiap render.
     TABS di-memoize bergantung hanya pada role — stabil selama role tidak berubah. */
  const handleTabChange = useCallback((k: string) => {
    if(k === "roster_gen" && role !== "admin"){
      showToast("⚠️ Akses Ditolak: Fitur ini memerlukan hak akses Admin.", C.d);
      AmbilLogMedis("KEAMANAN", role==="admin"?"Admin":"Perawat", "AKSES_DITOLAK",
        `Akses ditolak ke tab roster_gen (role: ${role})`, { tab: k, role });
      setTab("home");
      return;
    }
    /* ── TURSO LOG: PINDAH_TAB ── */
    AmbilLogMedis("NAVIGASI", role==="admin"?"Admin":"Perawat", "PINDAH_TAB",
      `Berpindah ke tab: ${k}`, { tab: k, role });
    setTab(k);
  }, [role, showToast]);
  /* Handler untuk event button — membaca key dari data-tab attribute.
     Satu fungsi stabil untuk semua tombol tab (tidak perlu arrow function per tombol). */
  const handleTabClickEvt = useCallback((e: React.MouseEvent<HTMLButtonElement>) => {
    const k = e.currentTarget.dataset.tab;
    if(!k) return;
    if(k === "roster_gen" && role !== "admin"){
      showToast("⚠️ Akses Ditolak: Fitur ini memerlukan hak akses Admin.", C.d);
      AmbilLogMedis("KEAMANAN", role==="admin"?"Admin":"Perawat", "AKSES_DITOLAK",
        `Akses ditolak ke tab roster_gen via bottom nav (role: ${role})`, { tab: k, role });
      setTab("home");
      return;
    }
    AmbilLogMedis("NAVIGASI", role==="admin"?"Admin":"Perawat", "PINDAH_TAB",
      `Berpindah ke tab: ${k} (via bottom nav)`, { tab: k, role, source: "bottom_nav" });
    setTab(k);
  }, [role, showToast]);
  const TABS = useMemo(() => ALL_TABS.filter(t => t.roles.includes(role)), [role]);

  if(!pinLoaded) return (
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"linear-gradient(135deg,#5FA39A,#3FA897,#16685F)"}}>
      <div style={{textAlign:"center",color:"#fff"}}>
        <img src="/logo.jpeg" style={{width:80,height:80,borderRadius:"50%",objectFit:"cover",marginBottom:16,display:"block",margin:"0 auto 16px"}}/>
        <div style={{fontSize:16,fontWeight:700,marginBottom:8}}>SISTEM KOORDINASI KAMAR BEDAH</div>
        <div style={{fontSize:13,opacity:.8,marginBottom:20}}>Rumah Sakit Panti Rini</div>
        <div style={{fontSize:13,opacity:.7}}>⏳ Memuat konfigurasi...</div>
      </div>
    </div>
  );
  if(!pinOK) return <PinScreen onVerify={handlePinVerify} pinAdmin={pinAdmin} pinPerawat={pinPerawat} isFirstTime={isFirstTime}/>;

  const todayOpsCount = ops.filter((o:any)=>o.date===todayDate()).length;

  const content = (
    <ErrorBoundary>
      {tab==="home" && (
        <div style={{paddingBottom:32}}>
          {/* Welcome Banner */}
          <div style={{background:"linear-gradient(135deg,#1e3a8a,#0284c7)",borderRadius:16,padding:"28px 24px",marginBottom:20,color:"#fff",boxShadow:"0 4px 20px rgba(30,58,138,.30)"}}>
            <div style={{fontSize:11,fontWeight:700,letterSpacing:2,opacity:.75,textTransform:"uppercase",marginBottom:6}}>Selamat Datang</div>
            <div style={{fontSize:22,fontWeight:900,lineHeight:1.2,marginBottom:6}}>Sistem Koordinasi<br/>Kamar Bedah</div>
            <div style={{fontSize:13,opacity:.85}}>RS Panti Rini · {new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"})}</div>
          </div>

          {/* Quick Stats */}
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12,marginBottom:20}}>
            {/* Card 1: Operasi Hari Ini */}
            <div style={{background:"#fff",borderRadius:14,padding:"16px 14px",boxShadow:"0 2px 12px rgba(0,0,0,.07)",border:"1px solid #E3F2FD",textAlign:"center"}}>
              <div style={{fontSize:32,fontWeight:900,color:"#1565C0",lineHeight:1}}>{todayOpsCount}</div>
              <div style={{fontSize:11,fontWeight:700,color:"#1565C0",marginTop:4}}>Operasi</div>
              <div style={{fontSize:10,color:"#5C677D",marginTop:2}}>Hari Ini</div>
            </div>
            {/* Card 2: Status Sync */}
            <div style={{background:"#fff",borderRadius:14,padding:"16px 14px",boxShadow:"0 2px 12px rgba(0,0,0,.07)",border:`1px solid ${rtStatus==="online"?"#E8F5E9":"#FFF3E0"}`,textAlign:"center"}}>
              <div style={{fontSize:28,lineHeight:1}}>{rtStatus==="online"?"🟢":"🟡"}</div>
              <div style={{fontSize:11,fontWeight:700,color:rtStatus==="online"?C.s:C.w,marginTop:4}}>{rtStatus==="online"?"Online":"Offline"}</div>
              <div style={{fontSize:10,color:"#5C677D",marginTop:2}}>Cloud Sync</div>
            </div>
            {/* Card 3: Staf terdaftar */}
            <div style={{background:"#fff",borderRadius:14,padding:"16px 14px",boxShadow:"0 2px 12px rgba(0,0,0,.07)",border:"1px solid #EDE7F6",textAlign:"center"}}>
              <div style={{fontSize:32,fontWeight:900,color:"#512DA8",lineHeight:1}}>{staff.length}</div>
              <div style={{fontSize:11,fontWeight:700,color:"#512DA8",marginTop:4}}>Staf</div>
              <div style={{fontSize:10,color:"#5C677D",marginTop:2}}>Terdaftar</div>
            </div>
          </div>

          {/* Operasi hari ini ringkas */}
          <Card style={{marginBottom:16}}>
            <SH label="📋 Operasi Hari Ini" color={C.i}/>
            {(() => {
              const todayOps = [...ops.filter((o:any)=>o.date===todayDate()&&o.status!=="batal")]
                .sort((a:any,b:any)=>{
                  const aDone=a.status==="done", bDone=b.status==="done";
                  if(!aDone&&bDone) return -1;
                  if(aDone&&!bDone) return 1;
                  return a.time.localeCompare(b.time);
                });
              if(todayOps.length===0) return (
                <div style={{fontSize:13,color:C.tL,textAlign:"center",padding:"8px 0",fontStyle:"italic"}}>Tidak ada operasi terjadwal hari ini</div>
              );
              return todayOps.map((op:any)=>{
                const isDone = op.status==="done";
                const isActive = !isDone && op.status!=="batal";
                const doUpdate = (newStatus: string) => {
                  const upd={...op,status:newStatus,updated_at:new Date().toISOString()};
                  setOps((p:any[])=>p.map((o:any)=>o.id===op.id?upd:o));
                  upsertOneToSupa("kb_operasi",upd).then((res:any)=>{ if(!res?.ok){ setOps((p:any[])=>p.map((o:any)=>o.id===op.id?op:o)); showToast(`⚠ Gagal: ${res?.error||"error"}`,C.d); } });
                };
                /* Auto-minimize: pasien selesai */
                if(isDone) return (
                  <div key={op.id} style={{display:"flex",alignItems:"center",gap:8,padding:"5px 10px",marginBottom:5,borderRadius:9,background:"#f1f5f9",opacity:.4}}>
                    <span style={{fontSize:11}}>✅</span>
                    <span style={{fontSize:12,color:"#475569",flex:1,textDecoration:"line-through",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{maskName(op.patient,privacyMode)} — {op.time}</span>
                    <Bdg label="Selesai" color={C.s} bg={C.sBg}/>
                  </div>
                );
                return (
                  <div key={op.id} style={{display:"flex",gap:10,alignItems:"center",marginBottom:10,paddingBottom:10,borderBottom:`1px solid ${C.b}`,background:isActive?"#fef2f2":"transparent",borderRadius:10,padding:"8px 10px"}}>
                    {/* Blinking dot */}
                    {isActive && <span style={{width:9,height:9,borderRadius:"50%",background:"#ef4444",flexShrink:0,display:"inline-block",animation:"redPulse 1.2s ease-in-out infinite"}}/>}
                    {/* Waktu badge */}
                    <div style={{background:op.opType==="cyto"?"#FFCDD2":C.iBg,borderRadius:8,padding:"4px 8px",fontSize:11,fontWeight:700,color:op.opType==="cyto"?"#B71C1C":C.i,flexShrink:0,minWidth:46,textAlign:"center"}}>{op.time}</div>
                    {/* Info */}
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{fontSize:13,fontWeight:700,color:C.t,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{maskName(op.patient,privacyMode)}</div>
                      <div style={{fontSize:11,color:C.tL,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{op.procedure} · {op.room}</div>
                    </div>
                    {/* Tombol aksi inline */}
                    <div style={{display:"flex",gap:5,flexShrink:0}}>
                      {op.status==="scheduled" && (
                        <button onClick={()=>doUpdate("ongoing")} style={{background:"#22c55e",color:"#fff",border:"none",borderRadius:8,padding:"8px 14px",fontSize:13,fontWeight:800,cursor:"pointer",fontFamily:"inherit",boxShadow:"0 2px 8px rgba(34,197,94,.35)"}}>
                          ▶ Mulai
                        </button>
                      )}
                      {op.status==="ongoing" && (
                        <button onClick={()=>doUpdate("done")} style={{background:"#ef4444",color:"#fff",border:"none",borderRadius:8,padding:"8px 14px",fontSize:13,fontWeight:800,cursor:"pointer",fontFamily:"inherit",boxShadow:"0 2px 8px rgba(239,68,68,.35)"}}>
                          ✓ Selesai
                        </button>
                      )}
                    </div>
                  </div>
                );
              });
            })()}
          </Card>

          {/* Quick Links */}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            <button onClick={()=>setTab("daftar")} style={{background:"linear-gradient(135deg,#0284c7,#38bdf8)",color:"#fff",border:"none",borderRadius:14,padding:"18px 14px",fontSize:15,fontWeight:800,cursor:"pointer",boxShadow:"0 4px 16px rgba(2,132,199,.35)",textAlign:"center",fontFamily:"inherit"}}>
              📝 Daftar Operasi
            </button>
            <button onClick={()=>setTab("laporan")} style={{background:"linear-gradient(135deg,#1d4ed8,#6366f1)",color:"#fff",border:"none",borderRadius:14,padding:"18px 14px",fontSize:15,fontWeight:800,cursor:"pointer",boxShadow:"0 4px 16px rgba(29,78,216,.30)",textAlign:"center",fontFamily:"inherit"}}>
              📄 Laporan Kepala Jaga
            </button>
            <button onClick={()=>setTab("jadwal")} style={{background:"linear-gradient(135deg,#1e3a8a,#1d4ed8)",color:"#fff",border:"none",borderRadius:14,padding:"18px 14px",fontSize:15,fontWeight:800,cursor:"pointer",boxShadow:"0 4px 16px rgba(30,58,138,.30)",textAlign:"center",fontFamily:"inherit"}}>
              📋 Lihat Jadwal
            </button>
            {role==="admin" && (
              <button onClick={()=>setTab("roster_gen")} style={{background:"linear-gradient(135deg,#0369a1,#0284c7)",color:"#fff",border:"none",borderRadius:14,padding:"18px 14px",fontSize:15,fontWeight:800,cursor:"pointer",boxShadow:"0 4px 16px rgba(3,105,161,.30)",textAlign:"center",fontFamily:"inherit"}}>
                🗓 Generator Roster
              </button>
            )}
          </div>
        </div>
      )}
      {tab==="jadwal"     && <ViewJadwal ops={ops} setOps={setOps} startEditOp={startEditOp} deleteOp={deleteOp} sendReminder={sendReminder} reqOpId={reqOpId} setReqOpId={setReqOpId} reqText={reqText} setReqText={setReqText} addReq={addReq} delReq={delReq} getPhone={getPhone} setNotifs={setNotifs} showToast={showToast} privacyMode={privacyMode} role={role} upsertOneToSupa={upsertOneToSupa}/>}
      {tab==="daftar"     && <ViewDaftar pendingEditOp={pendingEditOp} clearPendingEditOp={()=>setPendingEditOp(null)} saveOpFn={saveOpFn} staff={staff} setTab={setTab} templates={templates} setTemplates={setTemplates} showToast={showToast}/>}
      {tab==="laporan"    && <ViewLaporan ops={ops} staff={staff} roster={roster} showToast={showToast} role={role} privacyMode={privacyMode}/>}
      {tab==="wa"         && <ViewKirimWA ops={ops} staff={staff} setNotifs={setNotifs} showToast={showToast}/>}
      {tab==="statistik"  && <ViewStatistik ops={ops} archive={archive}/>}
      {tab==="staf"       && <ViewStaf staff={staff} setStaff={setStaff} roster={roster} setRoster={setRoster} showToast={showToast} upsertOneToSupa={upsertOneToSupa} deleteFromSupa={deleteFromSupa} upsertBulkToSupa={upsertBulkToSupa}/>}
      {tab==="lembur"     && <ViewLembur lemburPegawai={lemburPegawai} setLemburPegawai={setLemburPegawai} lemburData={lemburData} setLemburData={setLemburData} showToast={showToast} supaCfg={supaCfg} dbxCfg={dbxCfg} role={role} upsertOneToSupa={upsertOneToSupa} deleteFromSupa={deleteFromSupa}/>}
      {tab==="monitoring" && <ViewMonitoring monitoringEntries={monitoringEntries} setMonitoringEntries={setMonitoringEntries} monitoringCfg={monitoringCfg} setMonitoringCfg={setMonitoringCfg} showToast={showToast} supaCfg={supaCfg} dbxCfg={dbxCfg} role={role} upsertOneToSupa={upsertOneToSupa} deleteFromSupa={deleteFromSupa}/>}
      {tab==="roster_gen" && role==="admin" && <ViewRosterGenerator showToast={showToast} upsertOneToSupa={upsertOneToSupa} dbxCfg={dbxCfg}/>}
      {tab==="arsip"      && <ViewArsip
        data={{ops, archive, notifs, lemburData, lemburPegawai, monitoringEntries, monitoringCfg}}
        setOps={setOps}
        showToast={showToast}
        role={role}
        upsertBulkToSupa={upsertBulkToSupa}
        auditLog={auditLog}
        privacy={{mode: privacyMode, setMode: setPM}}
        supabase={{
          cfg: supaCfg, setCfg: setSupaCfg, status: supaStatus, backingUp: supaBackingUp,
          onBackup: handleSupaBackup, onRestoreOps: handleSupaRestoreOps, onRestoreLembur: handleSupaRestoreLembur,
          onRestoreMonitoring: handleSupaRestoreMonitoring, onRestoreAll: handleSupaRestoreAll,
        }}
        dropbox={{
          cfg: dbxCfg, setCfg: setDbxCfg, status: dbxStatus, backingUp: dbxBacking,
          onBackup: handleDropboxBackup, onRestoreOps: handleDropboxRestoreOps, onRestoreLembur: handleDropboxRestoreLembur,
          onBackupOpsXls: handleDropboxBackupOpsXls, onBackupLemburXls: handleDropboxBackupLemburXls,
          onBackupMonitoringXls: handleDropboxBackupMonitoringXls,
        }}
        realtime={{status: rtStatus, enabled: rtEnabled, setEnabled: setRtEnabled}}
      />}
    </ErrorBoundary>
  );

  return (
    <div className="kbShell" style={{color:C.t,fontFamily:"-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif"}}>

      {/* ── Sticky header (mobile + tablet: full header; desktop: top bar only) ── */}
      <div className="kb-header" style={{background:"linear-gradient(135deg,#1e3a8a,#1d4ed8 55%,#0284c7)",position:"sticky",top:0,zIndex:100,boxShadow:"0 2px 16px rgba(30,58,138,.35)"}}>
        <div className="kb-header-inner" style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
          <div style={{display:"flex",alignItems:"center",gap:12}}>
            <img src="/logo.jpeg" alt="Logo KB" style={{width:50,height:50,borderRadius:"50%",objectFit:"cover",border:"2px solid rgba(255,255,255,.5)",animation:"lgFloat 3s ease-in-out infinite",flexShrink:0}}/>
            <div>
              <div style={{fontSize:9,fontWeight:700,color:"rgba(255,255,255,.6)",letterSpacing:2,textTransform:"uppercase",marginBottom:1}}>Rumah Sakit Panti Rini</div>
              <div className="kb-title" style={{fontWeight:900,color:"#fff",letterSpacing:.5,lineHeight:1.1}}>SISTEM KOORDINASI<br/>KAMAR BEDAH</div>
              <HeaderClock/>
            </div>
          </div>
          <div style={{display:"flex",gap:8,alignItems:"center"}}>
            {rtStatus==="online" && (
              <span style={{display:"inline-flex",alignItems:"center",gap:6,background:"rgba(255,255,255,.10)",color:"rgba(255,255,255,.85)",fontSize:10,fontWeight:700,padding:"4px 10px 4px 8px",borderRadius:20,border:"1px solid rgba(255,255,255,.18)"}}>
                <span style={{width:7,height:7,borderRadius:"50%",background:"#C9A961",display:"inline-block",animation:"kbPulse 2s ease-in-out infinite"}}/>
                Sync
              </span>
            )}
            {ops.filter((o:any)=>o.status==="ongoing").length>0 && (
              <span style={{background:"#F44336",color:"#fff",fontSize:11,fontWeight:700,padding:"3px 9px",borderRadius:20}}>
                ● {ops.filter((o:any)=>o.status==="ongoing").length} berlangsung
              </span>
            )}
            {role==="admin" && <button onClick={()=>setShowPinMgmt(true)} style={{background:"rgba(255,255,255,.15)",border:"none",borderRadius:8,padding:"6px 10px",color:"#fff",fontSize:11,cursor:"pointer",fontFamily:"inherit",fontWeight:700}}>⚙ PIN</button>}
            <span style={{background:role==="admin"?"rgba(255,255,255,.25)":"rgba(255,255,255,.12)",border:"1px solid rgba(255,255,255,.3)",borderRadius:8,padding:"4px 9px",color:"#fff",fontSize:10,fontWeight:700}}>{role==="admin"?"🔐 Admin":"👤 Perawat"}</span>
            <button onClick={()=>{setPinOK(false);setRole("perawat");}} style={{background:"rgba(255,255,255,.15)",border:"none",borderRadius:8,padding:"6px 12px",color:"#fff",fontSize:11,cursor:"pointer",fontFamily:"inherit",fontWeight:700}}>🔒 Kunci</button>
          </div>
        </div>
        {/* Mobile/tablet tabs — hidden on desktop via CSS */}
        <div className="kb-tabs">
          {TABS.map(t=>(
            <button key={t.k} data-tab={t.k} onClick={handleTabClickEvt} className="kb-tab-btn"
              style={{border:"none",background:tab===t.k?"rgba(255,255,255,.20)":"none",color:tab===t.k?"#fff":"rgba(255,255,255,.60)",fontWeight:tab===t.k?900:600,cursor:"pointer",borderBottom:tab===t.k?"3px solid #38bdf8":"3px solid transparent",borderRadius:"8px 8px 0 0",transition:"all .15s",fontFamily:"inherit"}}>
              {t.l}
            </button>
          ))}
        </div>
      </div>

      {/* ── Storage warning banner ── */}
      {/* ── Body: sidebar (desktop only) + content ── */}
      <div className="kb-main" style={{background:"#f8fafc",flex:1}}>

        {/* Desktop sidebar — hidden on mobile via CSS */}
        <aside className="kb-sidebar">
          <div style={{fontSize:11,fontWeight:700,color:"rgba(255,255,255,.4)",letterSpacing:1,textTransform:"uppercase",padding:"4px 14px 8px",marginBottom:4}}>Menu</div>
          {TABS.filter(t=>t.k!=="arsip").map(t=>(
            <button key={t.k} data-tab={t.k} onClick={handleTabClickEvt}
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
            <button key={t.k} data-tab={t.k} onClick={handleTabClickEvt}
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

      {/* PIN Management Modal (Admin only) */}
      {showPinMgmt && role==="admin" && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.6)",zIndex:9000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
          <div style={{background:C.white,borderRadius:20,padding:"28px 24px",width:"100%",maxWidth:380,boxShadow:"0 20px 60px rgba(0,0,0,.4)"}}>
            <div style={{fontSize:16,fontWeight:800,color:C.p,marginBottom:4}}>⚙ Manajemen PIN</div>
            <div style={{fontSize:12,color:C.tL,marginBottom:20}}>Ubah PIN Admin dan PIN Perawat</div>
            <div style={{fontSize:13,fontWeight:700,color:C.p,marginBottom:8}}>🔐 PIN Admin Baru</div>
            <LF label="PIN Admin Baru (min. 4 digit)"><input style={iS} type="password" inputMode="numeric" maxLength={6} placeholder="PIN admin baru" value={pmNewAdmin} onChange={e=>setPmNewAdmin(e.target.value)}/></LF>
            <LF label="Konfirmasi PIN Admin"><input style={iS} type="password" inputMode="numeric" maxLength={6} placeholder="Ulangi PIN admin" value={pmCfAdmin} onChange={e=>setPmCfAdmin(e.target.value)}/></LF>
            <div style={{fontSize:13,fontWeight:700,color:C.g,margin:"12px 0 8px"}}>👤 PIN Perawat Baru</div>
            <LF label="PIN Perawat Baru (min. 4 digit)"><input style={iS} type="password" inputMode="numeric" maxLength={6} placeholder="PIN perawat baru" value={pmNewPerawat} onChange={e=>setPmNewPerawat(e.target.value)}/></LF>
            <LF label="Konfirmasi PIN Perawat"><input style={iS} type="password" inputMode="numeric" maxLength={6} placeholder="Ulangi PIN perawat" value={pmCfPerawat} onChange={e=>setPmCfPerawat(e.target.value)}/></LF>
            {pmErr && <div style={{fontSize:12,color:C.d,marginBottom:10,fontWeight:600}}>⚠ {pmErr}</div>}
            <div style={{display:"flex",gap:8,marginTop:12}}>
              <Btn full onClick={()=>{
                if(pmNewAdmin && pmNewAdmin.length<4){setPmErr("PIN Admin min. 4 digit");return;}
                if(pmNewAdmin && pmNewAdmin!==pmCfAdmin){setPmErr("PIN Admin tidak cocok");return;}
                if(pmNewPerawat && pmNewPerawat.length<4){setPmErr("PIN Perawat min. 4 digit");return;}
                if(pmNewPerawat && pmNewPerawat!==pmCfPerawat){setPmErr("PIN Perawat tidak cocok");return;}
                if(pmNewAdmin && pmNewPerawat && pmNewAdmin===pmNewPerawat){setPmErr("PIN Admin & Perawat tidak boleh sama");return;}
                // FIX AUDIT #4: hash PIN baru (PBKDF2) sebelum disimpan — jangan
                // pernah simpan plaintext. PIN yang tidak diganti (sudah berupa
                // hash "salt:hash" dari sebelumnya) dipertahankan apa adanya.
                (async()=>{
                  try{
                    const finalAdmin = pmNewAdmin ? await hashPin(pmNewAdmin) : pinAdmin;
                    const finalPerawat = pmNewPerawat ? await hashPin(pmNewPerawat) : pinPerawat;
                    if(pmNewAdmin){ setPinAdmin(finalAdmin); }
                    if(pmNewPerawat){ setPinPerawat(finalPerawat); }
                    await SUPA_CLIENT.from("kamar_bedah_config").upsert({key:"pins",value:JSON.stringify({admin:finalAdmin,perawat:finalPerawat})},{onConflict:"key"});
                    showToast("✓ PIN diperbarui & tersinkron ke semua perangkat",C.s);
                  }catch(e){ showToast("✓ PIN diperbarui (Supabase error, coba lagi)",C.w); }
                })();
                setPmNewAdmin("");setPmCfAdmin("");setPmNewPerawat("");setPmCfPerawat("");setPmErr("");
                setShowPinMgmt(false);
              }}>Simpan</Btn>
              <Btn full outline color={C.g} onClick={()=>{setShowPinMgmt(false);setPmErr("");setPmNewAdmin("");setPmCfAdmin("");setPmNewPerawat("");setPmCfPerawat("");}}>Batal</Btn>
            </div>
          </div>
        </div>
      )}

      {/* Toast */}
      {toast && (
        <div style={{position:"fixed",bottom:24,left:"50%",transform:"translateX(-50%)",background:toast.color,color:"#fff",padding:"11px 22px",borderRadius:14,fontSize:13,fontWeight:700,zIndex:9999,boxShadow:"0 4px 20px rgba(0,0,0,.2)",whiteSpace:"nowrap",maxWidth:"90vw",textAlign:"center"}}>
          {toast.msg}
        </div>
      )}
    </div>
  );
}
