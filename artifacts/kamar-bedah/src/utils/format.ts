/* src/utils/format.ts
   Fungsi utilitas murni untuk formatting & validasi — tanggal/waktu (WIB-
   aware), usia, ID, timeout Promise, sanitasi sel Excel, normalisasi role
   staf & nomor telepon, parsing pesan error Supabase. Dipindah dari
   App.tsx (langkah 3 refactor — pisah utils, batch 2), TIDAK ADA
   perubahan logika, termasuk seluruh FIX AUDIT & FIX BUG di dalamnya. */

import { FORMULA_TRIGGER_CHARS, EXCEL_EPOCH, BATAS_SIANG_MENIT, BATAS_SORE_MENIT, BATAS_MALAM_MENIT, ST } from "../constants";
import { gId, sanitize } from "./security";

export function sanitizeCellForExcel(value: any): any {
  if (typeof value !== "string" || value.length === 0) return value;
  return FORMULA_TRIGGER_CHARS.includes(value[0]) ? `'${value}` : value;
}

/* ─── FIX AUDIT #25 (MEDIUM — Query Tanpa Timeout, "Freeze" saat Network
   Hang) ──────────────────────────────────────────────────────────────
   Tanpa batas waktu eksplisit, sebuah query yang "menggantung" (jaringan
   RS lambat/overload, BUKAN disconnect total) bisa membuat tombol Simpan/
   Masuk/Pulang tetap disabled TANPA BATAS WAKTU — bagi perawat terasa
   seperti aplikasi freeze, tanpa opsi retry. withTimeout() membungkus
   sebuah Promise dan memaksanya reject setelah durasi tertentu, sehingga
   UI selalu punya kepastian (sukses ATAU gagal-dengan-pesan-jelas) dalam
   waktu wajar, bukan menunggu selamanya. */
export function withTimeout<T>(promise: Promise<T>, ms: number, label: string): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(`${label}: waktu tunggu habis (>${ms/1000}s) — periksa koneksi internet`)), ms);
    promise.then(v => { clearTimeout(timer); resolve(v); }, e => { clearTimeout(timer); reject(e); });
  });
}

/* ensureId: pakai id yang sudah ada selama berupa string non-kosong (kolom `id`
   di tabel kb_* bertipe text, BUKAN uuid — dibuktikan oleh row legacy
   id="lembur_data" yang sudah berjalan). Sengaja TIDAK memaksa format UUID di
   sini: beberapa tabel memakai id deterministik (mis. "roster_2026-06-17",
   "mon_2026-06-17_07:00", "${pegId}_${period}") supaya re-upload/re-save dengan
   kunci yang sama melakukan UPSERT (timpa baris yg sama) bukan menumpuk baris
   duplikat baru — memaksa UUID di sini akan menghancurkan mekanisme itu.
   Hanya id yang benar-benar kosong/rusak yang digantikan id baru. */
export const ensureId = (id: any): string => (typeof id === "string" && id.trim().length > 0) ? id : gId();

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
export const toLocalISODate = (d: Date): string => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
export const todayDate = () => toLocalISODate(new Date());
export const tmrwDate  = () => { const d=new Date(); d.setDate(d.getDate()+1); return toLocalISODate(d); };
export const fD        = (d: string) => { if(!d)return"-"; try{ return new Date(d+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}); }catch{ return d; } };
/* fDMY: format DD / MM / YYYY — dipakai di semua tab untuk display tanggal seragam */
export const fDMY      = (d: string): string => { if(!d)return"—"; try{ const [y,m,day]=d.split("-"); if(!y||!m||!day)return d; return `${day.padStart(2,"0")}/${m.padStart(2,"0")}/${y}`; }catch{ return d; } };
/* fNowMY: tanggal+waktu sekarang, tanggal dalam format DD/MM/YYYY */
export const fNow      = () => { const n=new Date(); const day=String(n.getDate()).padStart(2,"0"); const mo=String(n.getMonth()+1).padStart(2,"0"); const y=n.getFullYear(); const hh=String(n.getHours()).padStart(2,"0"); const mm=String(n.getMinutes()).padStart(2,"0"); return `${day}/${mo}/${y} ${hh}:${mm}`; };
export const fTR       = (t: string) => t?t.replace(":",".") : "";
/* FIX AUDIT #11 (MEDIUM): gWord sebelumnya memanggil new Date() dua kali —
   jika dipanggil tepat saat jam berganti (mis. 11:59:59.999 → 12:00:00.000),
   getHours() dan getMinutes() bisa berasal dari momen berbeda dan memberikan
   hasil yang salah. Sekarang menggunakan satu instance Date yang sama. */
/* FIX BUG #35 (audit lanjutan): magic number 660/900/1110 diganti konstanta
   bernama (jam*60+menit) supaya batas pagi/siang/sore/malam langsung
   terbaca tanpa perlu hitung manual — dipakai bersama oleh gWord() &
   timeOfDayLabel() supaya kalau batas waktu ini perlu diubah, cukup di
   satu tempat. */
export const gWord     = (): string => { const now = new Date(); const t = now.getHours() * 60 + now.getMinutes(); return t < BATAS_SIANG_MENIT ? "pagi" : t < BATAS_SORE_MENIT ? "siang" : t < BATAS_MALAM_MENIT ? "sore" : "malam"; };
/* timeOfDayLabel: sama seperti gWord(), tapi berdasarkan jam OPERASI (bukan
   waktu saat ini) — dipakai utk frasa "rencana operasi besok [pagi/siang/sore]"
   di msgSurgeon, yang merujuk ke jam operasinya, bukan jam pengiriman pesan. */
export const timeOfDayLabel = (time: string): string => {
  if(!time) return "";
  const [h,m] = time.split(":").map(Number);
  const t = (h||0)*60 + (m||0);
  return t < BATAS_SIANG_MENIT ? "pagi" : t < BATAS_SORE_MENIT ? "siang" : t < BATAS_MALAM_MENIT ? "sore" : "malam";
};
/* fDLower: varian fD() khusus utk template pesan WA Laporan/msgSurgeon/msgAnest
   sesuai spesifikasi format ("sabtu 20 juni 2026" — lowercase, tanpa koma).
   fD() ASLI TIDAK diubah supaya tempat lain yang sudah memakainya (cetak,
   export, dsb) tidak terdampak/regresi. */
export const fDLower   = (d: string): string => fD(d).replace(",", "").toLowerCase();
/* formatUsia: helper wajib sesuai spesifikasi — dipakai di seluruh modul
   rendering teks Laporan & WA (buildLaporan, msgSurgeon, msgAnest). */
export function formatUsia(tahun: any, bulan: any): string {
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
export const usiaOp    = (o: any): string => (o.age || o.ageMonths) ? formatUsia(o.age, o.ageMonths) : "";
/* titleCaseDokter: khusus utk SAPAAN/header pesan WA (contoh: "dokter Timor")
   — menghapus prefix "dr"/"Dr" jika ada lalu Title Case huruf pertama.
   Nama di BADAN pesan (mis. "dr timor" dalam daftar pasien) TETAP memakai
   nilai asli dari database (hanya di-sanitize, tidak diubah case-nya) —
   sesuai contoh output literal yang diberikan di spesifikasi. */
export const titleCaseDokter = (n: string): string => {
  const c = sanitize(n).replace(/^dr\.?\s*/i, "").trim();
  return c ? c.charAt(0).toUpperCase()+c.slice(1) : c;
};
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
export const toISODateStrict = (value: any): string | null => {
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

export const normalizeStaffType = (raw: string): keyof typeof ST => {
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
export const normalizePhone62 = (raw: string): string => {
  let d = String(raw||"").replace(/[^0-9]/g,"");
  if (d.startsWith("0")) d = "62" + d.slice(1);
  return d;
};
export const isValidPhone62 = (p: string): boolean => /^62\d{8,13}$/.test(p);
/* formatSupaError: ekstrak pesan deskriptif dari PostgrestError Supabase
   (message/details/hint/code) — dipakai agar UI tidak hanya menampilkan
   "gagal, coba lagi" generik melainkan alasan sebenarnya (tipe data salah,
   kolom wajib null, RLS, FK, dsb). */
export const formatSupaError = (error: any): string => {
  if (!error) return "Kesalahan tidak diketahui";
  const parts = [error.message, error.details, error.hint].filter(Boolean);
  const base = parts.length ? parts.join(" — ") : (error.toString?.() || "Kesalahan tidak diketahui");
  return error.code ? `[${error.code}] ${base}` : base;
};
