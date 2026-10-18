import { KOLOM_TAMBAHAN_OPS, nilaiTambahanOps } from "./utils/operasi-excel";
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
import { type RealtimeChannel } from "@supabase/supabase-js";
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ReferenceLine, ResponsiveContainer } from "recharts";

/* ─── ENV TYPES (FIX #1 — AUDIT CRITICAL #1) ────────────────────────────
   Tambahkan ke vite-env.d.ts pada project nyata:
     interface ImportMetaEnv {
       readonly VITE_SUPABASE_URL: string;
       readonly VITE_SUPABASE_ANON_KEY: string;
     }
   Dideklarasikan inline di sini agar file ini tetap valid TypeScript
   tanpa harus menambah file terpisah.

   FIX AUDIT #29 (CRITICAL — Kebocoran Credential Turso di Client Bundle):
   VITE_TURSO_URL / VITE_TURSO_TOKEN SEBELUMNYA dipakai di sini untuk
   membuat @libsql/client langsung di browser. Karena Vite meng-inline
   SEMUA variabel berprefix VITE_ ke dalam bundle publik, token Turso
   bisa diekstrak siapa pun lewat DevTools → Sources, memberi akses
   baca/tulis PENUH ke database audit log. Token ini DIHAPUS TOTAL dari
   sisi client. Audit log sekarang dikirim lewat Supabase Edge Function
   "turso-audit-proxy" (lihat supabase/functions/turso-audit-proxy/index.ts),
   persis seperti pola yang sudah dipakai untuk Dropbox (DBX_PROXY_SECRET).
   Token Turso asli HANYA hidup sebagai secret server-side (Deno.env). */
declare global {
  interface ImportMetaEnv {
    readonly VITE_SUPABASE_URL?: string;
    readonly VITE_SUPABASE_ANON_KEY?: string;
  }
  interface ImportMeta {
    readonly env: ImportMetaEnv;
  }
}

/* ─── CONFIG ────────────────────────────────────────────────────────── */
// CONFIG, HOSPITAL, LOCK_MS sekarang diimpor dari ./constants (lihat import di atas)

/* ─── TURSO AUDIT LOGGER — via Supabase Edge Function Proxy ──────────
   FIX AUDIT #29 (CRITICAL): sebelumnya modul ini membuat @libsql/client
   LANGSUNG di browser dengan token Turso mentah (VITE_TURSO_TOKEN),
   yang berarti token itu ikut ter-bundle ke JS publik dan bisa dicuri
   siapa pun lewat DevTools. Sekarang audit log dikirim ke Supabase Edge
   Function "turso-audit-proxy" — persis pola yang sudah dipakai untuk
   Dropbox (lihat dbxApiRaw/DBX_PROXY_SECRET). Token & URL Turso asli
   HANYA disimpan sebagai secret di sisi server (Deno.env di Edge
   Function), TIDAK PERNAH dikirim ke client.

   Deploy Edge Function berikut sebagai
   supabase/functions/turso-audit-proxy/index.ts (contoh):

     import { createClient as createTursoClient } from "npm:@libsql/client";
     const TURSO = createTursoClient({
       url: Deno.env.get("TURSO_URL")!,
       authToken: Deno.env.get("TURSO_TOKEN")!,
     });
     Deno.serve(async (req) => {
       const { id, waktu, kategori, operator, aksi, detail, metadata } = await req.json();
       await TURSO.execute({
         sql: `INSERT OR IGNORE INTO audit_log_kamar_bedah
                 (id, waktu, kategori, operator, aksi, detail, metadata)
               VALUES (?, ?, ?, ?, ?, ?, ?)`,
         args: [id, waktu, kategori, operator, aksi, detail, JSON.stringify(metadata)],
       });
       return new Response(JSON.stringify({ ok: true }), { headers: { "Content-Type": "application/json" } });
     });

   Jalankan DDL CREATE TABLE IF NOT EXISTS di atas SEKALI lewat Turso CLI
   / migrasi terpisah — bukan lagi otomatis dari client saat modul dimuat.
   ─────────────────────────────────────────────────────────────────── */

// fWIB dan gHex16 sekarang diimpor dari ./utils (lihat import di atas)

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
// AmbilLogMedis sekarang diimpor dari ./services (lihat import di atas)

/* ROOMS, BT, ST, OT, STS, C, EOP, PAGE_SIZE, DEFAULT_RECIPIENT sekarang
   diimpor dari ./constants (lihat import di atas) */

/* ─── DOMAIN INTERFACES (FIX AUDIT #9 — HIGH: TypeScript any masif) ────
   Mendefinisikan kontrak tipe yang eksplisit untuk entitas domain inti.
   Dengan interface ini TypeScript dapat mendeteksi prop hilang, tipe salah,
   dan refactoring yang memecah kontrak — sebelumnya semua terbungkam 'any'. */

/* REFACTOR (langkah 1 — pisah types): semua interface/type yang tadinya
   didefinisikan langsung di file ini (Operation, StaffMember, RosterEntry,
   Notif, ShowToastFn, UpsertOneFn, DeleteFromSupaFn, UpsertBulkFn,
   RtPayload, SupabaseConfig, DropboxConfig, OcrExtractResult, AssistantMsg,
   MonitoringEntry, MonitoringCfg, NurseType, RosterNurse, RosterGenProps,
   NurseMaster, ViewJadwalProps, SaveOpFnArgs, ViewDaftarProps,
   ViewLaporanProps, ViewKirimWAProps, ViewStafProps, ViewArsipProps,
   ScanFormulirProps, OfflineQueueItem, BeforeInstallPromptEvent) sekarang
   dipindah ke src/types/ (dikelompokkan per domain) — TIDAK ADA perubahan
   definisi, murni pindah lokasi supaya App.tsx lebih ramping & types bisa
   dipakai ulang dari file lain nanti (mis. saat komponen View dipisah di
   langkah refactor berikutnya). */
import type {
  Operation, StaffMember, RosterEntry, Notif,
  ShowToastFn, UpsertOneFn, DeleteFromSupaFn, UpsertBulkFn, RtPayload,
  SupabaseConfig, DropboxConfig,
  OcrExtractResult, AssistantMsg,
  MonitoringEntry, MonitoringCfg,
  NurseType, RosterNurse, RosterGenProps, NurseMaster,
  ViewJadwalProps, SaveOpFnArgs, ViewDaftarProps, ViewLaporanProps,
  ViewKirimWAProps, ViewStafProps, ViewArsipProps, ScanFormulirProps,
  OfflineQueueItem, BeforeInstallPromptEvent,
} from "./types";

/* REFACTOR (langkah 2 — pisah constants): semua konstanta module-level yang
   tadinya didefinisikan langsung di file ini sekarang dipindah ke
   src/constants/ (dikelompokkan per domain) — TIDAK ADA perubahan nilai,
   murni pindah lokasi. Konstanta yang TIDAK dipindah (mis. SUPA_CLIENT,
   DBX_PROXY_SECRET) sengaja dibiarkan karena tergolong "services" (dibuat
   dari pemanggilan fungsi runtime seperti createClient()), rencananya
   dipindah di langkah refactor terpisah, bukan langkah constants ini. */
import {
  CONFIG, HOSPITAL, LOCK_MS, ROOMS, BT, ST, OT, STS, EOP, PAGE_SIZE, DEFAULT_RECIPIENT,
  C, iS,
  FORMULA_TRIGGER_CHARS, PIN_MAX_ATTEMPTS, PIN_LOCKOUT_MS, PIN_LOCK_STORAGE_KEY,
  BATAS_SIANG_MENIT, BATAS_SORE_MENIT, BATAS_MALAM_MENIT, EXCEL_EPOCH, BULAN_ID, BULAN_NAMA_ID,
  OCR_SCAN_ENDPOINT, MAX_OCR_BATCH, MAX_OCR_FILE_MB, MAX_OCR_PASTE_CHARS, LS_OFFLINE_QUEUE_KEY,
  MON_JAMS, MON_ROOMS, MON_MONTHS_ID, MC, MON_ROOM_COLORS,
  MASTER_PERAWAT_LIST, SHIFT_CODES, KB_ROSTER_MASTER_TABLE, SEED_MASTER, SUMMARY_COLS, LS_DRAFT_KEY,
  ALL_TABS,
  KB_INSTALL_DISMISS_KEY, KB_INSTALL_DISMISS_MS,
} from "./constants";

/* REFACTOR (langkah 3 — pisah utils, batch 1/beberapa: security & sanitasi):
   fWIB, gHex16, safeJSONParse, hashPin, constantTimeEqual, verifyPin, gId,
   isValidUUID, sanitize, maskName, maskRM, escHtml dipindah ke src/utils/ —
   TIDAK ADA perubahan logika. Fungsi utils lain (format, CSV/Excel,
   OCR-matching, dll) masih di App.tsx, menyusul di batch berikutnya. */
import {
  fWIB, gHex16, safeJSONParse, hashPin, constantTimeEqual, verifyPin,
  gId, isValidUUID, sanitize, maskName, maskRM, escHtml,
  sanitizeCellForExcel, withTimeout, ensureId,
  toLocalISODate, todayDate, tmrwDate, fD, fDMY, fNow, fTR, gWord, timeOfDayLabel, fDLower,
  formatUsia, usiaOp, titleCaseDokter, toISODateStrict,
  normalizeStaffType, normalizePhone62, isValidPhone62, formatSupaError,
  xlsxCell, xlsxSetCell, downloadBlob, sanitizeForCSV, parseCSV,
  coerceOcrResult, capOcrText, normalizeDoctorName, levenshteinRatio,
  looksLikeNonDoctorAnesthesiaValue, findBestStaffMatch,
  validateOcrScalarFields, buildOcrSafeApply,
  compressImageToBase64, rotateImageFile,
  readPinLockState, writePinLockState,
  readOfflineQueue, writeOfflineQueue, enqueueOfflineOp, dequeueOfflineOp, countOfflineQueue,
  msgCyto,
  monIsOK,
} from "./utils";

/* REFACTOR (langkah 4 — pisah services): SUPA_CLIENT, AmbilLogMedis,
   defaultSupaCfg (Supabase); DBX_PROXY_SECRET, DBX_PROXY_FN, defaultDbxCfg,
   callDbxProxy, dropboxUpload, dropboxDownload, dropboxUploadExcel (Dropbox);
   callOcrScanEndpoint, runWorkerOcrScan, runWorkerOcrScanText (OCR) sekarang
   dipindah ke src/services/ — TIDAK ADA perubahan logika. Ini SATU-SATUNYA
   tempat createClient() dipanggil sekarang (di services/supabase.ts). */
import {
  SUPA_URL, SUPA_ANON, SUPA_CLIENT, AmbilLogMedis, defaultSupaCfg,
  DBX_PROXY_SECRET, DBX_PROXY_FN, defaultDbxCfg, callDbxProxy, dropboxUpload, dropboxDownload, dropboxUploadExcel, dropboxUploadBytes,
  callOcrScanEndpoint, runWorkerOcrScan, runWorkerOcrScanText,
} from "./services";

/* REFACTOR (langkah 5 — pisah components, batch 1: effects dekoratif):
   AuroraBackground, MeteorsBackground, GlareHover, Marquee, SplitText,
   TwinklingStars, RotatingText, SquaresBackground, AnimatedListItem,
   MagneticButton, SpotlightCard, DecryptedText, TiltCard, ShinyText,
   fireConfettiBurst, fireClickSpark dipindah ke src/components/effects.tsx
   — TIDAK ADA perubahan logika. Semua komponen ini murni presentational,
   tanpa dependency ke state/types/constants aplikasi. */
import {
  AuroraBackground, MeteorsBackground, GlareHover, Marquee, SplitText,
  TwinklingStars, RotatingText, SquaresBackground, AnimatedListItem,
  MagneticButton, SpotlightCard, DecryptedText, TiltCard, ShinyText,
  fireConfettiBurst, fireClickSpark,
} from "./components/effects";

/* REFACTOR (langkah 5 — pisah components, batch 2: primitives UI dasar):
   LF, TF, TA, SF, Btn, WaBt, Card, ensureSkeletonCss, ensureCustomFontLink,
   SkeletonCard, SkeletonList, AnimatedNumber, HomeParallaxBanner,
   ScrollReveal, Row, Bdg, SH, Toggle, MiniBar, ErrorBoundary dipindah ke
   src/components/primitives.tsx — TIDAK ADA perubahan logika. Ini dipakai
   di HAMPIR SELURUH tab aplikasi, jadi blast radius-nya lebih besar dari
   batch 1 (effects) — sudah diverifikasi diff 100% terhadap kode asli. */
import {
  LF, TF, TA, SF, Btn, WaBt, Card,
  ensureSkeletonCss, ensureCustomFontLink, SkeletonCard, SkeletonList,
  AnimatedNumber, HomeParallaxBanner, ScrollReveal,
  Row, Bdg, SH, Toggle, MiniBar, ErrorBoundary,
} from "./components/primitives";
import { TabIcon, ActionIcon } from "./components/icons";
import { HeaderClock } from "./components/header-clock";
import { InstallPrompt } from "./components/install-prompt";
import { FloatingAssistant } from "./components/floating-assistant";
import { PinScreen } from "./components/pin-screen";
import { DateTxt } from "./components/date-txt";
import { ScanFormulir } from "./components/scan-formulir";
import { ViewKirimWA } from "./components/view-kirim-wa";
import { ViewStatistik } from "./components/view-statistik";
import { StatistikBulanan } from "./components/statistik-bulanan";
import { ViewStaf } from "./components/view-staf";
import { ImportExcel } from "./components/import-excel";
import { ViewArsip } from "./components/view-arsip";
import { ViewJadwal } from "./components/view-jadwal";
import { ViewLaporan } from "./components/view-laporan";
import { ViewLaporanHarian } from "./components/view-laporan-harian";
import { ViewJadwalOK } from "./components/view-jadwal-ok";
// Catatan: __ocrSelfTest tidak perlu diimpor eksplisit — mengimpor nama lain
// dari ocr-matching.ts (file yang sama) sudah memicu window.__ocrSelfTest
// terpasang otomatis saat modul itu dievaluasi.


/* Exception jika data dari cloud terkorupsi (terpotong saat upload, dll).
   Tanpa guard ini seluruh aplikasi bisa Blank Screen sesaat setelah login.
   Helper ini selalu aman: mengembalikan fallback jika parse gagal. */
/* FIX BUG #31 (audit lanjutan): SEBELUMNYA hasil parse dikembalikan apa
   adanya tanpa validasi TIPE — kalau data yang tersimpan ternyata number,
   boolean, atau string mentah (korup/salah simpan), fungsi ini tetap
   mengembalikannya apa adanya, padahal SEMUA pemanggil di file ini
   mengharapkan object/array. guard() menolak primitif (number/boolean/
   string sisa) dan mengembalikan fallback — object & array (bentuk yang
   memang diharapkan pemanggil) tetap lolos apa adanya. */
// safeJSONParse sekarang diimpor dari ./utils (lihat import di atas)

/* ─── PIN HASHING (FIX AUDIT #4 — CRITICAL: PIN Plaintext) ──────────────
   PIN sebelumnya disimpan plaintext di Supabase (kamar_bedah_config).
   Sekarang di-hash dengan PBKDF2-SHA256 (100k iterasi) + salt acak per-PIN
   menggunakan Web Crypto API sebelum dikirim/disimpan. Format string yang
   disimpan: "<salt_hex>:<hash_hex>". Verifikasi PIN dilakukan dengan
   me-re-derive hash dari salt yang tersimpan lalu membandingkan hasilnya —
   PIN asli tidak pernah disimpan maupun dikirim dalam bentuk plaintext
   ke Supabase. */
// hashPin sekarang diimpor dari ./utils (lihat import di atas)
/* verifyPin: bandingkan PIN polos yang diketik user terhadap string
   "<salt>:<hash>" yang tersimpan. Mendukung mode legacy (plaintext lama,
   tanpa format salt:hash) untuk migrasi mulus dari data existing. */
/* FIX AUDIT #16 (CRITICAL — Timing Attack pada verifikasi PIN):
   Sebelumnya `rehashed === stored` memakai perbandingan string biasa, yang
   TIDAK constant-time — waktu eksekusi bisa berbeda marginal tergantung
   berapa karakter awal yang cocok, membuka celah timing-attack teoretis.
   Sekarang memakai constantTimeEqual() yang selalu membandingkan SEMUA
   karakter tanpa early-exit, terlepas dari hasil sebagian. */
/* FIX BUG #37 (audit lanjutan): early-return `if(a.length!==b.length)` di
   versi lama membuat kasus panjang-beda SELALU lebih cepat daripada kasus
   panjang-sama-tapi-isi-beda — perbedaan waktu ini tetap membocorkan info
   (walau minor). Sekarang SELALU loop sampai panjang maksimum, memakai
   0 sebagai pengganti char yang tidak ada di string lebih pendek, supaya
   total waktu eksekusi tidak lagi bergantung pada apakah panjangnya sama. */
// constantTimeEqual sekarang diimpor dari ./utils (lihat import di atas)
/* FIX AUDIT #17 (CRITICAL — Timing Side-Channel pada Fallback Legacy):
   Percabangan `stored.includes(":")` sebelumnya membuat DUA JALUR EKSEKUSI
   dengan biaya komputasi sangat berbeda: PIN ter-hash menjalankan PBKDF2
   100k iterasi (lambat, konstan), sementara PIN legacy plaintext langsung
   dibandingkan tanpa PBKDF2 sama sekali (sangat cepat). Selisih waktu ini
   BISA DIUKUR dari luar (network timing) untuk menyimpulkan apakah sebuah
   akun masih pakai PIN plaintext (target termudah) — bahkan tanpa tahu
   sedikit pun isi PIN-nya. Sekarang KEDUA jalur SELALU menjalankan PBKDF2
   penuh (walau hasilnya dibuang di jalur legacy), sehingga total waktu
   eksekusi kedua cabang seragam dan tidak lagi membocorkan format
   penyimpanan lewat timing. */
// verifyPin sekarang diimpor dari ./utils (lihat import di atas)
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
// gId dan isValidUUID sekarang diimpor dari ./utils (lihat import di atas)

/* ─── FIX AUDIT #18 (MEDIUM — CSV/Excel Formula Injection, CWE-1236) ────
   Field teks bebas dari user (keterangan, keperluan, nama, dsb) yang
   diawali karakter =, +, -, @, tab, atau CR bisa DIINTERPRETASI SEBAGAI
   FORMULA oleh Excel/LibreOffice saat file .xlsx dibuka orang lain —
   berpotensi eksekusi perintah sistem via formula seperti
   =cmd|'/c calc'!A1 atau kebocoran data via =WEBSERVICE(...).
   sanitizeCellForExcel membubuhkan apostrof (') di depan sel yang diawali
   karakter pemicu formula, sehingga Excel selalu memperlakukannya sebagai
   TEKS MURNI, bukan formula — tanpa mengubah tampilan visual sel bagi
   pengguna (apostrof depan tidak tampak di Excel). Dipakai khusus pada
   kolom teks bebas yang diisi manual oleh perawat/admin (keterangan,
   keperluan lembur, nama/paraf TTD, dll), BUKAN pada kolom angka/tanggal
   yang sudah terkontrol formatnya. */
// FORMULA_TRIGGER_CHARS sekarang diimpor dari ./constants (lihat import di atas)
// sanitizeCellForExcel, withTimeout, ensureId sekarang diimpor dari ./utils (lihat import di atas)

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
// toLocalISODate, todayDate, tmrwDate, fD, fDMY, fNow, fTR, gWord,
// timeOfDayLabel, fDLower, formatUsia, usiaOp, titleCaseDokter sekarang
// diimpor dari ./utils (lihat import di atas)
/* FIX BUG #34 (audit lanjutan): strip-tag SATU KALI (`replace(/<[^>]*>/g,"")`
   sekali jalan) secara teoretis bisa lolos untuk pola tag bersarang yang
   dikonstruksi khusus untuk "reformasi" tag setelah 1 pass dihapus. Sekarang
   diulang (loop) sampai tidak ada lagi perubahan, menutup celah itu. Catatan:
   ini pertahanan lapis KEDUA untuk teks bebas (nama, keterangan) — untuk
   HTML yang benar-benar dirender (export Word), pertahanan UTAMA tetap
   escHtml() yang sudah proper-escape sejak FIX AUDIT #3, bukan strip-tag. */
// sanitize, maskName, maskRM sekarang diimpor dari ./utils (lihat import di atas)
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
// EXCEL_EPOCH sekarang diimpor dari ./constants (lihat import di atas)
// toISODateStrict sekarang diimpor dari ./utils (lihat import di atas)
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
// xlsxCell dan xlsxSetCell sekarang diimpor dari ./utils (lihat import di atas)
/* normalizeStaffType: pemetaan label bebas (hasil ketik manual di Excel)
   ke key baku yang dipakai aplikasi (ST). Mencegah staf "hilang" dari
   pengelompokan UI karena typo/label berbeda, walau tersimpan di DB. */
// normalizeStaffType, normalizePhone62, isValidPhone62, formatSupaError sekarang diimpor dari ./utils (lihat import di atas)

/* ─── REAL FILE HELPERS ─────────────────────────────────────────────── */
/* FIX BUG #32 (audit lanjutan): SEBELUMNYA a.click() tidak dibungkus
   try/catch — di sebagian browser (Safari/iOS lama, atau saat popup/
   download diblokir kebijakan browser) pemanggilan ini bisa melempar
   exception yang TIDAK tertangani, membuat elemen <a> & Object URL
   tertinggal (memory leak kecil) dan pengguna tidak tahu unduhan gagal.
   Sekarang dibungkus try/finally: cleanup (hapus elemen, revoke URL)
   SELALU jalan baik sukses maupun gagal, dan kegagalan di-log + dikembalikan
   sebagai boolean supaya pemanggil BISA (opsional) menampilkan pesan error
   ke pengguna alih-alih gagal senyap. */
// downloadBlob sekarang diimpor dari ./utils (lihat import di atas)
/* FIX AUDIT #7 (HIGH — sanitizeForCSV Incomplete / CSV Injection Bypass):
   Implementasi lama hanya menangkap karakter berbahaya di AWAL string.
   Penyerang bisa menyisipkan "\n=cmd|..." di dalam field multi-line —
   Excel memproses baris setelah newline sebagai sel baru dengan formula aktif.
   Fix:
   1. Hapus semua \r\n dari dalam field (normalize ke spasi)
   2. Strip karakter kontrol lain (U+0000–U+001F kecuali tab yang sudah dihandle)
   3. Neutralisasi formula prefix (=, +, -, @, |, \t, \r) di awal hasil bersih
   Dengan urutan ini, injeksi via embedded newline tidak bisa lolos. */
// sanitizeForCSV sekarang diimpor dari ./utils (lihat import di atas)

/* parseCSV: parser CSV minimal tapi AMAN terhadap field yang dibungkus tanda
   kutip (mis. `"Doe, John"` atau keterangan yang mengandung koma) — kasus ini
   sangat umum pada CSV hasil "Save As" dari Excel. Implementasi lama
   (`text.split("\n").map(r=>r.split(","))`) memecah SETIAP koma tanpa
   peduli tanda kutip, sehingga 1 baris bisa pecah menjadi kolom yang salah
   (nama/telepon/jabatan bergeser) TANPA error apa pun yang terlihat —
   silent data corruption saat import staf/roster. Parser di bawah
   menangani: koma di dalam tanda kutip, tanda kutip ganda escaped (""),
   serta akhir baris CRLF/LF. */
// parseCSV sekarang diimpor dari ./utils (lihat import di atas)

/* ─── SUPABASE BACKUP ───────────────────────────────────────────────── */
// SupabaseConfig sekarang diimpor dari ./types (lihat import type di atas)
// defaultSupaCfg sekarang diimpor dari ./services (lihat import di atas)

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
// DBX_PROXY_SECRET, DBX_PROXY_FN, defaultDbxCfg, callDbxProxy, dropboxUpload,
// dropboxDownload, dropboxUploadExcel sekarang diimpor dari ./services (lihat import di atas)

/* ═══════════════════════════════════════════════════════════════════
   FITUR: SCAN FORMULIR OPERASI (Kamera / Screenshot / Galeri) — GRATIS
   ───────────────────────────────────────────────────────────────────
   Menggunakan Tesseract.js (OCR open-source, 100% GRATIS, jalan
   sepenuhnya di BROWSER pasien/perawat — TIDAK ADA API berbayar,
   TIDAK ADA server call, TIDAK ADA biaya per-scan).

   KONSEKUENSI PENTING (perlu dipahami sebelum pakai):
   - Tesseract HANYA membaca teks mentah dari gambar — tidak "paham"
     makna. Pemetaan ke field (nama/diagnosa/tanggal/dll) dilakukan
     lewat PENCARIAN LABEL/KATA KUNCI (regex) di teks hasil OCR,
     BUKAN oleh model AI yang mengerti konteks.
   - Akurasi BAIK untuk teks cetak/ketik dan screenshot digital.
   - Akurasi RENDAH untuk tulisan tangan — bisa gagal total atau
     salah baca huruf/angka. confidence selalu "low" untuk kasus ini.
   - Karena parsing berbasis label, formulir HARUS mengandung kata
     kunci yang dikenali (mis. "Nama:", "Diagnosa:", "Tgl Operasi:")
     agar bisa terpetakan — teks bebas tanpa label tidak akan terbaca
     ke field manapun.
   - Verifikasi manual oleh perawat WAJIB sebelum simpan (sama seperti
     sebelumnya) — anggap ini "bantuan starting point", bukan auto-fill
     yang bisa dipercaya penuh.
   ═══════════════════════════════════════════════════════════════════ */

/* ═══════════════════════════════════════════════════════════════════
   FITUR: SCAN FORMULIR OPERASI (Kamera / Screenshot / Galeri)
   ───────────────────────────────────────────────────────────────────
   Menggunakan CLOUDFLARE WORKERS AI (model vision), dipanggil lewat
   route BARU di Worker yang SAMA dengan yang men-serve App ini
   (lihat berkas worker_ocr_route.ts terpisah untuk kode server-nya).

   GRATIS s.d. 10.000 Neuron/hari (reset tiap 00:00 UTC) — tidak perlu
   API key eksternal apa pun, karena binding `env.AI` sudah otomatis
   tersedia untuk semua Cloudflare Worker. Tidak ada Supabase Edge
   Function tambahan; permintaan cukup ke endpoint sendiri, misalnya
   POST /api/ocr-scan pada domain kamarbedahrsrn.my.id yang sama.

   Foto TIDAK disimpan di mana pun — hanya lewat di memori Worker
   selama request berlangsung untuk diekstrak datanya, lalu dibuang.
   Hasil ekstraksi WAJIB diverifikasi manual oleh perawat sebelum
   disimpan lewat alur saveOpFn yang sudah ada.
   ═══════════════════════════════════════════════════════════════════ */

// OCR_SCAN_ENDPOINT sekarang diimpor dari ./constants (lihat import di atas)

// OcrExtractResult sekarang diimpor dari ./types (lihat import type di atas)

/* FIX AUDIT #34 (HIGH — Validasi Runtime pada Respons Eksternal):
   Sebelumnya `json.data as OcrExtractResult[]` langsung di-cast paksa tanpa
   validasi bentuk/tipe. Kalau model vision/Worker mengembalikan field
   dengan tipe tak terduga (mis. `room` berupa angka), pemanggilan method
   khusus-string (`.match`, `.toUpperCase`) di applyOcrResult/autoSaveOcrBatch
   akan throw dan bisa membuat komponen crash di tengah alur kerja perawat.
   coerceOcrResult() memaksa SEMUA field jadi string yang aman dipakai,
   apa pun tipe asli dari respons server. */
// coerceOcrResult sekarang diimpor dari ./utils (lihat import di atas)

/* FIX AUDIT #35 (MEDIUM — Tidak Ada Batas Panjang Teks Bebas): model vision
   yang berhalusinasi pada gambar buram/kompleks bisa mengembalikan teks
   sangat panjang. capOcrText() memotong ke panjang wajar untuk field
   nama/diagnosa/tindakan/dll sebelum disimpan ke form atau database. */
/* FIX BUG #41 (audit lanjutan): `.slice(0,max)` beroperasi per unit UTF-16,
   bukan per karakter — untuk emoji/karakter di luar BMP (dibentuk dari
   SEPASANG unit UTF-16/"surrogate pair"), potongan bisa jatuh di TENGAH
   pasangan, menghasilkan karakter rusak (mis. tampil sebagai "�"). Array.from
   memecah string per code point (aman terhadap surrogate pair), jadi
   pemotongan tidak akan pernah membelah satu karakter jadi dua. */
// capOcrText sekarang diimpor dari ./utils (lihat import di atas)

/* PENAMBAHAN #1 (fuzzy-match dokter ke Master Data Staf): hasil OCR untuk
   nama dokter sering beda tipis dari yang terdaftar (typo kecil, gelar
   tidak lengkap, dsb — mis. "dr. Adhinanda" hasil OCR vs "dr. Adhinanda,
   Sp.B" yang terdaftar). Fungsi ini mencari nama staf TERDAFTAR yang
   paling mirip dengan hasil OCR, dan HANYA mengganti kalau similarity-nya
   sangat tinggi (>=0.72) — supaya tidak salah cocokkan ke dokter lain yang
   namanya kebetulan agak mirip. Kalau tidak ada yang cukup mirip, nilai
   OCR asli tetap dipakai apa adanya (fail-safe: tidak menebak). */
/* FIX BUG #45 (audit lanjutan): SEBELUMNYA regex `\bsp\.?\s*[a-z]{1,4}\b`
   cuma menangkap 1-4 huruf tepat setelah "sp" — untuk gelar subspesialis
   dengan akhiran tanda kurung/strip (mis. "Sp.B(K)Onk", "Sp.An-KIC"), sisa
   akhirannya ("(k)onk", "-kic") LOLOS tidak terhapus dan ikut masuk ke
   nama yang dinormalisasi, menurunkan akurasi fuzzy-match (skor similarity
   turun karena ada "sampah" karakter tambahan). Sekarang gelar spesialis
   dipotong dari kemunculan pertama "sp" SAMPAI AKHIR STRING — karena gelar
   selalu berada di BELAKANG nama, ini jauh lebih tahan terhadap variasi
   format gelar apa pun (termasuk yang belum pernah dilihat sebelumnya)
   dibanding mencoba menangkap tiap kemungkinan pola gelar satu-satu. */
// normalizeDoctorName sekarang diimpor dari ./utils (lihat import di atas)

// levenshteinRatio sekarang diimpor dari ./utils (lihat import di atas)

/** Cari nama staf paling mirip dari `candidates` (sudah difilter role-nya
 *  di pemanggil). Return nama TERDAFTAR (bentuk resmi) kalau similarity
 *  cukup tinggi, atau `null` kalau tidak ada yang cukup mirip (pemanggil
 *  harus tetap pakai nilai OCR asli dalam kasus ini). */
/** Nilai anestesi yang BUKAN nama dokter (opsi valid di dropdown — lihat
 *  SF label="Dokter Anestesi" di ViewDaftar) — kalau OCR mengembalikan
 *  salah satu dari ini, JANGAN coba fuzzy-match ke nama dokter, karena
 *  ini memang bukan nama orang dan bisa salah "dikoreksi" jadi nama
 *  dokter yang kebetulan mirip hurufnya. */
// looksLikeNonDoctorAnesthesiaValue sekarang diimpor dari ./utils (lihat import di atas)

// findBestStaffMatch sekarang diimpor dari ./utils (lihat import di atas)

/* FIX AUDIT #30 (CRITICAL, bagian OCR): validasi terpusat untuk field
   semi-terstruktur (umur, tanggal, jam, kamar, gol. darah) hasil OCR.
   SEBELUMNYA hanya diterapkan di applyOcrResult (jalur 1-kandidat) — jalur
   autoSaveOcrBatch (≥2 kandidat) menulis field ini MENTAH langsung ke
   Supabase tanpa validasi apa pun. Sekarang kedua jalur memakai validator
   yang SAMA, supaya tidak ada lagi celah data tidak valid yang lolos ke
   database produksi hanya karena lewat jalur auto-save batch. */
// validateOcrScalarFields sekarang diimpor dari ./utils (lihat import di atas)

/* SARAN PENINGKATAN #11 (unit test untuk validasi OCR) — CATATAN JUJUR:
   App.tsx ini adalah satu file komponen React, BUKAN proyek dengan test
   runner (Vitest/Jest) terpasang — menambah test framework sungguhan
   perlu perubahan di luar file ini (package.json, config, folder test
   terpisah) yang tidak bisa dilakukan hanya dengan mengedit App.tsx.
   Sebagai pengganti sementara yang tetap berguna: fungsi ini menjalankan
   beberapa kasus tepi terhadap coerceOcrResult/validateOcrScalarFields/
   buildOcrSafeApply secara manual, dan mencetak hasilnya ke console.
   Cara pakai: buka DevTools Console di browser lalu panggil
   `window.__ocrSelfTest()`. Kalau nanti proyek ini benar-benar dipecah
   jadi beberapa file dengan Vitest, kasus-kasus di bawah ini tinggal
   dipindah jadi `it(...)` beneran. */
// __ocrSelfTest sekarang diimpor dari ./utils (lihat import di atas) — window.__ocrSelfTest otomatis terpasang saat module ./utils di-import

/* SARAN PENINGKATAN #10 + #2 (SATU sumber kebenaran untuk aturan
   kehati-hatian OCR — sebelumnya applyOcrResult & autoSaveOcrBatch
   masing-masing punya logika "jangan auto-percaya CYTO" sendiri-sendiri,
   yang terbukti bisa DRIFT: satu jalur diperbaiki (FIX AUDIT #30 untuk
   batch), satu lagi tertinggal berbulan-bulan (FIX AUDIT #47 untuk
   1-kandidat) sampai ketahuan lewat audit manual. Supaya kelas bug ini
   tidak terulang, SEMUA keputusan "field mana yang butuh perhatian
   ekstra & kenapa" sekarang hidup di SATU fungsi ini. Kalau nanti perlu
   menambah field berisiko lain (mis. nama obat kemoterapi), cukup ubah
   di sini — otomatis berlaku ke jalur 1-kandidat maupun batch.

   Field yang diberi perlakuan berbeda:
   - opType "cyto": TIDAK di-auto-fill sama sekali (beda dari field lain)
     karena statusnya bukan cuma "data salah" tapi memicu EFEK SAMPING
     otomatis di seluruh app (banner darurat, lompat ke urutan teratas
     jadwal untuk SEMUA staf) — risiko salah-baca di sini jauh lebih
     mahal daripada field naratif biasa.
   - bloodType & allergy: TETAP di-auto-fill (beda dari cyto, karena
     keduanya tidak memicu efek samping otomatis apa pun di app — cuma
     data yang tersimpan), TAPI diberi styling "kritis" (merah, sama
     seperti RM & tanggal) supaya tidak tenggelam di antara field
     naratif lain yang risikonya lebih rendah kalau salah baca (nama
     dokter, ruang asal, dll — salah di situ gampang ketahuan & gampang
     diperbaiki; salah gol.darah/alergi baru ketahuan saat tindakan
     medis berlangsung). */
// buildOcrSafeApply sekarang diimpor dari ./utils (lihat import di atas)

/** Kompresi + resize gambar di client sebelum dikirim ke Worker.
 *  Mengecilkan payload — penting untuk jaringan RS yang lambat, dan
 *  mengurangi token/neuron yang dipakai model vision (lebih hemat
 *  kuota harian). */
// compressImageToBase64 sekarang diimpor dari ./utils (lihat import di atas)

/** SARAN PENINGKATAN #7: merotasi gambar 90°/180°/270° via canvas SEBELUM
 *  dikirim ke OCR. Dipakai di langkah preview ScanFormulir supaya foto yang
 *  miring/terbalik (kasus umum saat foto formulir kertas buru-buru) bisa
 *  diluruskan dulu tanpa perlu foto ulang dari awal — akurasi OCR untuk
 *  teks yang terbalik/miring 90° biasanya jauh lebih buruk. */
// rotateImageFile sekarang diimpor dari ./utils (lihat import di atas)
/** Kirim gambar ke route OCR di Worker sendiri (same-origin — tidak
 *  perlu CORS khusus, tidak perlu API key/secret di client karena
 *  request ini berjalan di domain yang sama seperti App di-serve). */
/** Fungsi inti bersama: kirim body apa pun ke /api/ocr-scan lalu urai
 *  responsnya dengan cara yang SAMA PERSIS baik untuk jalur gambar maupun
 *  jalur paste-teks — supaya keduanya tunduk pada validasi & coerce yang
 *  identik (coerceOcrResult dkk) dan tidak bisa drift satu sama lain,
 *  sama semangatnya dengan buildOcrSafeApply() untuk validasi di sisi
 *  form (lihat SARAN PENINGKATAN #10). */
// callOcrScanEndpoint, runWorkerOcrScan, runWorkerOcrScanText sekarang diimpor dari ./services (lihat import di atas)


/* FIX AUDIT #3 (CRITICAL — XSS via downloadAsWord): semua data dinamis
   (judul, header, isi sel) WAJIB di-escape sebelum disisipkan ke string
   HTML. Tanpa ini, data pasien yang mengandung "<script>" atau atribut
   "onerror=" bisa tereksekusi saat file .doc dibuka di Word/LibreOffice
   (stored XSS). */
// escHtml sekarang diimpor dari ./utils (lihat import di atas)

/* Download sebagai Word (.doc via HTML) */
// downloadAsWord sekarang diimpor dari ./utils (lihat import di atas)

/* supabaseBackup & supabaseRestore kini menerima SUPA_CLIENT sebagai parameter
   agar dapat digunakan sebelum modul App selesai di-mount. */
async function supabaseBackup(_cfg: SupabaseConfig, data: any, client: any): Promise<{ok: boolean; msg: string}> {
  try {
    /* Pengaman: jangan simpan backup kalau data lokal masih kosong
       (misal browser/perangkat baru yang datanya belum termuat). */
    const _d: any = data || {};
    if(!(_d.ops?.length) && !(_d.staff?.length) && !(_d.roster?.length) && !(_d.lemburPegawai?.length))
      return {ok:false, msg:"Backup dibatalkan: data lokal masih kosong (belum termuat)"};
    const payload = { created_at: new Date().toISOString(), data: JSON.stringify(data) };
    const {error} = await client.from("kamar_bedah_backup").insert(payload);
    if(error) return {ok:false, msg:"Supabase error: " + error.message.slice(0,120)};
    /* Padatkan: sisakan hanya backup TERBARU per hari (hari WIB).
       Dijalankan SETELAH insert sukses, jadi kalau gagal tidak ada data hilang. */
    try {
      const wib = new Date(Date.now() + 7*3600*1000);
      const awalHariWIB = new Date(Date.UTC(wib.getUTCFullYear(), wib.getUTCMonth(), wib.getUTCDate()) - 7*3600*1000);
      await client.from("kamar_bedah_backup").delete()
        .gte("created_at", awalHariWIB.toISOString())
        .lt("created_at", payload.created_at);
    } catch { /* pemadatan non-kritis */ }
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
// msgSurgeon, msgAnest, msgCyto sekarang diimpor dari ./utils (lihat import di atas)
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
// buildLaporan sekarang diimpor dari ./utils (lihat import di atas)

// LF, TF, TA, SF, Btn, WaBt, Card, ensureSkeletonCss, ensureCustomFontLink,
// SkeletonCard, SkeletonList, AnimatedNumber, HomeParallaxBanner, ScrollReveal,
// Row, Bdg, SH, Toggle, MiniBar, ErrorBoundary sekarang diimpor dari
// ./components/primitives (lihat import di atas)

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
// PIN_MAX_ATTEMPTS, PIN_LOCKOUT_MS sekarang diimpor dari ./constants (lihat import di atas)
/* FIX AUDIT #31 (HIGH — Bypass Lockout via Refresh): sebelumnya `attempts`
   dan `lockedUntil` HANYA berupa useState React biasa, sehingga refresh
   halaman (F5) langsung mereset percobaan ke 0 dan lockout ke null. Ini
   membuat brute-force PIN 4 digit (10.000 kombinasi) jadi trivial — tinggal
   refresh tiap 5 kali gagal. Sekarang state ini di-persist ke sessionStorage
   agar bertahan lewat refresh (mitigasi client-side). CATATAN: ini tetap
   bisa dilewati lewat mode incognito/browser lain — perbaikan PERMANEN
   yang benar adalah enforce rate-limit di server (Supabase Edge Function
   / tabel kb_login_attempts dengan RLS), karena client tidak pernah bisa
   dipercaya sepenuhnya untuk self-enforce security control. */
// PIN_LOCK_STORAGE_KEY sekarang diimpor dari ./constants (lihat import di atas)
// readPinLockState dan writePinLockState sekarang diimpor dari ./utils (lihat import di atas)
/* PENAMBAHAN — efek visual terinspirasi ReactBits.dev (reactbits.dev),
   direkonstruksi pure CSS/JS TANPA menambah dependency baru (tidak perlu
   pnpm install apa pun) — supaya bisa langsung ditempel & jalan. Warna
   disesuaikan dengan palet teal/biru yang sudah dipakai di app.

   1) AuroraBackground — mirip komponen "Aurora" ReactBits: beberapa blob
      gradient besar yang bergerak lambat & saling tumpang tindih, memberi
      kesan cerah & hidup dibanding gradient statis.
   2) ShinyText — teks dengan highlight yang "berjalan" melintas (mirip
      komponen "Shiny Text" ReactBits), dipakai untuk judul.
   3) fireClickSpark — percikan kecil di titik klik tombol (mirip
      komponen "ClickSpark" ReactBits), dipakai saat tombol utama ditekan. */
// AuroraBackground, MeteorsBackground, GlareHover, Marquee, SplitText,
// TwinklingStars, RotatingText, SquaresBackground, AnimatedListItem,
// MagneticButton, SpotlightCard, DecryptedText, TiltCard, ShinyText,
// fireClickSpark sekarang diimpor dari ./components/effects (lihat import di atas)


// PinScreen sekarang diimpor dari ./components/pin-screen (lihat import di atas)


/* ─── VIEW JADWAL ────────────────────────────────────────────────────── */
// ViewJadwal sekarang diimpor dari ./components/view-jadwal (lihat import di atas)


/* ─── DATE TEXT INPUT ────────────────────────────────────────────────── */
/* Menggantikan <input type="date"> di semua tab.
   Alasan: browser Windows dengan locale US menampilkan type="date" dalam
   format MM/DD/YYYY — tidak bisa diubah via CSS/atribut HTML.
   Komponen ini pakai type="text" + auto-insert slash sehingga format
   DD/MM/YYYY selalu konsisten di semua OS/browser.
   Nilai internal tetap YYYY-MM-DD (kompatibel DB & perbandingan string). */
// DateTxt sekarang diimpor dari ./components/date-txt (lihat import di atas)

/* ─── SCAN FORMULIR OPERASI (Kamera / Screenshot / Galeri) ──────────────
   Komponen tombol yang dipakai di dalam ViewDaftar. Mengirim foto ke
   Edge Function "ocr-proxy", menerima array hasil ekstraksi (bisa >1
   entri jika 1 screenshot berisi banyak pasien), lalu meneruskan ke
   callback onResults milik ViewDaftar untuk auto-fill form. ────────── */
// ScanFormulirProps sekarang diimpor dari ./types (lihat import type di atas)
/* FIX AUDIT #36 (MEDIUM — Resource Exhaustion): batas jumlah kandidat yang
   diproses dari 1 scan, mencegah auto-save massal kalau server/model
   mengembalikan array besar akibat bug atau respons cacat. */
// MAX_OCR_BATCH, MAX_OCR_FILE_MB, MAX_OCR_PASTE_CHARS sekarang diimpor dari ./constants (lihat import di atas)

/* ════════════════════════════════════════════════════════════════════
   PENAMBAHAN #8 — Offline queue untuk entri pasien (kb_operasi)
   ════════════════════════════════════════════════════════════════════
   Modul roster jadwal (lihat syncToCloudSilent di ViewLembur/generator
   roster) SUDAH punya pola offline-first: simpan ke localStorage dulu,
   baru sync ke Supabase, dengan auto-retry saat koneksi kembali. Modul
   input pasien (ViewDaftar → saveOpFn/autoSaveOcrBatch) BELUM — kalau
   offline atau upsert gagal karena network, data pasien HANYA ada di
   state React (hilang kalau tab ditutup/di-refresh sebelum sempat sync).

   Utility di bawah ini menambal celah itu: setiap upsert kb_operasi yang
   gagal (offline atau error network) otomatis masuk antrean localStorage
   yang PERSISTEN (bertahan meski tab ditutup), dan dicoba lagi otomatis
   begitu browser mendeteksi koneksi kembali online — pola yang sama
   persis semangatnya dengan yang sudah dipakai di roster generator. */
// LS_OFFLINE_QUEUE_KEY sekarang diimpor dari ./constants (lihat import di atas)

// OfflineQueueItem sekarang diimpor dari ./types (lihat import type di atas)

// readOfflineQueue, writeOfflineQueue, enqueueOfflineOp, dequeueOfflineOp,
// countOfflineQueue sekarang diimpor dari ./utils (lihat import di atas)

/* PENAMBAHAN — Live Camera Scan: buka kamera live (getUserMedia) langsung
   di browser, tampilkan preview real-time, perawat arahkan ke formulir
   lalu tekan tombol "Ambil" saat sudah pas fokusnya — TIDAK auto-capture
   otomatis (sengaja): deteksi "gambar sudah fokus/stabil" itu kompleks dan
   rawan salah jepret di kondisi cahaya kamar bedah yang bervariasi: lebih
   aman & dapat diandalkan kalau perawat sendiri yang menentukan kapan
   gambar sudah pas, baru menekan tombol.

   Frame yang di-capture diubah jadi `File` lalu diteruskan ke
   `onCapture(file)` — pemanggil (ScanFormulir) meneruskannya ke
   `onFilePicked` yang SAMA PERSIS dipakai jalur "Ambil Foto"/"Upload
   Screenshot" — jadi TIDAK ADA duplikasi logika kompresi/preview/rotate/
   kirim-ke-OCR sama sekali, cuma sumber gambarnya saja yang berbeda. */
// LiveCameraScanModal dan ScanFormulir sekarang diimpor dari
// ./components/scan-formulir (lihat import di atas)


/* ─── VIEW DAFTAR ────────────────────────────────────────────────────── */
/* FIX #1: opForm state dipindah ke dalam ViewDaftar sebagai local state.
   Sebelumnya state ini ada di root App, sehingga setiap ketukan huruf di
   kolom "Nama Pasien" memicu re-render seluruh aplikasi — sangat berat di
   tablet/HP. Sekarang hanya ViewDaftar yang re-render saat mengetik.
   pendingEditOp (prop, lihat FIX AUDIT #24) memungkinkan root App (via
   startEditOp) mengisi form lokal ini — menggantikan mekanisme editOpRef
   lama yang rentan race condition saat ViewDaftar belum/baru mount. */
function ViewDaftar({pendingEditOp,clearPendingEditOp,saveOpFn,autoSaveOcrBatch,staff,setTab,templates,setTemplates,showToast}: ViewDaftarProps) {
  // ── Form state: lokal di sini, tidak lagi di root App ──
  // FIX AUDIT #24: initial state langsung diisi dari pendingEditOp kalau ada
  // (penting untuk RENDER PERTAMA — tidak perlu menunggu useEffect sama sekali
  // untuk kasus edit yang dibuka dari tab lain, karena component baru mount).
  const [opForm,   setOpForm]   = useState<Partial<Operation>>(()=> pendingEditOp ? {...pendingEditOp} : {...EOP, date:todayDate()});
  const [editingOp,setEditingOp]= useState<any>(pendingEditOp || null);
  const [opErrors, setOpErrors] = useState<any>({});
  const [dupWarn,  setDupWarn]  = useState(false);

  /* SARAN PENINGKATAN #5 (draft OCR tersimpan lokal): kalau tab/app force-
     close atau reload di tengah alur "scan → periksa → simpan" (mis. HP
     perawat kehabisan baterai, atau app ke-refresh tidak sengaja), tanpa
     ini seluruh hasil scan HILANG dan formulir harus difoto ulang dari
     nol. localStorage AMAN dipakai di sini (ini aplikasi web sungguhan
     yang di-serve dari domain sendiri, BUKAN Claude Artifact) — datanya
     murni draft form biasa, bukan kredensial. Draft disimpan HANYA saat
     ada field yang berasal dari OCR & belum disimpan (ocrFilledFields.size
     >0), dan SELALU dihapus begitu form berhasil disimpan/dibatalkan. */
  const OCR_DRAFT_KEY = "kb_ocr_draft_v1";
  const [showDraftRecovery, setShowDraftRecovery] = useState(false);
  const [draftPreview, setDraftPreview] = useState<any>(null);

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

  // SARAN #5: cek draft tersisa saat komponen baru mount (sekali saja) —
  // hanya tawarkan pemulihan kalau TIDAK sedang membuka mode Edit (supaya
  // tidak menimpa data pasien yang memang sedang diedit dari tab lain).
  useEffect(()=>{
    if(pendingEditOp) return;
    try{
      const raw = localStorage.getItem(OCR_DRAFT_KEY);
      if(!raw) return;
      const parsed = safeJSONParse(raw, null);
      if(parsed && parsed.opForm && Array.isArray(parsed.ocrFilledFields) && parsed.ocrFilledFields.length>0){
        setDraftPreview(parsed);
        setShowDraftRecovery(true);
      }
    }catch{ /* localStorage bisa saja diblokir (mode privat dsb) — abaikan, fitur ini best-effort */ }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  },[]);

  const restoreDraft = () => {
    if(!draftPreview) return;
    setOpForm(draftPreview.opForm);
    setOcrFilledFields(new Set(draftPreview.ocrFilledFields||[]));
    setOcrCriticalFields(new Set(draftPreview.ocrCriticalFields||["rm","date"]));
    setOcrConfidence(draftPreview.ocrConfidence||null);
    setOcrOriginalValues(draftPreview.ocrOriginalValues||{});
    setShowDraftRecovery(false);
    showToast("✓ Draft OCR sebelumnya dipulihkan — periksa lagi sebelum simpan.", C.s);
  };
  const discardDraft = () => {
    try{ localStorage.removeItem(OCR_DRAFT_KEY); }catch{}
    setShowDraftRecovery(false);
    setDraftPreview(null);
  };

  /* FIX AUDIT #39 (lanjutan) + SARAN #5: reset SEMUA state terkait OCR
     (bukan cuma ocrFilledFields) & hapus draft localStorage — tanpa ini,
     banner "form terisi dari OCR, BELUM TERSIMPAN" atau draft lama bisa
     nyangkut/salah tampil lagi setelah form BERHASIL disimpan & dikosongkan
     untuk entri berikutnya. */
  const resetOp = () => {
    setOpForm({...EOP,date:todayDate()}); setEditingOp(null); setOpErrors({}); setDupWarn(false);
    setOcrFilledFields(new Set()); setOcrCriticalFields(new Set(["rm","date"])); setOcrConfidence(null); setOcrOriginalValues({});
    try{ localStorage.removeItem(OCR_DRAFT_KEY); }catch{}
  };
  const saveOp  = () => saveOpFn({opForm, editingOp, setOpErrors, setDupWarn, resetOp, ocrOriginalValues});
  const dupWarning = dupWarn;

  const byT = (t: string) => staff.filter((s:any)=>s.type===t).map((s:any)=>s.name);
  const allNurses = [...new Set([...byT("circulating"),...byT("anesthesia_nurse"),...byT("onloop"),...byT("katim")])];
  const [foc,setFoc] = useState<string|null>(null);
  const [showTpl,setShowTpl] = useState(false);
  /* SARAN PENINGKATAN #8 (bedakan "field terisi OCR belum diverifikasi" vs
     "sudah diedit/dikonfirmasi user"): SEBELUMNYA highlight kuning/merah
     hanya hilang saat form disimpan atau dikosongkan — kalau perawat
     mengedit satu field yang tadinya di-highlight, highlight itu TETAP
     nempel padahal perawat sudah jelas-jelas melihat & mengoreksinya.
     Sekarang set()/setR() (dipakai SEMUA field form, lihat pemakaiannya
     di JSX di bawah) otomatis mencopot field itu dari ocrFilledFields
     begitu user mengetik/memilih ulang — mengedit dianggap sebagai
     tindakan konfirmasi eksplisit, sama seperti Edit+Simpan yang
     membersihkan needsReview di FIX AUDIT #45. */
  const clearOcrHighlight = (k: string) => setOcrFilledFields((prev: Set<string>) => {
    if(!prev.has(k)) return prev;
    const next = new Set(prev); next.delete(k); return next;
  });
  /* FIX BUG (SARAN #5 — ditemukan saat audit lanjutan): banner "Ada draft
     OCR belum tersimpan" SEBELUMNYA hanya ditutup lewat tombol Pulihkan/
     Abaikan eksplisit. Kalau perawat MENGABAIKAN banner itu begitu saja
     lalu langsung mulai scan/isi form BARU, banner tetap nempel di layar
     dengan data draft LAMA yang sudah tidak relevan — kalau perawat lalu
     (tanpa sengaja atau reflex) menekan "Pulihkan Draft", form yang BARU
     SAJA diisi (mis. hasil scan pasien lain) akan TERTIMPA draft lama,
     data yang baru dikerjakan hilang senyap. Sekarang: begitu user
     berinteraksi dengan field apa pun (lewat set/setR, dipakai semua
     field form), banner otomatis dianggap tidak relevan lagi & ditutup. */
  const dismissStaleDraftBanner = () => setShowDraftRecovery(prev => prev ? false : prev);
  const set  = (k: string,v: any) => { setOpForm((p: any)=>({...p,[k]:v})); setOpErrors((p: any)=>({...p,[k]:undefined})); clearOcrHighlight(k); dismissStaleDraftBanner(); };
  const setR = (k: string,v: any) => { setOpForm((p: any)=>({...p,[k]:v})); setOpErrors((p: any)=>({...p,[k]:undefined})); clearOcrHighlight(k); dismissStaleDraftBanner(); };
  const e = opErrors;

  // ── Scan Formulir (OCR): state kandidat hasil scan + field yang ter-auto-fill ──
  // ocrFilledFields dipakai untuk menandai border kuning/merah pada field yang
  // terisi otomatis dari scan, supaya perawat tahu field mana yang wajib dicek
  // ulang sebelum menyimpan. Field "kritis" (merah) sekarang DINAMIS lewat
  // buildOcrSafeApply() — dasarnya selalu RM & tanggal, ditambah golongan
  // darah+alergi kalau hasil scan memang mengisi keduanya (saran #2).
  const [ocrFilledFields, setOcrFilledFields] = useState<Set<string>>(new Set());
  const [ocrCriticalFields, setOcrCriticalFields] = useState<Set<string>>(new Set(["rm","date"]));
  /* SARAN PENINGKATAN #1: field `confidence` dari model ("high"/"low")
     SEBELUMNYA di-parsing tapi tidak pernah dipakai di mana pun — murni
     dibuang. Sekarang dipakai untuk mewarnai highlight lebih tegas
     (oranye lebih pekat + border lebih tebal) saat model sendiri
     melaporkan tidak yakin dengan bacaannya — sinyal ini sudah ada dari
     awal, cuma belum dimanfaatkan. */
  const [ocrConfidence, setOcrConfidence] = useState<"high"|"low"|null>(null);
  /* SARAN PENINGKATAN #3 (audit trail OCR): menyimpan nilai ASLI hasil OCR
     (sebelum sempat diedit user) untuk field yang ter-auto-fill, supaya
     saat form disimpan, saveOpFn bisa membandingkan nilai akhir vs nilai
     asli OCR dan mencatat field mana yang dikoreksi user vs dipakai apa
     adanya. Log ini dikirim lewat AmbilLogMedis (infrastruktur audit yang
     sudah ada di app, dipakai di 30+ titik lain) supaya kalau suatu saat
     ada laporan data pasien salah, bisa ditelusuri apakah itu salah baca
     OCR yang luput dikoreksi, atau salah input manual dari awal. */
  const [ocrOriginalValues, setOcrOriginalValues] = useState<Record<string,string>>({});

  // SARAN #5 (lanjutan): auto-simpan draft ke localStorage setiap kali ada
  // perubahan SELAMA form masih berisi data OCR yang belum tersimpan.
  // Debounce ringan (400ms) supaya tidak menulis localStorage di setiap
  // ketukan huruf saat perawat mengoreksi field.
  useEffect(()=>{
    if(ocrFilledFields.size===0) return;
    const t = setTimeout(()=>{
      try{
        localStorage.setItem(OCR_DRAFT_KEY, JSON.stringify({
          opForm, ocrFilledFields: Array.from(ocrFilledFields), ocrCriticalFields: Array.from(ocrCriticalFields),
          ocrConfidence, ocrOriginalValues, savedAt: new Date().toISOString(),
        }));
      }catch{ /* localStorage penuh/diblokir — draft cuma best-effort, jangan sampai mengganggu alur utama */ }
    }, 400);
    return ()=>clearTimeout(t);
  },[opForm, ocrFilledFields, ocrCriticalFields, ocrConfidence, ocrOriginalValues]);

  const applyOcrResult = (r: OcrExtractResult) => {
    const filled = new Set<string>();
    // FIX BUG (lanjutan, lihat komentar lengkap di dismissStaleDraftBanner):
    // hasil scan BARU selalu menggantikan relevansi banner draft LAMA.
    setShowDraftRecovery(false);
    /* FIX AUDIT #30/#35 (konsistensi) + SARAN #10: pakai validator DAN
       aturan kehati-hatian yang SAMA PERSIS dengan autoSaveOcrBatch lewat
       buildOcrSafeApply() — satu sumber kebenaran, lihat komentar di
       definisinya untuk alasan lengkap kenapa ini penting. */
    const safe = buildOcrSafeApply(r);
    const v = safe.scalars;
    let dateWasInvalid = !!r.date && !v.date;
    const suspectedCyto = safe.suspectedCyto;
    const originalVals: Record<string,string> = {};
    setOpForm((prev: any) => {
      const next = {...prev};
      if(r.patient)         { next.patient = capOcrText(r.patient); filled.add("patient"); originalVals.patient = next.patient; }
      if(v.age)             { next.age = v.age; filled.add("age"); originalVals.age = v.age; }
      if(v.ageMonths)       { next.ageMonths = v.ageMonths; filled.add("ageMonths"); originalVals.ageMonths = v.ageMonths; }
      if(r.rm)              { next.rm = capOcrText(r.rm, 50); filled.add("rm"); originalVals.rm = next.rm; }
      if(r.diagnosis)       { next.diagnosis = capOcrText(r.diagnosis); filled.add("diagnosis"); originalVals.diagnosis = next.diagnosis; }
      if(r.procedure)       { next.procedure = capOcrText(r.procedure); filled.add("procedure"); originalVals.procedure = next.procedure; }
      if(r.opType && !suspectedCyto) { next.opType = r.opType; filled.add("opType"); originalVals.opType = r.opType; }
      // FIX BUG: field "Kamar Operasi" adalah dropdown dengan opsi tetap
      // (ROOMS = ["Kamar Operasi 1"..."5"]). Teks bebas hasil OCR (mis.
      // "OK 2" atau "Ruang Bedah 2") tidak akan cocok dengan opsi manapun
      // kalau ditaruh mentah — <select> akan tampak kosong/tidak terpilih
      // secara visual walau state sebenarnya berisi string yang salah.
      if(v.room)            { next.room = v.room; filled.add("room"); originalVals.room = v.room; }
      // FIX BUG (CRITICAL): tanggal SEBELUMNYA masuk mentah dari OCR tanpa
      // lewat toISODateStrict() — validator yang sama dipakai DateTxt di
      // seluruh app. Kalau tidak valid, field DIKOSONGKAN dan diberi tanda
      // merah agar perawat WAJIB isi manual.
      if(v.date)            { next.date = v.date; filled.add("date"); originalVals.date = v.date; }
      else if(r.date)        { next.date = ""; }
      if(v.time)            { next.time = v.time; filled.add("time"); originalVals.time = v.time; }
      if(r.surgeon)         {
        const rawName = capOcrText(r.surgeon, 80);
        const matched = findBestStaffMatch(rawName, byT("surgeon"));
        next.surgeon = matched || rawName;
        filled.add("surgeon"); originalVals.surgeon = next.surgeon;
      }
      if(r.anesthesiologist){
        const rawName = capOcrText(r.anesthesiologist, 80);
        // FIX BUG: "Lokal Anestesi" BUKAN nama dokter (opsi valid di
        // dropdown) — jangan coba fuzzy-match, bisa salah "dikoreksi"
        // jadi nama dokter yang kebetulan mirip hurufnya.
        const matched = looksLikeNonDoctorAnesthesiaValue(rawName) ? null : findBestStaffMatch(rawName, byT("anesthesiologist"));
        next.anesthesiologist = matched || rawName;
        filled.add("anesthesiologist"); originalVals.anesthesiologist = next.anesthesiologist;
      }
      if(r.allergy)         { next.allergy = capOcrText(r.allergy, 120); filled.add("allergy"); originalVals.allergy = next.allergy; }
      // FIX BUG: "Golongan Darah" juga dropdown dengan opsi tetap (BT).
      if(v.bloodType)       { next.bloodType = v.bloodType; filled.add("bloodType"); originalVals.bloodType = v.bloodType; }
      if(safe.cautionNote){
        /* Cegah catatan dobel kalau perawat scan ulang formulir yang sama
           (mis. foto pertama buram lalu difoto ulang) — tanpa guard ini,
           setiap scan berulang menumpuk baris catatan yang identik. */
        if(!next.specialNeeds || !next.specialNeeds.includes(safe.cautionNote)){
          next.specialNeeds = next.specialNeeds ? `${safe.cautionNote}\n${next.specialNeeds}` : safe.cautionNote;
        }
        filled.add("specialNeeds");
      }
      return next;
    });
    setOcrFilledFields(filled);
    setOcrCriticalFields(new Set(safe.criticalFieldKeys));
    setOcrConfidence(r.confidence);
    setOcrOriginalValues(originalVals);

    setOpErrors({});
    /* FIX AUDIT #39 (BUG UTAMA yang dilaporkan user): SEBELUMNYA fungsi ini
       hanya menampilkan toast kalau tanggal invalid — untuk kasus normal
       (semua field valid) TIDAK ADA toast dari applyOcrResult sama sekali,
       sehingga toast hijau generik "✓ ... data terbaca" dari ScanFormulir
       (yang terlihat seperti "berhasil disimpan") jadi satu-satunya pesan
       yang perawat lihat — padahal data BARU mengisi form di tab "Daftar",
       BELUM tersimpan ke Supabase dan BELUM masuk tab "Jadwal". Ini akar
       penyebab keluhan "notifikasi berhasil tapi tidak muncul di Jadwal".
       Sekarang applyOcrResult SELALU menampilkan toast oranye (bukan hijau)
       yang eksplisit bilang data BELUM tersimpan, apa pun hasil validasinya. */
    if(suspectedCyto){
      showToast("🔴 OCR mendeteksi kemungkinan CYTO/EMERGENCY — TIDAK dipilih otomatis. Periksa formulir asli & pilih manual di \"Jenis Operasi\" jika benar, lalu \"Simpan\".", C.d);
    } else if(dateWasInvalid){
      showToast("⚠ Tanggal hasil scan tidak valid — mohon isi tanggal secara manual, lalu tekan \"Simpan\"", C.d);
    } else if(r.confidence==="low"){
      /* FIX BUG (audit menyeluruh): SEBELUMNYA lowConfidenceFields sudah
         dihitung & dikirim server, tapi TIDAK PERNAH sampai dipakai di
         UI — perawat cuma dapat pesan generik "periksa teliti SEMUA
         field" tanpa tahu field MANA yang sebenarnya perlu dicek. Sekarang
         sebutkan field spesifiknya (kalau tersedia), pakai label yang
         mudah dipahami perawat — bukan nama key teknis seperti "date". */
      const fieldLabels: Record<string,string> = {
        patient:"Nama Pasien", age:"Usia (Tahun)", ageMonths:"Usia (Bulan)", rm:"No. RM",
        diagnosis:"Diagnosa", procedure:"Tindakan", opType:"Jenis Operasi",
        date:"Tanggal", time:"Jam", surgeon:"Dokter Operator",
        anesthesiologist:"Dokter Anestesi", room:"Kamar Operasi",
        allergy:"Riwayat Alergi", bloodType:"Golongan Darah",
      };
      const fieldNames = (r.lowConfidenceFields||[]).map(f=>fieldLabels[f]||f).filter(Boolean);
      const fieldHint = fieldNames.length>0
        ? ` Terutama: ${fieldNames.join(", ")}.`
        : "";
      showToast(`📝 Data terbaca (keyakinan AI rendah — tulisan/teks kurang jelas).${fieldHint} PERIKSA TELITI sebelum "Simpan".`, C.w);
    } else {
      showToast("📝 Data terbaca & form terisi otomatis — GULIR KE BAWAH, periksa, lalu tekan \"Simpan\" agar masuk ke Jadwal.", C.w);
    }
  };

  const iStyle = (k: string): React.CSSProperties => {
    if(ocrFilledFields.has(k) && !e[k]){
      const critical = ocrCriticalFields.has(k);
      const lowConf = ocrConfidence==="low";
      // SARAN #1: confidence rendah → warna lebih tegas & border lebih
      // tebal, supaya field yang model-nya sendiri kurang yakin lebih
      // mencolok dibanding field auto-fill biasa.
      const color = critical ? C.d : (lowConf ? "#EA580C" : "#F59E0B");
      const bg = critical ? "#FEF2F2" : (lowConf ? "#FFF1E6" : "#FFFBEB");
      return {...iS, borderColor: color, borderWidth: (critical||lowConf) ? 2 : 1, background: bg, boxShadow: foc===k?`0 0 0 3px ${color}22`:"none"};
    }
    return {...iS, borderColor:e[k]?C.d:foc===k?C.p:C.b, boxShadow:foc===k?`0 0 0 3px ${C.p}22`:e[k]?`0 0 0 3px ${C.d}22`:"none"};
  };
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
        <div><div key={editingOp?"edit":"new"} style={{fontSize:17,fontWeight:800,color:C.t}}><SplitText text={editingOp?"✏️ Edit Jadwal":"📝 Pendaftaran Jadwal Operasi"}/></div><div style={{fontSize:12,color:C.tL,marginTop:2}}>Field bertanda <span style={{color:C.d,fontWeight:700}}>✱</span> wajib diisi</div></div>
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
      {/* SARAN PENINGKATAN #5: banner pemulihan draft OCR — tampil kalau ada
          draft tersimpan dari sesi sebelumnya yang belum sempat disimpan
          (mis. app sempat ke-refresh/force-close di tengah proses periksa). */}
      {showDraftRecovery && draftPreview && (
        <div style={{background:"#EFF6FF",border:`1.5px solid ${C.i}`,borderRadius:10,padding:"10px 14px",marginBottom:14,fontSize:13,color:C.i}}>
          <div style={{fontWeight:700,marginBottom:6}}>💾 Ada draft hasil scan OCR yang belum tersimpan dari sesi sebelumnya</div>
          <div style={{marginBottom:8}}>Pasien: <b>{draftPreview.opForm?.patient || "(nama belum terisi)"}</b>{draftPreview.savedAt ? ` — disimpan otomatis ${new Date(draftPreview.savedAt).toLocaleString("id-ID")}` : ""}</div>
          <div style={{display:"flex",gap:8}}>
            <Btn sm onClick={restoreDraft} style={{flex:1}}>↩ Pulihkan Draft</Btn>
            <Btn sm outline color={C.g} onClick={discardDraft} style={{flex:1}}>🗑 Abaikan</Btn>
          </div>
        </div>
      )}
      <ScanFormulir
        showToast={showToast}
        onResults={(results)=>{
          if(results.length===1) applyOcrResult(results[0]);
          else { setShowDraftRecovery(false); autoSaveOcrBatch(results); }
        }}
      />
      {/* CLEANUP: blok "pilih salah satu kandidat OCR" DIHAPUS — ini kode mati
          peninggalan desain lama. Sejak alur routing berdasar jumlah hasil
          (results.length===1 → applyOcrResult, ≥2 → autoSaveOcrBatch),
          ocrCandidates tidak pernah lagi diisi data (hanya pernah di-set ke
          array kosong), sehingga blok ini TIDAK PERNAH tampil ke pengguna.
          Dibiarkan ada sebelumnya hanya membingungkan pembaca kode. */}
      {/* FIX AUDIT #39 (lanjutan): banner PERSISTEN (bukan toast yang hilang
          dalam 3.2 detik) yang tetap tampil selama form masih berisi data
          hasil OCR yang belum disimpan — pengingat kuat agar perawat tidak
          lupa/salah kira sudah tersimpan hanya karena sempat melihat toast
          sekilas. Hilang otomatis begitu form disimpan (ocrFilledFields
          direset) atau dikosongkan ulang. */}
      {ocrFilledFields.size>0 && (
        <div style={{background:"#FFFBEB",border:"1.5px solid #F59E0B",borderRadius:10,padding:"10px 14px",marginBottom:14,fontSize:13,color:"#92400E",fontWeight:700}}>
          📝 Form ini terisi otomatis dari hasil scan OCR — <u>BELUM TERSIMPAN</u>. Periksa data di bawah, lalu tekan tombol "Simpan" agar masuk ke Jadwal Operasi.
        </div>
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
          <div style={{flex:1}}><LF label="Tanggal" req><DateTxt style={iStyle("date")} value={opForm.date||""} onChange={(v:string)=>setR("date",v)}/></LF>{e.date&&<div style={{fontSize:11,color:C.d,marginTop:-8,marginBottom:8}}>⚠ {e.date}</div>}</div>
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
// ViewLaporan sekarang diimpor dari ./components/view-laporan (lihat import di atas)


/* ─── VIEW KIRIM WA ──────────────────────────────────────────────────── */
// ViewKirimWA sekarang diimpor dari ./components/view-kirim-wa (lihat import di atas)


/* ─── VIEW STATISTIK ─────────────────────────────────────────────────── */
// ViewStatistik sekarang diimpor dari ./components/view-statistik (lihat import di atas)


/* ─── VIEW STAF ──────────────────────────────────────────────────────── */
// ViewStaf sekarang diimpor dari ./components/view-staf,
// ImportExcel dari ./components/import-excel (lihat import di atas)


/* ─── VIEW ARSIP ─────────────────────────────────────────────────────── */
// ViewArsip sekarang diimpor dari ./components/view-arsip (lihat import di atas)


/* ─── VIEW LEMBUR ────────────────────────────────────────────────────── */
// BULAN_ID sekarang diimpor dari ./constants (lihat import di atas)
const nowYM = () => { const d=new Date(); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`; };
const ymLabel = (ym:string) => { const [y,m]=ym.split("-"); return `${BULAN_ID[parseInt(m)-1]} ${y}`; };

function ViewLembur({lemburPegawai, setLemburPegawai, lemburData, setLemburData, showToast, supaCfg, dbxCfg, role, upsertOneToSupa, deleteFromSupa, autoSelPeg, onAutoSelPegUsed}: any) {
  const [sub, setSub]         = useState<"catat"|"rekap"|"pegawai">("catat");
  const [selPeg, setSelPeg]   = useState<string>(autoSelPeg||"");
  const [selYM,  setSelYM]    = useState(nowYM());
  const [editRow, setEditRow] = useState<string|null>(null);
  const [rowForm, setRowForm] = useState<any>({});
  /* FIX AUDIT #44 (MEDIUM/HIGH — Stale Closure di handleSelfieKeluarOtomatis):
     handleSelfieKeluarOtomatis menunggu GPS+kamera (Promise.all, bisa makan
     waktu 1-3 detik atau lebih di jaringan RS yang lambat) sebelum membangun
     rowFormFinal dengan `{...rowForm, ...}` — tapi `rowForm` di situ adalah
     variabel closure yang dibekukan pada render SAAT FUNGSI MULAI dipanggil.
     Kalau user sempat mengetik di field lain (mis. "Keperluan Lembur",
     "Keterangan", jam masuk) selagi GPS/kamera masih diproses, ketikan itu
     TERSIMPAN di state rowForm (re-render terjadi normal), TAPI closure yang
     sedang berjalan tidak pernah "melihat" state baru itu — begitu proses
     selfie selesai, rowFormFinal dibangun dari rowForm versi LAMA, sehingga
     ketikan user H HILANG SENYAP saat setRowForm(rowFormFinal) dipanggil dan
     data yang salah/kurang itu ikut ter-upsert ke Supabase lewat saveEntry().
     Perbaikan: rowFormRef selalu disinkronkan ke rowForm terbaru lewat
     useEffect (pola yang sama seperti latestDataRef di komponen induk),
     lalu handleSelfieKeluarOtomatis membaca rowFormRef.current (bukan
     `rowForm` langsung) saat membangun objek final. */
  const rowFormRef = useRef(rowForm);
  useEffect(()=>{ rowFormRef.current = rowForm; },[rowForm]);
  /* REVISI: tandai sesi ini berasal dari login cepat "⏰ Input Lembur" (bukan
     dari klik manual "✚ Tambah Baris Lembur" di tab Lembur biasa). Dipakai
     untuk menyederhanakan tampilan form: HANYA tombol selfie+GPS yang
     tampil, tanpa field tanggal/jam/keperluan/keterangan/TTD/Simpan/Batal —
     karena semua sudah otomatis terekam & tersimpan begitu proses selesai. */
  const [quickLemburMode, setQuickLemburMode] = useState(false);
  /* Jam berjalan khusus tampilan mode sederhana (Masuk/Pulang) ala contoh
     aplikasi absensi resmi — update tiap detik hanya saat mode ini aktif. */
  const [quickNow, setQuickNow] = useState(() => new Date());
  useEffect(() => {
    if (!quickLemburMode) return undefined;
    const t = setInterval(() => setQuickNow(new Date()), 1000);
    return () => clearInterval(t);
  }, [quickLemburMode]);
  /* OPSI A: saat ViewLembur pertama mount dengan autoSelPeg dari PinScreen,
     auto-set selPeg dan reset autoSelPeg di root agar tidak terpicu lagi
     kalau user navigasi ke tab lain lalu kembali ke tab Lembur.
     ALUR OTOMATIS: sekaligus auto-buka form baris baru dan set flag
     autoTriggerSelfie agar selfie langsung terpicu tanpa klik tombol. */
  useEffect(()=>{
    if(autoSelPeg){
      setSelPeg(autoSelPeg);
      if(onAutoSelPegUsed) onAutoSelPegUsed();
      /* REVISI: mode tampilan sederhana (2 tombol besar Masuk/Pulang) —
         tidak lagi auto-buka form/auto-trigger selfie seperti sebelumnya.
         Perawat sendiri yang klik "Masuk" atau "Pulang" sesuai kebutuhan. */
      setQuickLemburMode(true);
    }
  },[autoSelPeg]);
  const [isLoadingSelfie, setIsLoadingSelfie] = useState(false);
  /* FIX BUG #3 + FIX AUDIT #20 (double-submit race — digeneralisasi):
     ref sinkron ini SEKARANG dijadikan guard umum di dalam writeLemburKey()
     itu sendiri (bukan hanya di handler Masuk/Pulang), sehingga SEMUA jalur
     penulisan (tombol "✓ Simpan" manual, Hapus, simpan TTD, import Excel,
     maupun Masuk/Pulang) sama-sama terlindungi dari klik ganda cepat tanpa
     perlu menduplikasi guard di setiap handler satu-satu. Dicek SEBELUM
     await apa pun — beda dengan state `isLoadingSelfie` yang baru efektif
     ke prop `disabled` setelah 1 render cycle. */
  const isSubmittingAbsenRef = useRef(false);
  /* ALUR OTOMATIS: setelah login via "⏰ Input Lembur", ViewLembur otomatis
     membuka form baru DAN langsung trigger selfie tanpa user perlu klik apapun.
  /* FALLBACK MANUAL: ditampilkan kalau GPS/kamera tidak aktif/ditolak,
     agar user tetap bisa isi jam keluar & tanggal secara manual */
  const [showManualKeluar, setShowManualKeluar] = useState(false);
  const [kepRuang, setKepRuang]   = useState("");
  const [kepBidang, setKepBidang] = useState("");
  const [pegForm,  setPegForm]    = useState({name:"", nik:""});
  const [rekapYM,  setRekapYM]    = useState(nowYM());
  const fileRef = useRef<HTMLInputElement>(null);
  /* FIX BUG GPS+SELFIE #2 & #5: simpan stream kamera aktif di ref supaya bisa
     dimatikan dari luar (tombol Batal / unmount), dan flag mounted untuk
     mencegah setState dipanggil setelah komponen unmount di tengah proses async */
  const activeSelfieStreamRef = useRef<MediaStream | null>(null);
  const isMountedSelfieRef = useRef(true);
  useEffect(() => {
    isMountedSelfieRef.current = true;
    return () => {
      isMountedSelfieRef.current = false;
      if (activeSelfieStreamRef.current) {
        activeSelfieStreamRef.current.getTracks().forEach(track => track.stop());
        activeSelfieStreamRef.current = null;
      }
    };
  }, []);
  /* Hentikan kamera secara aman dari mana saja (dipanggil saat Batal ditekan) */
  const stopActiveSelfieStream = () => {
    if (activeSelfieStreamRef.current) {
      activeSelfieStreamRef.current.getTracks().forEach(track => track.stop());
      activeSelfieStreamRef.current = null;
    }
  };
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
    let res;
    try {
      /* FIX AUDIT #25: timeout 15 detik — cukup toleran untuk jaringan RS
         yang lambat, tapi tetap memberi kepastian ke UI dalam waktu wajar
         alih-alih menggantung tanpa batas. */
      res = await withTimeout(
        upsertOneToSupa?.("kb_lembur_data", {id:key, ...newRecForKey, updated_at:new Date().toISOString()}) as Promise<any>,
        15000, "Simpan data lembur"
      );
    } catch (e: any) {
      showToast(`⚠ ${e?.message || "Gagal menyimpan — koneksi bermasalah"}`, C.d);
      return false;
    }
    if(!res?.ok){
      showToast(`⚠ Gagal menyimpan ke Supabase: ${res?.error||"kesalahan tidak diketahui"}`,C.d);
      return false;
    }
    setLemburData((p:Record<string,any>)=>({...p,[key]:newRecForKey}));
    return true;
  };

  /* FIX AUDIT #30 (CRITICAL — Race Condition / Lost Update pada writeLemburKey):
     writeLemburKey() menulis SELURUH array `entries` hasil hitung ulang di
     client. Kalau device/tab lain menulis ke `key` yang sama dalam window
     waktu network round-trip (mis. user submit absen masuk dari 2 tab, atau
     retry sesudah timeout), upsert kedua akan MENIMPA array `entries` yang
     baru saja ditulis device pertama — entry milik device pertama HILANG
     tanpa error apa pun, karena keduanya "berhasil" secara independen.

     Perbaikan: entri BARU (absen masuk / baris baru) tidak lagi ditulis lewat
     read-modify-write penuh, tapi lewat RPC Postgres yang meng-append secara
     atomik di sisi database (jsonb concat dalam satu statement UPDATE),
     sehingga dua penulisan bersamaan TIDAK saling menimpa.

     Jalankan migrasi berikut SEKALI di Supabase SQL Editor:

       create or replace function append_lembur_entry(p_key text, p_entry jsonb, p_base jsonb)
       returns void as $$
         insert into kb_lembur_data (id, data, updated_at)
         values (p_key, jsonb_set(p_base, '{entries}', jsonb_build_array(p_entry)), now())
         on conflict (id) do update
           set data = jsonb_set(
                 kb_lembur_data.data,
                 '{entries}',
                 coalesce(kb_lembur_data.data->'entries', '[]'::jsonb) || p_entry
               ),
               updated_at = now();
       $$ language sql;

     Jika RPC ini belum di-deploy, fallback otomatis ke writeLemburKey (mode
     lama) supaya aplikasi tetap jalan — tapi tetap rawan race, jadi migrasi
     di atas SANGAT direkomendasikan untuk segera dijalankan. */
  const writeLemburKeyAppend = async (baseRecForKey: any, entryBaru: any): Promise<boolean> => {
    if(!key) return false;
    try {
      const { error } = await withTimeout(
        SUPA_CLIENT.rpc("append_lembur_entry", {
          p_key: key,
          p_entry: entryBaru,
          p_base: { ...baseRecForKey, entries: [] },
        }) as unknown as Promise<any>,
        15000, "Simpan absen (atomic)"
      );
      if(!error){
        /* Sinkron optimis di state lokal — data otoritatif tetap datang lewat
           realtime/loadFromSupaTable, ini hanya agar UI langsung responsif. */
        setLemburData((p:Record<string,any>)=>{
          const prevEntries = p[key!]?.entries || baseRecForKey.entries || [];
          return {...p, [key!]: {...baseRecForKey, entries:[...prevEntries, entryBaru]}};
        });
        return true;
      }
      console.warn("[LEMBUR] RPC append_lembur_entry gagal/belum ada, fallback ke write penuh:", error);
    } catch (e) {
      console.warn("[LEMBUR] RPC append_lembur_entry error, fallback ke write penuh:", e);
    }
    /* Fallback: RPC belum di-deploy → pola lama (tetap ada risiko race,
       lebih baik daripada gagal total sebelum migrasi dijalankan). */
    const entriesNow: any[] = (key && lemburData[key]?.entries) || baseRecForKey.entries || [];
    return writeLemburKey({ ...baseRecForKey, entries: [...entriesNow, entryBaru] });
  };

  /* REVISI: saveEntry sekarang menerima parameter opsional `formOverride`.
     Dipakai oleh auto-save (login lembur) agar menyimpan data rowForm yang
     PALING BARU (hasil GPS+selfie) tanpa risiko stale-state, sementara
     pemakaian normal (tombol "✓ Simpan" manual) tetap jalan seperti biasa
     dengan memakai rowForm dari state seperti sebelumnya. */
  const saveEntry = async (formOverride?: any) => {
    /* FIX AUDIT #20 (MEDIUM — double-submit "✓ Simpan" manual): guard sinkron
       yang sama dipakai handler Masuk/Pulang, diterapkan juga di sini supaya
       klik ganda cepat pada tombol Simpan tidak membuat baris terduplikasi. */
    if (isSubmittingAbsenRef.current) return;
    isSubmittingAbsenRef.current = true;
    const src = formOverride || rowForm;
    const jenis = src.jenisEntri || "kerja";
    /* UPDATE: Tanggal Awal Lembur & Jam Absen Masuk TIDAK WAJIB diisi saat simpan —
       boleh kosong dan diisi belakangan (sewaktu-waktu lewat Edit). Baris tetap
       bisa disimpan sebagai draft. Validasi wajib lama (tanggalAwal, jamMasuk)
       sengaja dihapus sesuai permintaan; jamKeluar tetap wajib lewat tombol
       GPS+Selfie HANYA kalau diisi (lihat handleSelfieKeluarOtomatis) — kalau
       belum diisi, baris kerja tetap bisa disimpan sebagai draft kosong juga. */
    if(jenis==="ambil" && (src.jumlahJamAmbil===undefined||src.jumlahJamAmbil===""||Number(src.jumlahJamAmbil)<=0)){
      isSubmittingAbsenRef.current = false;
      showToast("Jumlah jam diambil wajib diisi (lebih dari 0) untuk Pengambilan Lembur",C.d); return;
    }
    /* Normalisasi: pastikan field yang tidak relevan untuk jenis ini tidak menyimpan sisa data lama */
    const cleanForm = jenis==="kerja"
      ? {...src, jenisEntri:"kerja", jumlahJamAmbil:""}
      /* UPDATE LEMBUR GPS+SELFIE: bersihkan lokasiKeluar & fotoKeluarUrl jika entri berupa Pengambilan Lembur,
         karena field ini hanya relevan untuk verifikasi selesai Kerja Lembur */
      : {...src, jenisEntri:"ambil", jamMasuk:"", jamKeluar:"", lokasiKeluar:"", fotoKeluarUrl:""};
    /* FIX BUG #7/#8: baca entries & rec TERBARU dari lemburData[key] tepat
       sebelum menulis (bukan snapshot closure `entries`/`rec` dari awal
       render) — meminimalkan risiko nomor (`no`) bentrok dan risiko
       menimpa perubahan device lain yang masuk lewat realtime sesaat
       sebelum tombol Simpan diklik. */
    const recNow: any = (key && lemburData[key]) || rec;
    const entriesNow: any[] = recNow.entries || [];
    showToast("⟳ Menyimpan...",C.p);
    /* FIX AUDIT #30: baris BARU pakai append atomic (RPC) agar tidak
       menimpa entry dari device/tab lain yang menulis di window waktu yang
       sama; baris EDIT tetap pakai write penuh karena mengubah field pada
       entry yang sudah ada (butuh RPC update-by-id terpisah untuk atomik
       penuh — lihat catatan pada writeLemburKeyAppend). */
    const ok = editRow
      ? await writeLemburKey({...recNow, entries: entriesNow.map((e:any)=>e.id===editRow?{...e,...cleanForm}:e)})
      : await writeLemburKeyAppend(recNow, {id:gId(),...cleanForm,no:entriesNow.length+1});
    isSubmittingAbsenRef.current = false;
    if(ok){ setEditRow(null); setRowForm({}); showToast("✓ Baris disimpan & tersinkron",C.s); }
  };

  /* ── UPDATE LEMBUR GPS+SELFIE (BARU DITAMBAHKAN) ──────────────────────
     Mengunci absen keluar Kerja Lembur secara otomatis: ambil koordinat GPS
     presisi tinggi, jepret foto selfie via kamera depan, lalu unggah ke
     Dropbox dan simpan tautan langsungnya. Ini murni fitur tambahan — tidak
     mengubah/menghapus alur input manual tanggalAwal/jamMasuk yang sudah ada.

     KOMPATIBILITAS LINTAS DEVICE:
     GPS & kamera browser TIDAK BISA dijamin 100% jalan di semua HP/browser —
     ini batasan hardware & platform, bukan sesuatu yang bisa "diperbaiki" lewat
     kode semata. Yang kita bisa lakukan: deteksi sedini mungkin, beri pesan
     yang jelas, coba beberapa strategi fallback otomatis, dan kalau tetap
     gagal — pengguna SELALU punya jalan keluar lewat field manual di bawah
     (lihat bagian "Jika selfie/GPS gagal" di form).
     1) Cek koneksi HTTPS dulu — geolocation & kamera diblokir browser modern
        kalau bukan https/localhost, ini penyebab kegagalan paling umum.
     2) GPS: coba presisi tinggi dulu (timeout pendek), kalau timeout/gagal,
        otomatis coba ulang dengan presisi rendah (lebih cepat & lebih
        kompatibel di dalam gedung / device GPS lemah) sebelum benar-benar
        menyerah.
     3) Kamera: dukung juga API getUserMedia versi lama (vendor-prefixed),
        dipakai sebagian browser Android/iOS lawas yang belum punya
        navigator.mediaDevices standar.
     4) Capture foto: pakai canvas.toBlob, dengan fallback ke toDataURL untuk
        browser sangat lama yang belum mendukung toBlob. */
  const getUserMediaCompat = (constraints: MediaStreamConstraints): Promise<MediaStream> => {
    if (navigator.mediaDevices?.getUserMedia) {
      return navigator.mediaDevices.getUserMedia(constraints);
    }
    /* Fallback untuk browser lawas (Android WebView lama, dsb.) yang masih
       memakai API vendor-prefixed, bukan navigator.mediaDevices standar. */
    const legacyGetUserMedia =
      (navigator as any).getUserMedia ||
      (navigator as any).webkitGetUserMedia ||
      (navigator as any).mozGetUserMedia ||
      (navigator as any).msGetUserMedia;
    if (legacyGetUserMedia) {
      return new Promise((resolve, reject) => {
        legacyGetUserMedia.call(navigator, constraints, resolve, reject);
      });
    }
    return Promise.reject(new Error("getUserMedia tidak tersedia di browser ini"));
  };

  const captureBlobCompat = (canvas: HTMLCanvasElement): Promise<Blob | null> => {
    return new Promise((resolve) => {
      if (typeof canvas.toBlob === "function") {
        canvas.toBlob((b) => resolve(b), "image/jpeg", 0.75);
        return;
      }
      /* Fallback untuk browser sangat lama tanpa dukungan canvas.toBlob:
         pakai toDataURL lalu ubah jadi Blob secara manual. */
      try {
        const dataUrl = canvas.toDataURL("image/jpeg", 0.75);
        const arr = dataUrl.split(",");
        const bstr = atob(arr[1]);
        const u8 = new Uint8Array(bstr.length);
        for (let i = 0; i < bstr.length; i++) u8[i] = bstr.charCodeAt(i);
        resolve(new Blob([u8], { type: "image/jpeg" }));
      } catch {
        resolve(null);
      }
    });
  };

  /* Bungkus getCurrentPosition jadi Promise + retry: presisi tinggi dulu,
     kalau gagal/timeout otomatis coba lagi dengan presisi rendah. Ini
     menambah peluang sukses di device/lokasi dengan sinyal GPS lemah
     (mis. dalam gedung rumah sakit) tanpa perlu campur tangan pengguna. */
  const getPositionWithFallback = (): Promise<GeolocationPosition> => {
    return new Promise((resolve, reject) => {
      navigator.geolocation.getCurrentPosition(
        resolve,
        () => {
          if (showToast) showToast("⏳ GPS presisi tinggi gagal, mencoba mode hemat (lebih cepat)...", C.p);
          navigator.geolocation.getCurrentPosition(
            resolve,
            reject,
            { enableHighAccuracy: false, timeout: 8000, maximumAge: 30000 }
          );
        },
        { enableHighAccuracy: true, timeout: 6000 }
      );
    });
  };

  /* Ambil 1 frame selfie dari kamera depan & upload ke Dropbox. Mengembalikan
     link foto, atau null jika kamera/izin tidak tersedia — TANPA melempar
     error ke pemanggil, supaya tanggal/jam tetap bisa terkunci walau bagian
     ini gagal sepenuhnya (lihat handleSelfieKeluarOtomatis di bawah). */
  const captureSelfieAndUpload = async (): Promise<string | null> => {
    let stream: MediaStream;
    try {
      /* facingMode "user" sebagai preferensi (ideal), bukan syarat mutlak —
         supaya device tanpa kamera depan tetap bisa pakai kamera yang ada. */
      stream = await getUserMediaCompat({ video: { facingMode: { ideal: "user" } } });
    } catch {
      return null; // kamera nonaktif/izin ditolak/tidak ada — bukan error fatal
    }
    if (!isMountedSelfieRef.current) {
      stream.getTracks().forEach(track => track.stop());
      return null;
    }
    activeSelfieStreamRef.current = stream;

    const video = document.createElement("video");
    video.srcObject = stream;
    video.muted = true;
    video.playsInline = true;
    (video as any).setAttribute("webkit-playsinline", "true"); // kompatibilitas Safari iOS lama

    /* Tunggu video siap dengan timeout singkat — dipersingkat dari versi
       sebelumnya supaya proses keseluruhan lebih cepat (sesuai permintaan:
       "dalam memproses data cepat"). */
    await new Promise<void>((resolve) => {
      let done = false;
      const finish = () => { if (!done) { done = true; resolve(); } };
      video.onloadedmetadata = () => { video.play().catch(() => {}); setTimeout(finish, 200); };
      setTimeout(finish, 1500); // fallback keras jika metadata tak kunjung siap
      video.play().catch(() => {});
    });

    if (!isMountedSelfieRef.current) {
      stream.getTracks().forEach(track => track.stop());
      activeSelfieStreamRef.current = null;
      return null;
    }

    const canvas = document.createElement("canvas");
    canvas.width = 400;
    canvas.height = 300;
    const ctx = canvas.getContext("2d");
    if (!ctx) {
      stream.getTracks().forEach(track => track.stop());
      activeSelfieStreamRef.current = null;
      return null;
    }
    try {
      ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
    } catch {
      stream.getTracks().forEach(track => track.stop());
      activeSelfieStreamRef.current = null;
      return null;
    }

    const blob = await captureBlobCompat(canvas);
    /* Segera matikan hardware kamera setelah frame didapat demi privasi & baterai */
    stream.getTracks().forEach(track => track.stop());
    activeSelfieStreamRef.current = null;
    if (!blob) return null;

    try {
      const namaFile = `selfie_lembur_${Date.now()}.jpg`;
      /* UPDATE LEMBUR GPS+SELFIE — PENYESUAIAN KEAMANAN:
         Upload disalurkan lewat callDbxProxy (token Dropbox aman di server,
         Supabase Edge Function "dropbox-proxy"), bukan API Dropbox langsung
         dari browser. ⚠ WAJIB: pastikan Edge Function sudah punya handler
         action "upload_image" sebelum fitur ini berfungsi di production. */
      const CHUNK = 8192;
      const bytes = new Uint8Array(await blob.arrayBuffer());
      let binary = "";
      for (let i = 0; i < bytes.length; i += CHUNK) {
        binary += String.fromCharCode(...bytes.subarray(i, i + CHUNK));
      }
      const base64Foto = btoa(binary);
      const uploadRes = await callDbxProxy({
        action: "upload_image",
        path: `/foto_lembur/${namaFile}`,
        base64: base64Foto,
        mime: "image/jpeg",
      });
      if (!uploadRes.ok) return null;
      return uploadRes.data?.url || uploadRes.data?.directLink || null;
    } catch {
      return null;
    }
  };

  /* ── UPDATE: tanggal & jam SELALU tercatat, GPS+kamera bersifat pelengkap ──
     Sebelumnya jika GPS atau kamera mati/ditolak izinnya, seluruh proses
     gagal dan tanggal/jam keluar TIDAK tercatat sama sekali (user terjebak
     harus isi manual). Sekarang:
     1) Tanggal & jam keluar dikunci LANGSUNG di awal (instan, tidak menunggu
        apa pun) — ini data yang paling penting dan tidak boleh gagal hanya
        karena GPS/kamera nonaktif.
     2) GPS & kamera dijalankan PARALEL (bukan berurutan seperti sebelumnya)
        supaya proses jauh lebih cepat — hasil keduanya bersifat tambahan
        (lokasi & foto), dilengkapi kalau berhasil, dilewati kalau tidak.
     3) Jika GPS/kamera aktif & berhasil → perilaku identik seperti sebelumnya
        (lokasi & foto ikut terkunci). Jika tidak aktif/ditolak → tanggal &
        jam tetap tersimpan, hanya lokasi/foto yang kosong. */
  const handleSelfieKeluarOtomatis = async () => {
    setIsLoadingSelfie(true);
    if (showToast) showToast("⏳ Memproses data keluar...", C.p);

    // STEP 1 — instan, selalu berhasil: kunci tanggal & jam sekarang juga,
    // tanpa menunggu izin GPS/kamera sama sekali.
    const sekarang = new Date();
    /* FIX AUDIT #21 (CRITICAL — Pergeseran Tanggal Timezone): sebelumnya
       memakai toISOString().split("T")[0] yang SELALU dalam UTC, sehingga
       pukul 00:00–06:59 WIB tercatat sebagai tanggal KEMARIN. Diganti ke
       toLocalISODate() yang memakai komponen tanggal LOKAL, konsisten
       dengan seluruh bagian aplikasi lain (lihat definisi & peringatan di
       baris ~489). */
    const tglKeluar = toLocalISODate(sekarang);
    const jamKeluar = sekarang.toTimeString().slice(0, 5);
    if (isMountedSelfieRef.current) {
      setRowForm((prev: any) => ({ ...prev, tanggalAkhir: tglKeluar, jamKeluar: jamKeluar }));
    }

    const secure = typeof window === "undefined" || window.isSecureContext;

    // STEP 2 — GPS & kamera berjalan BERSAMAAN (paralel) untuk mempercepat proses.
    // Masing-masing menelan error sendiri (catch -> null) sehingga satu gagal
    // tidak menggagalkan yang lain ataupun tanggal/jam yang sudah terkunci di STEP 1.
    const gpsPromise: Promise<string | null> = (secure && navigator.geolocation)
      ? getPositionWithFallback()
          .then(pos => `Lat: ${pos.coords.latitude.toFixed(5)}, Lon: ${pos.coords.longitude.toFixed(5)}`)
          .catch(() => null)
      : Promise.resolve(null);

    const fotoPromise: Promise<string | null> = secure
      ? captureSelfieAndUpload().catch(() => null)
      : Promise.resolve(null);

    const [lokasiString, directLink] = await Promise.all([gpsPromise, fotoPromise]);

    if (!isMountedSelfieRef.current) return;

    /* REVISI: bangun objek rowForm final secara eksplisit (bukan hanya lewat
       setRowForm functional-update) supaya nilai terbarunya bisa langsung
       dipakai untuk auto-save di bawah tanpa menunggu re-render/stale state.
       FIX AUDIT #44: dibangun dari rowFormRef.current (SELALU terkini),
       bukan dari `rowForm` closure yang beku sejak fungsi ini mulai
       dipanggil — lihat penjelasan lengkap di deklarasi rowFormRef. */
    const rowFormFinal = {
      ...rowFormRef.current,
      tanggalAkhir: tglKeluar,
      jamKeluar: jamKeluar,
      lokasiKeluar: lokasiString || rowFormRef.current.lokasiKeluar || "",
      fotoKeluarUrl: directLink || rowFormRef.current.fotoKeluarUrl || ""
    };
    setRowForm(rowFormFinal);

    const gpsOk = !!lokasiString, fotoOk = !!directLink;

    if (showToast) {
      if (gpsOk && fotoOk) {
        showToast("✓ Parameter Keluar Berhasil Dikunci (Tanggal, Jam, GPS & Foto)!", C.s);
        setShowManualKeluar(false);
      } else if (!gpsOk && !fotoOk) {
        showToast("✓ Tanggal & Jam keluar tercatat. GPS/Kamera tidak aktif — lengkapi manual jika perlu.", C.w);
        /* FALLBACK: tampilkan input manual agar user bisa lengkapi */
        setShowManualKeluar(true);
      } else {
        showToast(`✓ Tanggal & Jam tercatat${gpsOk ? " + GPS" : ""}${fotoOk ? " + Foto" : ""}. ${!gpsOk ? "GPS" : "Kamera"} tidak tersedia.`, C.w);
        setShowManualKeluar(false);
      }
    }
    setIsLoadingSelfie(false);

    /* REVISI: LOGIN LEMBUR & TOMBOL MANUAL — auto-save SELALU berjalan.
       Sebelumnya hanya alur login otomatis yang auto-save; sekarang tombol
       manual "📷 Klik untuk Selfie & Ambil Jam Keluar" juga langsung
       menyimpan baris begitu tanggal, jam, GPS & foto selesai diproses —
       perawat tidak perlu menyentuh apa pun lagi, cukup login dan boleh
       langsung keluar. */
    await saveEntry(rowFormFinal);
  };

  /* ── MODE TAMPILAN SEDERHANA (2 tombol besar: Masuk & Pulang) ─────────
     Sesuai contoh aplikasi absensi resmi RS Panti Rini: begitu login lewat
     "⏰ Input Lembur", perawat hanya lihat nama + tanggal/jam + 2 tombol
     besar. Klik "Masuk" → GPS+selfie+save langsung otomatis (buat baris
     baru). Klik "Pulang" → GPS+selfie+save langsung otomatis (lengkapi
     baris kerja lembur milik pegawai ini yang masih terbuka — sudah ada
     jamMasuk tapi belum ada jamKeluar; kalau tidak ada baris terbuka,
     buat baris baru berisi jam pulang saja). Kedua tombol bisa diklik
     kapan saja, tidak saling mengunci satu sama lain. */
  const captureGpsSelfie = async (): Promise<{lokasi:string; foto:string}> => {
    const secure = typeof window === "undefined" || window.isSecureContext;
    const gpsPromise: Promise<string | null> = (secure && navigator.geolocation)
      ? getPositionWithFallback()
          .then(pos => `Lat: ${pos.coords.latitude.toFixed(5)}, Lon: ${pos.coords.longitude.toFixed(5)}`)
          .catch(() => null)
      : Promise.resolve(null);
    const fotoPromise: Promise<string | null> = secure
      ? captureSelfieAndUpload().catch(() => null)
      : Promise.resolve(null);
    const [lokasi, foto] = await Promise.all([gpsPromise, fotoPromise]);
    return { lokasi: lokasi || "", foto: foto || "" };
  };

  const handleAbsenMasuk = async () => {
    /* FIX BUG #1 (race selPeg/key null) + FIX BUG #3 (double-submit race):
       cek DUA guard sinkron sebelum apa pun lain dijalankan. */
    if (isSubmittingAbsenRef.current) return;
    if (!key) { showToast("⚠ Data pegawai belum siap, coba lagi sesaat.", C.w); return; }
    isSubmittingAbsenRef.current = true;
    setIsLoadingSelfie(true);
    if (showToast) showToast("⏳ Memproses absen masuk...", C.p);
    const sekarang = new Date();
    /* FIX AUDIT #21 (CRITICAL — Pergeseran Tanggal Timezone): pakai
       toLocalISODate() bukan toISOString(), lihat penjelasan lengkap di
       definisi toLocalISODate baris ~489. */
    const tgl = toLocalISODate(sekarang);
    const jam = sekarang.toTimeString().slice(0, 5);
    const { lokasi, foto } = await captureGpsSelfie();
    if (!isMountedSelfieRef.current) { isSubmittingAbsenRef.current = false; return; }
    /* FIX BUG #4 (nomor duplikat): baca entries TERBARU dari lemburData saat
       ini (bukan closure lama) tepat sebelum menulis, meminimalkan risiko
       nomor bentrok akibat penulisan bersamaan dari device/tab lain. */
    const entriesNow: any[] = (lemburData[key!]?.entries) || entries;
    const entryBaru = {
      id: gId(), no: entriesNow.length + 1,
      jenisEntri: "kerja", tanggalAwal: tgl, tanggalAkhir: "",
      jamMasuk: jam, jamKeluar: "",
      lokasiMasuk: lokasi, fotoMasukUrl: foto,
      lokasiKeluar: "", fotoKeluarUrl: "",
      jumlahJamAmbil: "", keperluanLembur: "", keterangan: "", ttd: ""
    };
    showToast("⟳ Menyimpan absen masuk...", C.p);
    /* FIX AUDIT #30: append atomic, bukan read-modify-write penuh */
    const ok = await writeLemburKeyAppend(lemburData[key!]||rec, entryBaru);
    isSubmittingAbsenRef.current = false;
    /* FIX AUDIT #19 (MEDIUM — setState setelah unmount): guard mounted
       harus dicek lagi SETELAH await KEDUA (writeLemburKey), bukan hanya
       setelah await pertama (captureGpsSelfie). Kalau komponen unmount
       persis saat writeLemburKey masih berjalan, setIsLoadingSelfie/
       showToast di bawah ini akan mencoba update komponen yang sudah
       tidak ada — sumber React warning & potensi memory leak. */
    if (!isMountedSelfieRef.current) return;
    setIsLoadingSelfie(false);
    if (ok) {
      /* FIX BUG #5: kalau GPS & foto GAGAL total, tampilkan warning kuning,
         bukan toast sukses hijau yang menutupi fakta tidak ada bukti lokasi/foto. */
      if (!lokasi && !foto) {
        showToast(`✓ Absen Masuk tercatat pukul ${jam}. ⚠ GPS & Kamera tidak aktif — tanpa bukti lokasi/foto.`, C.w);
      } else if (!lokasi || !foto) {
        showToast(`✓ Absen Masuk tercatat pukul ${jam}${lokasi?" + GPS":""}${foto?" + Foto":""}. ⚠ ${!lokasi?"GPS":"Kamera"} tidak aktif.`, C.w);
      } else {
        showToast(`✓ Absen Masuk tercatat pukul ${jam} + GPS + Foto`, C.s);
      }
    }
  };

  const handleAbsenPulang = async () => {
    if (isSubmittingAbsenRef.current) return;
    if (!key) { showToast("⚠ Data pegawai belum siap, coba lagi sesaat.", C.w); return; }
    isSubmittingAbsenRef.current = true;
    setIsLoadingSelfie(true);
    if (showToast) showToast("⏳ Memproses absen pulang...", C.p);
    const sekarang = new Date();
    /* FIX AUDIT #21 (CRITICAL — Pergeseran Tanggal Timezone): sama seperti
       handleAbsenMasuk — pakai toLocalISODate() bukan toISOString(). */
    const tgl = toLocalISODate(sekarang);
    const jam = sekarang.toTimeString().slice(0, 5);
    const { lokasi, foto } = await captureGpsSelfie();
    if (!isMountedSelfieRef.current) { isSubmittingAbsenRef.current = false; return; }
    /* FIX BUG #2 (salah baris lintas tanggal): baca entries TERBARU + filter
       WAJIB tanggalAwal === tgl (hari ini) saat mencari baris terbuka. Baris
       "Masuk" dari hari sebelumnya yang lupa di-Pulang-kan TIDAK BOLEH
       tertimpa oleh absen pulang hari ini — kalau ada, biarkan tetap
       terbuka (perawat/admin bisa lengkapi manual lewat Edit) dan buat
       baris baru untuk hari ini. */
    const entriesNow: any[] = (lemburData[key!]?.entries) || entries;
    const idxTerbuka = [...entriesNow].reverse().findIndex((e: any) =>
      e.jenisEntri !== "ambil" && e.jamMasuk && !e.jamKeluar && e.tanggalAwal === tgl
    );
    let upd: any[];
    if (idxTerbuka >= 0) {
      const realIdx = entriesNow.length - 1 - idxTerbuka;
      upd = entriesNow.map((e: any, i: number) => i === realIdx
        ? { ...e, tanggalAkhir: tgl, jamKeluar: jam, lokasiKeluar: lokasi || e.lokasiKeluar || "", fotoKeluarUrl: foto || e.fotoKeluarUrl || "" }
        : e);
    } else {
      upd = [...entriesNow, {
        id: gId(), no: entriesNow.length + 1,
        jenisEntri: "kerja", tanggalAwal: tgl, tanggalAkhir: tgl,
        jamMasuk: "", jamKeluar: jam,
        lokasiKeluar: lokasi, fotoKeluarUrl: foto,
        jumlahJamAmbil: "", keperluanLembur: "", keterangan: "", ttd: ""
      }];
    }
    showToast("⟳ Menyimpan absen pulang...", C.p);
    const ok = await writeLemburKey({ ...(lemburData[key!]||rec), entries: upd });
    isSubmittingAbsenRef.current = false;
    /* FIX AUDIT #19: sama seperti handleAbsenMasuk — cek mounted lagi
       SETELAH await writeLemburKey, bukan hanya setelah captureGpsSelfie. */
    if (!isMountedSelfieRef.current) return;
    setIsLoadingSelfie(false);
    if (ok) {
      if (!lokasi && !foto) {
        showToast(`✓ Absen Pulang tercatat pukul ${jam}. ⚠ GPS & Kamera tidak aktif — tanpa bukti lokasi/foto.`, C.w);
      } else if (!lokasi || !foto) {
        showToast(`✓ Absen Pulang tercatat pukul ${jam}${lokasi?" + GPS":""}${foto?" + Foto":""}. ⚠ ${!lokasi?"GPS":"Kamera"} tidak aktif.`, C.w);
      } else {
        showToast(`✓ Absen Pulang tercatat pukul ${jam} + GPS + Foto`, C.s);
      }
    }
  };

  const delEntry = async (id:string) => {
    /* FIX BUG #8: baca rec/entries terbaru dari lemburData[key], bukan
       closure lama, sebelum menghapus & menulis ulang. */
    const recNow: any = (key && lemburData[key]) || rec;
    const entriesNow: any[] = recNow.entries || [];
    const upd = entriesNow.filter((e:any)=>e.id!==id).map((e:any,i:number)=>({...e,no:i+1}));
    const ok = await writeLemburKey({...recNow, entries:upd});
    if(ok) showToast("Baris dihapus & tersinkron");
  };

  const saveSigs = async () => {
    if(!key) return;
    /* FIX BUG #8: pakai rec TERBARU dari lemburData[key] — kalau tidak,
       menyimpan tanda tangan bisa ikut menimpa `entries` dengan versi lama
       (closure) walau field yang diubah cuma kepRuang/kepBidang. */
    const recNow: any = lemburData[key] || rec;
    const ok = await writeLemburKey({...recNow, kepRuang, kepBidang, savedAt:fNow()});
    if(ok) showToast("✓ Tanda tangan tersimpan & tersinkron",C.s);
  };

  const handleImportJadwal = (file:File) => {
    const reader = new FileReader();
    /* FIX AUDIT #28: sama seperti handleFile di ImportExcel. */
    reader.onerror = () => { showToast("Gagal membaca file — coba pilih file lain", C.d); };
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
        /* FIX BUG #8: mulai dari entries TERBARU (lemburData[key]), bukan
           closure `entries` lama — mencegah import Excel menimpa absen yang
           baru masuk lewat realtime sesaat sebelum proses import selesai. */
        const recNow: any = (key && lemburData[key]) || rec;
        let newEntries: any[] = [...(recNow.entries || [])];
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
        const added = newEntries.length-(recNow.entries||[]).length;
        showToast(`⟳ Menyimpan ${added} baris lembur...`,C.p);
        const ok = await writeLemburKey({...recNow, entries:newEntries});
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
    /* Table header
       UPDATE LEMBUR GPS+SELFIE: 2 kolom baru ditambahkan di ujung kanan (Kolom I & J).
       Kolom A-H (index 0-7) ORISINAL — urutan & isinya tidak berubah sama sekali,
       sehingga rumus rekap manual/makro Excel yang sudah ada tetap aman. */
    const TH=["No","Tgl Awal Lembur","Tgl Akhir Lembur","Jam Absen Masuk","Jam Absen Keluar","Keperluan Lembur","Keterangan","Tanda Tangan","Lokasi Keluar (GPS)","Tautan Foto Selfie"];
    TH.forEach((h,i)=>setC(7,i,h,true,true,true,"16685F"));
    /* Table rows */
    entries.forEach((e:any,i:number)=>{
      const r=8+i;
      const isAmbil = e.jenisEntri==="ambil";
      const masukCol  = isAmbil ? "—" : e.jamMasuk;
      const keluarCol = isAmbil ? "—" : e.jamKeluar;
      const kepCol    = isAmbil ? `[PENGAMBILAN -${Number(e.jumlahJamAmbil)||0} jam] ${e.keperluanLembur||""}`.trim() : e.keperluanLembur;
      /* Data lama (sebelum sistem selfie ada) otomatis terisi strip (—) — aman, tidak error */
      const gpsCol    = e.lokasiKeluar || "—";
      const selfieCol = e.fotoKeluarUrl || "—";
      /* FIX AUDIT #41: pakai i+1 (index array), bukan e.no statis yang bisa bentrok */
      [i+1,e.tanggalAwal||e.jadwalDinas||"",e.tanggalAkhir||"",masukCol,keluarCol,kepCol,e.keterangan,e.ttd||"",gpsCol,selfieCol].forEach((v,c)=>setC(r,c,String(v||""),false,c===0||c===3||c===4,true,i%2===0?"F0FFF8":undefined));
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
    /* Merges
       UPDATE LEMBUR GPS+SELFIE: merge judul/header diperlebar sampai Kolom J (index 9)
       agar tampilan tetap rapi penuh selebar tabel — TIDAK mengubah posisi merge
       ringkasan saldo & tanda tangan yang tetap di kolom A-H (0-7) seperti semula. */
    ws["!merges"]=[
      merge(0,0,0,9), merge(1,0,1,9), merge(2,0,2,9),
      merge(3,0,3,1), merge(3,2,3,7),
      merge(4,0,4,1), merge(4,2,4,7),
      merge(5,0,5,1), merge(5,2,5,7),
      merge(sigRow+1,0,sigRow+1,3), merge(sigRow+1,4,sigRow+1,7),
      merge(sigRow+2,0,sigRow+2,3), merge(sigRow+2,4,sigRow+2,7),
      merge(sigRow+6,0,sigRow+6,3), merge(sigRow+6,4,sigRow+6,7),
    ];
    ws["!cols"]=[{wch:4},{wch:16},{wch:16},{wch:16},{wch:16},{wch:26},{wch:22},{wch:16},{wch:24},{wch:32}];
    ws["!ref"]=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:sigRow+7,c:9}});
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
        /* FIX AUDIT #41: pakai i+1 (index array), bukan e.no statis yang bisa bentrok */
        [String(i+1),e.tanggalAwal||e.jadwalDinas||"",e.tanggalAkhir||"",masukCol,keluarCol,kepCol,e.keterangan||""].forEach((v,c)=>s2(row,c,v,false,c===0||c===3||c===4,true,bg));
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
          const rows=ents.map((e:any,i:number)=>[i+1,e.tanggalAwal||"",e.tanggalAkhir||"",e.jamMasuk||"",e.jamKeluar||"",sanitizeCellForExcel(e.keperluanLembur||""),sanitizeCellForExcel(e.keterangan||"")]);
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

  /* ── LAYAR MODE SEDERHANA: hanya nama + tanggal/jam + 2 tombol besar
     (Masuk & Pulang), sesuai contoh aplikasi absensi resmi RS Panti Rini.
     Ditampilkan sebagai layar penuh menggantikan seluruh UI tab Lembur
     biasa (sub-tab, tabel, dsb) selama quickLemburMode aktif. */
  if (quickLemburMode) {
    const tglLabel = quickNow.toLocaleDateString("id-ID",{day:"2-digit",month:"2-digit",year:"numeric"}).split("/").join("-");
    const jamLabel = quickNow.toTimeString().slice(0,5).replace(":",".");
    return (
      <div style={{minHeight:"70vh",display:"flex",flexDirection:"column",alignItems:"center",paddingTop:24,paddingBottom:24}}>
        <div style={{fontSize:15,color:C.tL,marginBottom:2}}>Sistem Koordinasi Kamar Bedah</div>
        <div style={{fontSize:17,fontWeight:800,color:C.b,marginBottom:14,textAlign:"center"}}>{peg?.name || "Pegawai"}</div>
        <div style={{fontSize:20,fontWeight:700,color:C.b}}>
          {tglLabel} <span style={{color:C.tL,fontWeight:500}}>{jamLabel}</span>
        </div>
        {isLoadingSelfie && (
          <div style={{marginTop:16,fontSize:13,color:C.p,fontWeight:600}}>⏳ Memproses GPS &amp; Selfie — mohon tunggu...</div>
        )}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,width:"100%",maxWidth:420,marginTop:32,padding:"0 16px"}}>
          <button
            onClick={handleAbsenMasuk}
            disabled={isLoadingSelfie || !key}
            style={{aspectRatio:"1",background:C.p,color:"#fff",border:"none",borderRadius:16,fontSize:22,fontWeight:800,cursor:(isLoadingSelfie||!key)?"default":"pointer",opacity:(isLoadingSelfie||!key)?0.6:1,boxShadow:"0 2px 10px rgba(0,0,0,.12)"}}
          >
            Masuk
          </button>
          <button
            onClick={handleAbsenPulang}
            disabled={isLoadingSelfie || !key}
            style={{aspectRatio:"1",background:C.p,color:"#fff",border:"none",borderRadius:16,fontSize:22,fontWeight:800,cursor:(isLoadingSelfie||!key)?"default":"pointer",opacity:(isLoadingSelfie||!key)?0.6:1,boxShadow:"0 2px 10px rgba(0,0,0,.12)"}}
          >
            Pulang
          </button>
        </div>
        <button
          type="button"
          onClick={()=>{stopActiveSelfieStream();setIsLoadingSelfie(false);setQuickLemburMode(false);}}
          style={{marginTop:28,background:"none",border:"none",color:C.tL,fontSize:12,fontWeight:600,cursor:"pointer",fontFamily:"inherit",textDecoration:"underline"}}
        >
          Buka tampilan lengkap / edit manual
        </button>
      </div>
    );
  }

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

              {/* Actions — UPDATE: disederhanakan jadi 1 tombol utama besar +
                  2 aksi sekunder kecil di bawahnya agar UI tidak crowded.
                  Import Excel & Download Excel tetap tersedia tapi tidak
                  mengambil visual weight yang sama dengan aksi utama. */}
              <div style={{marginBottom:14}}>
                <Btn full onClick={()=>{stopActiveSelfieStream();setIsLoadingSelfie(false);setShowManualKeluar(false);setEditRow(null);setQuickLemburMode(false);setRowForm({jenisEntri:"kerja",tanggalAwal:"",tanggalAkhir:"",jamMasuk:"",jamKeluar:"",jumlahJamAmbil:"",keperluanLembur:"",keterangan:"",ttd:""});}} style={{marginBottom:8,padding:"14px",fontSize:15,fontWeight:800}}>✚ Tambah Baris Lembur</Btn>
                <div style={{display:"flex",gap:8}}>
                  <button onClick={()=>fileRef.current?.click()} style={{flex:1,background:"#F4F6F8",border:`1px solid ${C.b}`,borderRadius:10,color:C.tL,fontSize:11,fontWeight:700,padding:"8px 10px",cursor:"pointer",fontFamily:"inherit"}}>📥 Import Excel</button>
                  <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={e=>{const f=e.target.files?.[0];if(f)handleImportJadwal(f);e.target.value="";}}/>
                  <button onClick={downloadExcel} style={{flex:1,background:"#F4F6F8",border:`1px solid ${C.b}`,borderRadius:10,color:C.tL,fontSize:11,fontWeight:700,padding:"8px 10px",cursor:"pointer",fontFamily:"inherit"}}>⬇ Download Excel</button>
                </div>
              </div>

              {/* Add/Edit row form — muncul kalau rowForm sudah diisi (Tambah/Edit diklik) */}
              {Object.keys(rowForm).length > 0 && (
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
                    <LF label="Tanggal Awal Lembur (opsional, bisa diisi nanti)">
                      <DateTxt style={iS} value={rowForm.tanggalAwal||""} onChange={(v:string)=>setRowForm((p:any)=>({...p,tanggalAwal:v}))}/>
                    </LF>
                    {(rowForm.jenisEntri||"kerja")==="ambil" && (
                      <LF label="Tanggal Akhir Lembur">
                        <DateTxt style={iS} value={rowForm.tanggalAkhir||""} onChange={(v:string)=>setRowForm((p:any)=>({...p,tanggalAkhir:v}))}/>
                      </LF>
                    )}
                    {(rowForm.jenisEntri||"kerja")==="kerja" ? (
                      <>
                        <LF label="Jam Absen Masuk (opsional, bisa diisi nanti)">
                          <input style={iS} type="time" value={rowForm.jamMasuk||""} onChange={e=>setRowForm((p:any)=>({...p,jamMasuk:e.target.value}))}/>
                        </LF>
                        {/* UPDATE LEMBUR GPS+SELFIE: input manual "Jam Absen Keluar" digantikan tombol
                            otomatisasi terintegrasi (kamera+GPS) demi mencegah manipulasi waktu.
                            Field rowForm.jamKeluar tetap sama, hanya cara pengisiannya yang berubah. */}
                        <div style={{gridColumn:"1 / -1",border:`1px dashed ${rowForm.jamKeluar?C.s:"#cbd5e1"}`,padding:14,borderRadius:10,background:rowForm.jamKeluar?C.sBg:"#f8fafc",marginTop:6}}>
                          <div style={{fontWeight:700,fontSize:12,marginBottom:8,color:C.s}}>
                            📸 VERIFIKASI SELESAI LEMBUR (GPS &amp; ANTI-MANIPULASI WAKTU)
                          </div>
                          <div style={{display:"flex",gap:12,alignItems:"center",marginBottom:8,flexWrap:"wrap"}}>
                            <Btn outline onClick={handleSelfieKeluarOtomatis} style={{background:C.pBg,color:C.p,fontSize:12}} disabled={isLoadingSelfie}>
                              {isLoadingSelfie ? "🔄 Mengunci Parameter Absen..." : "📷 Klik untuk Selfie & Ambil Jam Keluar"}
                            </Btn>
                            {rowForm.fotoKeluarUrl && (
                              <img
                                src={rowForm.fotoKeluarUrl}
                                alt="Preview Selfie"
                                style={{width:45,height:45,borderRadius:6,objectFit:"cover",border:`1px solid ${C.s}`}}
                              />
                            )}
                          </div>
                          {rowForm.jamKeluar && (
                            <div style={{fontSize:12,color:C.s,fontWeight:600,lineHeight:"1.4"}}>
                              ✓ Waktu Terkunci Pada Tanggal: {rowForm.tanggalAkhir?fDMY(rowForm.tanggalAkhir):"—"} | Jam: {rowForm.jamKeluar}<br/>
                              📍 Koordinat Lokasi: {rowForm.lokasiKeluar||"—"}
                            </div>
                          )}
                          {/* FALLBACK MANUAL: hanya tampil kalau GPS+kamera keduanya gagal/tidak aktif.
                              Kalau selfie berhasil → tersembunyi. Kalau gagal → muncul otomatis via
                              setShowManualKeluar(true) di handleSelfieKeluarOtomatis. */}
                          {showManualKeluar && (
                            <div style={{marginTop:10,paddingTop:10,borderTop:`1px dashed ${C.w}`}}>
                              <div style={{fontSize:11,color:"#92400e",background:"#FEF3C7",borderRadius:8,padding:"6px 10px",marginBottom:8,fontWeight:600}}>
                                ⚠ GPS/Kamera tidak aktif — isi jam keluar secara manual:
                              </div>
                              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                                <LF label="Tanggal Akhir Lembur">
                                  <DateTxt style={iS} value={rowForm.tanggalAkhir||""} onChange={(v:string)=>setRowForm((p:any)=>({...p,tanggalAkhir:v}))}/>
                                </LF>
                                <LF label="Jam Absen Keluar">
                                  <input style={iS} type="time" value={rowForm.jamKeluar||""} onChange={e=>setRowForm((p:any)=>({...p,jamKeluar:e.target.value}))}/>
                                </LF>
                              </div>
                            </div>
                          )}
                        </div>
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
                    <Btn full onClick={()=>saveEntry()} disabled={isLoadingSelfie} style={{flex:2}}>✓ Simpan</Btn>
                    <Btn full outline color={C.g} onClick={()=>{stopActiveSelfieStream();setIsLoadingSelfie(false);setShowManualKeluar(false);setEditRow(null);setRowForm({});}} style={{flex:1}}>Batal</Btn>
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
                    <div style={{fontSize:12,marginTop:4}}>Klik "✚ Tambah Baris Lembur" untuk mulai mencatat</div>
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
                              {/* FIX AUDIT #41 (MEDIUM — Nomor Urut Bentrok): kolom "No"
                                  SEBELUMNYA menampilkan field statis e.no yang dihitung sekali
                                  saat entri dibuat (entriesNow.length+1). Field ini bisa jadi
                                  tidak akurat/bentrok kalau ada race penulisan bersamaan atau
                                  urutan entries berubah (mis. setelah hapus baris tengah tanpa
                                  semua jalur ikut renumbering). Nomor urut murni tampilan —
                                  jadi diturunkan langsung dari index array saat render, yang
                                  SELALU konsisten & tidak pernah bentrok/berlubang. */}
                              <td style={{padding:"8px 10px",textAlign:"center",fontWeight:700,color:C.p}}>{i+1}</td>
                              <td style={{padding:"8px 10px",fontWeight:600,color:C.t}}>{fDMY(e.tanggalAwal||e.jadwalDinas)||"—"}</td>
                              <td style={{padding:"8px 10px",fontWeight:600,color:C.t}}>{e.tanggalAkhir?fDMY(e.tanggalAkhir):"—"}</td>
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
                                <button onClick={()=>{stopActiveSelfieStream();setIsLoadingSelfie(false);setQuickLemburMode(false);setEditRow(e.id);setRowForm({...e});}} style={{background:C.iBg,border:`1px solid ${C.i}`,borderRadius:6,color:C.i,padding:"3px 8px",cursor:"pointer",fontSize:11,marginRight:4}}>✏</button>
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
// MonitoringEntry & MonitoringCfg sekarang diimpor dari ./types (lihat import type di atas)
const defaultMonCfg: MonitoringCfg = { suhuMin:19, suhuMax:24, rhMin:45, rhMax:65, lokasiRuang:"Kamar Bedah", kepalaKamarBedah:"" };
// MON_JAMS dan MON_ROOMS sekarang diimpor dari ./constants (lihat import di atas)
/* monId: id deterministik berbasis (ruang,tanggal,jam) — bukan random — sehingga
   menyimpan ulang slot yang sama akan UPSERT (timpa baris yang sama) bukan
   membuat baris baru yang menumpuk sebagai duplikat "hantu" di Supabase.
   KRITIS: ruang disertakan agar data antar-ruangan tidak saling menimpa. */
const monId = (ruang: string, tanggal: string, jam: string) =>
  `mon_${ruang.replace(/\s+/g,"_")}_${tanggal}_${jam}`;
// MON_MONTHS_ID, MC, MON_ROOM_COLORS sekarang diimpor dari ./constants (lihat import di atas)

// monIsOK sekarang diimpor dari ./utils (lihat import di atas)

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
      sanitizeCellForExcel(e.petugas||""), // FIX AUDIT #33 (lanjutan): field bebas-teks
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

  /* Isi ulang form slots ketika tanggal berubah.
     CATATAN (setelah dikaji ulang): SENGAJA tidak menambahkan
     monitoringEntries ke deps meski itu berarti form tidak auto-refresh
     saat device lain menyimpan data monitoring pada tanggal yang sama.
     Alasannya: `slots` adalah DRAFT LOKAL yang sedang diketik user secara
     aktif (nilai suhu/kelembaban per slot belum tersimpan). Menambahkan
     monitoringEntries ke deps akan memicu mkSlots() dan MERESET SELURUH
     draft form — termasuk input yang sedang diketik user tapi belum
     diklik "Simpan" — setiap kali ADA SAJA perubahan monitoringEntries
     dari device lain, walau untuk ruangan/tanggal yang tidak terkait sama
     sekali. Risiko kehilangan input aktif ini jauh lebih berbahaya
     daripada risiko form menampilkan data sedikit usang. Trade-off ini
     disengaja — biarkan apa adanya. */
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
        headerRow.eachCell((cell: ExcelJS.Cell) => {
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
              <DateTxt value={formDate} onChange={(v:string)=>setFormDate(v)} style={{...iSm,width:"auto"}}/>
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
// SUPA_URL, SUPA_ANON, SUPA_CLIENT sekarang diimpor dari ./services (lihat import di atas)

/* ─── DROPBOX: token DIHAPUS dari sini sepenuhnya — lihat DBX_PROXY_SECRET
   di atas. Token asli sekarang hanya hidup sbg secret server-side di
   Supabase Edge Function "dropbox-proxy" (Deno.env.get("DROPBOX_ACCESS_TOKEN")). */

/* ─── HEADER CLOCK (terisolasi) ─────────────────────────────────────────
   Sengaja dipisah jadi komponennya sendiri — detak 1 detik di sini TIDAK
   memicu re-render App ataupun tab yang sedang aktif (lihat catatan perf
   di komponen App: state currentTime sebelumnya ada di root App). */
// HeaderClock sekarang diimpor dari ./components/header-clock (lihat import di atas)

/* ─── MAIN APP ───────────────────────────────────────────────────────── */

/* ─── MASTER DATA PERAWAT RS PANTI RINI ────────────────────────────── */
// MASTER_PERAWAT_LIST, SHIFT_CODES, KB_ROSTER_MASTER_TABLE, SEED_MASTER,
// BULAN_NAMA_ID, SUMMARY_COLS, LS_DRAFT_KEY sekarang diimpor dari
// ./constants (lihat import di atas)

/* ─── VIEW ROSTER GENERATOR ─────────────────────────────────────────── */
// NurseType, RosterNurse, RosterGenProps sekarang diimpor dari ./types (lihat import type di atas)

/* ── Master Perawat — interface (terpisah dari RosterNurse) ──
   FIX (Audit Sentralisasi DB): localStorage DIHAPUS TOTAL sebagai sumber
   data. Tabel Supabase "kb_roster_master" kini menjadi satu-satunya
   sumber kebenaran (1 row per perawat: id, data jsonb, updated_at),
   konsisten dengan pola tabel kb_* lain di file ini (kb_operasi,
   kb_staf, dst). SEED_MASTER hanya dipakai sekali untuk mengisi
   tabel Supabase bila tabel masih kosong (first-run), bukan sebagai
   fallback offline. */
// NurseMaster sekarang diimpor dari ./types (lihat import type di atas)

/* ── Helper: hitung ringkasan 7 kolom per perawat (untuk Excel export) ── */
function hitungRingkasan(row: string[], daysInMonth: number) {
  const cnt = (code: string) => row.slice(0, daysInMonth).filter(c=>c===code).length;
  const P=cnt("P"), PG=cnt("PG"), S=cnt("S"), SG=cnt("SG"),
        T=cnt("T"), LG=cnt("LG"), L=cnt("L");
  return [P, PG, S, SG, LG+L, T, P+PG+S+SG+T+LG+L];
}

function mkNurse(tipe: NurseType = "junior"): RosterNurse {
  return { id: gId(), nik:"", pk:"", name:"", sisaCuti: 12, tipe };
}

function fisherYates<T>(arr: T[]): T[] {
  const a = [...arr];
  for(let i=a.length-1;i>0;i--){ const j=Math.floor(Math.random()*(i+1)); [a[i],a[j]]=[a[j],a[i]]; }
  return a;
}

function ViewRosterGenerator({ showToast, upsertOneToSupa, deleteFromSupa, dbxCfg }: RosterGenProps) {
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

  /* ── FITUR 1: State Master Perawat — SEPENUHNYA dari Supabase ──
     FIX (Audit Sentralisasi DB): localStorage dihapus total. Saat mount,
     fetch seluruh baris tabel kb_roster_master. Jika tabel masih kosong
     (deployment baru), seed sekali dengan SEED_MASTER lalu upload ke
     Supabase — setelah itu Supabase menjadi satu-satunya sumber data,
     sama seperti modul kb_operasi/kb_staf/dst. */
  const [nurseMasterList, setNurseMasterList] = useState<NurseMaster[]>([]);
  const [masterForm, setMasterForm] = useState({ nik:"", name:"", pk:"III" });
  const [masterFormErr, setMasterFormErr] = useState("");
  const [masterSaving, setMasterSaving]   = useState(false);
  const [masterLoading, setMasterLoading] = useState(true);

  /* Fetch master perawat dari Supabase sekali saat komponen mount */
  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const { data, error } = await SUPA_CLIENT.from(KB_ROSTER_MASTER_TABLE).select("data").order("updated_at",{ascending:true});
        if(cancelled) return;
        if(!error && data?.length){
          setNurseMasterList(data.map((x:any)=>x.data));
        } else if(!error && data?.length===0){
          /* Tabel kosong (first-run) → seed sekali dari SEED_MASTER */
          setNurseMasterList(SEED_MASTER);
          const ts = new Date().toISOString();
          await SUPA_CLIENT.from(KB_ROSTER_MASTER_TABLE).upsert(
            SEED_MASTER.map(m=>({ id:m.id, data:{...m, updated_at:ts}, updated_at:ts })),
            {onConflict:"id"}
          );
        } else if(error){
          console.warn("[kb_roster_master] fetch error:", formatSupaError(error));
          showToast("⚠ Gagal memuat data master dari Supabase", "#DC2626");
        }
      } catch(e) {
        if(!cancelled) showToast("⚠ Koneksi ke Supabase gagal saat memuat data master", "#DC2626");
      } finally {
        if(!cancelled) setMasterLoading(false);
      }
    })();
    return () => { cancelled = true; };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  /* FIX TDZ: newGrid_safe dipindah ke sini (sebelum rowSummaries/colSummaries)
     karena kedua useMemo di bawah memanggilnya saat render pertama — definisi
     const yang berada di bawah useMemo menyebabkan ReferenceError TDZ. */
  const newGrid_safe = (ni:number, d:number) => grid[ni]?.[d] ?? "";

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

  /* ── FITUR 1: CRUD Master Perawat — sentralisasi penuh ke Supabase ──
     Setiap perawat = 1 row di tabel kb_roster_master (id, data, updated_at),
     memakai upsertOneToSupa/deleteFromSupa yang sudah ada — pola yang sama
     dipakai modul kb_staf, kb_operasi, dst. Tidak ada lagi localStorage. */
  const addToMaster = async () => {
    const nik  = masterForm.nik.trim();
    const name = masterForm.name.trim();
    const pk   = masterForm.pk.trim();
    if(!nik)  { setMasterFormErr("NIK tidak boleh kosong."); return; }
    if(!name) { setMasterFormErr("Nama tidak boleh kosong."); return; }
    if(nurseMasterList.some(m => m.nik === nik))
      { setMasterFormErr(`NIK ${nik} sudah ada di database.`); return; }
    const newEntry: NurseMaster = { id: gId(), nik, name, pk };
    setMasterFormErr("");
    setMasterSaving(true);
    try {
      const res = await upsertOneToSupa(KB_ROSTER_MASTER_TABLE, {...newEntry});
      if(res.ok){
        setNurseMasterList(p => [...p, newEntry]);
        setMasterForm({ nik:"", name:"", pk:"III" });
        showToast(`✅ "${name}" (NIK: ${nik}) disimpan ke Supabase!`, "#2E7D32");
      } else {
        showToast(`✗ Gagal menyimpan ke Supabase: ${res.error||"unknown error"}`, "#DC2626");
      }
    } catch(e:any) {
      showToast(`✗ Koneksi ke Supabase gagal: ${e?.message||"unknown error"}`, "#DC2626");
    }
    setMasterSaving(false);
  };

  const removeFromMaster = async (id: string) => {
    const target = nurseMasterList.find(m => m.id === id);
    if(!target) return;
    try {
      const res = await deleteFromSupa(KB_ROSTER_MASTER_TABLE, id);
      if(res.ok){
        setNurseMasterList(p => p.filter(m => m.id !== id));
        showToast(`🗑 Perawat "${target.name}" dihapus dari database master.`, "#5C677D");
      } else {
        showToast(`✗ Gagal menghapus dari Supabase: ${res.error||"unknown error"}`, "#DC2626");
      }
    } catch(e:any) {
      showToast(`✗ Koneksi ke Supabase gagal: ${e?.message||"unknown error"}`, "#DC2626");
    }
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
      const res = await dropboxUploadBytes(path, await wb.xlsx.writeBuffer() as unknown as ArrayBuffer);
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
            ☁ Supabase (real-time)
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
        {masterLoading ? (
          <div style={{fontSize:12,color:C.tL,textAlign:"center",padding:"10px 0"}}>⟳ Memuat data master dari Supabase...</div>
        ) : nurseMasterList.length > 0 ? (
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
        ) : (
          <div style={{fontSize:12,color:C.tL,textAlign:"center",padding:"10px 0"}}>Belum ada data master perawat. Tambahkan lewat form di atas.</div>
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

// ALL_TABS sekarang diimpor dari ./constants (lihat import di atas)

/* PENAMBAHAN #6 (konsistensi ikon): set ikon SVG line-style (stroke-based,
   konsisten strokeWidth) untuk tab sidebar/nav — pengganti emoji yang
   sebelumnya dipakai (🏠📋📝 dst). Dibuat inline (BUKAN import library
   seperti lucide-react) supaya TIDAK menambah dependency baru yang perlu
   di-install — cukup tempel file ini dan langsung jalan. CATATAN JUJUR:
   ini baru diterapkan di tab sidebar (elemen paling sering dilihat),
   BELUM di seluruh app (ratusan emoji lain di form/pesan toast/dll masih
   emoji) — mengganti semuanya butuh sesi terpisah karena scope-nya besar
   dan risiko regresi visual kalau dikerjakan tergesa. */
// TabIcon dan ActionIcon sekarang diimpor dari ./components/icons (lihat import di atas)

// ===== Install Prompt PWA (Android/desktop pakai event browser, iOS pakai instruksi manual) =====
// BeforeInstallPromptEvent sekarang diimpor dari ./types (lihat import type di atas)

// KB_INSTALL_DISMISS_KEY dan KB_INSTALL_DISMISS_MS sekarang diimpor dari ./constants (lihat import di atas)

// kbIsStandalone, kbIsIos, kbWasDismissedRecently, InstallPrompt sekarang
// diimpor dari ./components/install-prompt (lihat import di atas)
// ===== End Install Prompt =====

/* ═══════════════════════════════════════════════════════════════════
   VIRTUAL ASSISTANT — widget chat melayang untuk pendaftaran pasien lewat
   teks natural (mis. "Ny. Siti 45th appendektomi besok jam 9 dr. Andika").

   SENGAJA TIDAK bikin endpoint/tabel/skema baru — reuse penuh 2 jalur
   yang SUDAH ADA & sudah teruji:
   1. Ekstraksi teks → runWorkerOcrScanText() → endpoint /api/ocr-scan
      YANG SAMA dipakai jalur "paste teks" di Tab Daftar. worker.ts TIDAK
      perlu diubah sama sekali untuk fitur ini.
   2. Penyimpanan → autoSaveOcrBatch() YANG SAMA dipakai batch OCR foto —
      otomatis tersimpan ke kb_operasi dengan needsReview:true. Efeknya:
      entri dari asisten ini muncul di Tab Jadwal dengan badge biru
      "🔍 Perlu Review (dari OCR)" — perawat WAJIB buka Edit untuk
      cek/lengkapi lalu Save (itu yang berfungsi sebagai "ACC"). TIDAK
      ADA jalur auto-commit data pasien tanpa direview manusia — sengaja,
      karena salah simpan jadwal operasi berisiko fatal (lihat diskusi
      sebelumnya). Pola "antrian ACC" ini konsisten dengan yang sudah
      berjalan untuk OCR foto/teks, bukan mekanisme baru yang terpisah. */
// AssistantMsg sekarang diimpor dari ./types (lihat import type di atas)

// FloatingAssistant sekarang diimpor dari ./components/floating-assistant (lihat import di atas)

export default function App() {
  useEffect(() => { ensureCustomFontLink(); }, []);
  /* FIX AUDIT #23 (MEDIUM — Silent Failure jika env var Supabase hilang):
     sebelumnya kegagalan ini hanya di-console.error (tidak terlihat user),
     lalu aplikasi tetap render normal — perawat bisa login & mengisi form
     tanpa tahu bahwa SEMUA operasi simpan akan gagal karena tidak pernah
     ada koneksi ke database sama sekali. Untuk sistem medis, "silent
     failure" seperti ini berbahaya: perawat mengira data tersimpan padahal
     tidak. Sekarang aplikasi berhenti di layar blocking yang jelas,
     mencegah interaksi lebih jauh sampai konfigurasi diperbaiki. */
  /* ⚠️ CATATAN RULES OF HOOKS: early-return ini AMAN karena SUPA_URL/SUPA_ANON
     adalah module-level const yang dihitung SEKALI saat file di-load dan
     TIDAK PERNAH berubah selama sesi browser berjalan — sehingga kondisi ini
     selalu identik di setiap render (tidak akan pernah mengubah jumlah/urutan
     hooks di bawahnya). JANGAN tiru pola early-return-sebelum-hooks ini untuk
     kondisi yang BISA berubah antar render (state/props) — itu akan
     melanggar Rules of Hooks dan menyebabkan bug "Rendered more hooks than
     during the previous render". */
  if (!SUPA_URL || !SUPA_ANON) {
    return (
      <div style={{minHeight:"100vh",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:24,background:"#FEF2F2",textAlign:"center"}}>
        <div style={{fontSize:40,marginBottom:12}}>⚠️</div>
        <div style={{fontSize:16,fontWeight:800,color:"#991B1B",marginBottom:8}}>Konfigurasi Server Belum Lengkap</div>
        <div style={{fontSize:13,color:"#7F1D1D",maxWidth:420,lineHeight:1.5}}>
          Aplikasi tidak dapat terhubung ke database (VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY belum diset).
          Semua data TIDAK AKAN TERSIMPAN sampai konfigurasi ini diperbaiki oleh admin/IT.
          Hubungi tim IT sebelum menggunakan aplikasi ini.
        </div>
      </div>
    );
  }
  const [pinOK,    setPinOK]    = useState(false);
  const [role,     setRole]     = useState<"admin"|"perawat">("perawat");
  const [pinAdmin,   setPinAdmin]   = useState<string>("2409"); // fallback
  const [pinPerawat, setPinPerawat] = useState<string>("1234"); // fallback
  const [pinLoaded,  setPinLoaded]  = useState(false);
  const [pinFromCloud, setPinFromCloud] = useState(false); // true jika PIN loaded dari Supabase
  const isFirstTime = pinLoaded && !pinFromCloud; // first time jika Supabase tidak punya PIN
  const [tab,      setTab]      = useState("home");
  /* MENU INPUT LEMBUR KHUSUS: jika true, perawat login lewat tombol
     "⏰ Input Lembur" di layar PIN, sehingga setelah login HANYA tab Lembur
     yang tersedia (tanpa tab lain) — mempermudah perawat yang cuma perlu
     mencatat lembur tanpa harus melewati semua halaman. Login biasa (tombol
     Masuk biasa) tetap membuka semua tab seperti biasa. */
  const [lemburOnlyMode, setLemburOnlyMode] = useState(false);

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
  /* PENAMBAHAN #3 (skeleton loading): true sampai fetch data AWAL dari
     Supabase (loadAllFromSupa) selesai — dipakai ViewJadwal untuk
     membedakan "data memang kosong" vs "masih dimuat" (lihat komentar di
     ViewJadwalProps.initialLoading). Diset false di FINALLY loadAllFromSupa
     (baik sukses maupun gagal) supaya tidak pernah macet true selamanya
     kalau fetch gagal/offline. */
  const [initialLoading, setInitialLoading] = useState(true);
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
  /* OPSI A: state untuk auto-pilih pegawai di ViewLembur setelah login via
     "⏰ Input Lembur" — diset saat user memilih nama di PinScreen, diteruskan
     ke ViewLembur sebagai initialSelPeg, lalu dikosongkan setelah dipakai */
  const [autoSelPeg, setAutoSelPeg] = useState<string>("");
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
    const run24h = async (currentOps: any[], currentLemburPegawai: any[], currentLemburData: Record<string,any>, currentCfg: DropboxConfig, currentMonitoringEntries: MonitoringEntry[], currentMonitoringCfg: any) => {
      const now = Date.now();
      const last = currentCfg.lastExcelBackupTs||0;
      if(now - last < 24*60*60*1000) return; /* 24h not yet elapsed */
      const folder = currentCfg.path.replace(/[^/]+$/, "");
      /* FIX AUDIT #22 (LOW/MEDIUM — Pergeseran Tanggal Timezone di nama file):
         sama seperti FIX AUDIT #21 — toISOString() memakai UTC, sehingga
         backup yang dibuat pukul 00:00–06:59 WIB diberi nama tanggal
         KEMARIN. Diganti ke toLocalISODate() supaya nama file backup selalu
         konsisten dengan tanggal kalender WIB yang dilihat admin. */
      const stamp = toLocalISODate(new Date());
      /* — Operasi Excel — */
      const wbOps = XLSX.utils.book_new();
      const h = ["No","Tanggal","Jam","Pasien","Usia","RM","Jenis","Diagnosa","Tindakan","Kamar","Operator","Anestesi","Asisten","Instrumen","P.Anestesi","Onloop","RR/Katim","Status","Dibuat",...KOLOM_TAMBAHAN_OPS];
      const rows = [...currentOps].sort((a:any,b:any)=>a.date.localeCompare(b.date)||a.time.localeCompare(b.time)).map((o:any,i:number)=>[i+1,o.date,o.time,sanitizeCellForExcel(o.patient||""),o.age||"",o.rm||"",o.opType||"",sanitizeCellForExcel(o.diagnosis||""),sanitizeCellForExcel(o.procedure||""),o.room||"",sanitizeCellForExcel(o.surgeon||""),sanitizeCellForExcel(o.anesthesiologist||""),sanitizeCellForExcel(o.assistantNurse||""),sanitizeCellForExcel(o.circulatingNurse||""),sanitizeCellForExcel(o.anesthesiaNurse||""),sanitizeCellForExcel(o.onloopNurse||""),sanitizeCellForExcel(o.rrKatim||""),o.status||"",o.createdAt||"",...nilaiTambahanOps(o)]);
      const wsOps = XLSX.utils.aoa_to_sheet([h,...rows]);
      XLSX.utils.book_append_sheet(wbOps, wsOps, "Jadwal Operasi");
      const resOps = await dropboxUploadExcel(`${folder}Auto_Operasi_${stamp}.xlsx`, wbOps);
      /* — Lembur Excel — */
      const wbLmb = XLSX.utils.book_new();
      currentLemburPegawai.forEach((p:any)=>{
        const entries: any[][]=[];
        Object.keys(currentLemburData).filter(k=>k.startsWith(p.id+"_")).sort().forEach(k=>{
          const rec=currentLemburData[k];
          /* UPDATE LEMBUR GPS+SELFIE: kolom lokasiKeluar & fotoKeluarUrl ditambahkan di ujung kanan,
             dengan fallback "—" untuk data lama agar tidak ada baris yang error/kosong */
          (rec.entries||[]).forEach((e:any)=>entries.push([sanitizeCellForExcel(p.name),p.nik||"",k.replace(p.id+"_",""),e.tanggalAwal||"",e.tanggalAkhir||"",e.jamMasuk||"",e.jamKeluar||"",sanitizeCellForExcel(e.keperluanLembur||""),sanitizeCellForExcel(e.keterangan||""),sanitizeCellForExcel(e.ttd||""),e.lokasiKeluar||"—",e.fotoKeluarUrl||"—"]));
        });
        if(entries.length){const ws=XLSX.utils.aoa_to_sheet([["Nama","NIK","Bulan","Tgl Awal","Tgl Akhir","Jam Masuk","Jam Keluar","Keperluan","Keterangan","TTD","Lokasi Keluar (GPS)","Tautan Foto Selfie"],...entries]);XLSX.utils.book_append_sheet(wbLmb,ws,p.name.slice(0,28));}
      });
      if(!wbLmb.SheetNames.length) XLSX.utils.book_append_sheet(wbLmb,XLSX.utils.aoa_to_sheet([["Belum ada data lembur"]]),"Data Lembur");
      const resLmb = await dropboxUploadExcel(`${folder}Auto_Lembur_${stamp}.xlsx`, wbLmb);
      /* — Monitoring Excel —
         FIX AUDIT #43 (MEDIUM/HIGH — Stale Closure pada Backup Otomatis 24 Jam):
         SEBELUMNYA bagian ini memakai `monitoringEntries`/`monitoringCfg` langsung
         dari closure render — BUKAN dari latestDataRef.current seperti currentOps/
         currentLemburPegawai/currentLemburData di atas. Karena efek ini hanya
         re-run saat `dbxCfg.autoExcelBackup` berubah (dependency array di baris
         paling bawah useEffect ini), closure atas monitoringEntries/monitoringCfg
         "membeku" pada nilai saat auto-backup terakhir kali diaktifkan/nonaktifkan
         — bukan pada data terkini. Akibatnya: setiap kali setInterval memanggil
         run24h() sejam sekali selama berhari-hari, sheet "Monitoring Suhu" &
         "Konfigurasi" di file Auto_Monitoring_*.xlsx akan TERUS berisi data lama
         (suhu/kelembaban yang dicatat SAAT toggle diaktifkan), padahal sheet
         Operasi & Lembur di file lain sudah benar memakai data terbaru — inkonsisten
         dan berpotensi membuat admin salah kira suhu/kelembaban ruangan sudah
         lama tidak sesuai standar padahal datanya memang tidak pernah ter-update.
         Sekarang currentMonitoringEntries/currentMonitoringCfg diteruskan sebagai
         parameter dari latestDataRef.current (sama seperti data lain), sehingga
         selalu segar setiap kali run24h dipanggil. */
      const wbMon = XLSX.utils.book_new();
      const monRows = currentMonitoringEntries.map((e:MonitoringEntry)=>[e.ruang??currentMonitoringCfg.lokasiRuang??"",e.tanggal,e.jam,e.suhu,e.kelembaban,monIsOK(e.suhu,e.kelembaban,currentMonitoringCfg)?"SESUAI":"TIDAK SESUAI",e.petugas]);
      const wsMonMain = XLSX.utils.aoa_to_sheet([["Ruangan","Tanggal","Jam","Suhu (°C)","Kelembaban (%)","Status","Petugas"],...monRows]);
      wsMonMain["!cols"]=[{wch:18},{wch:12},{wch:8},{wch:12},{wch:14},{wch:16},{wch:22}];
      XLSX.utils.book_append_sheet(wbMon, wsMonMain, "Monitoring Suhu");
      const wsMonCfg = XLSX.utils.aoa_to_sheet([["Parameter","Nilai"],["Ruangan 1","Kamar Bedah 1"],["Ruangan 2","Kamar Bedah 2"],["Ruangan 3","Ruang Instrumen"],["Suhu Min",currentMonitoringCfg.suhuMin],["Suhu Max",currentMonitoringCfg.suhuMax],["RH Min",currentMonitoringCfg.rhMin],["RH Max",currentMonitoringCfg.rhMax],["Kepala KB",currentMonitoringCfg.kepalaKamarBedah]]);
      XLSX.utils.book_append_sheet(wbMon, wsMonCfg, "Konfigurasi");
      const resMon = await dropboxUploadExcel(`${folder}Auto_Monitoring_${stamp}.xlsx`, wbMon);
      if(resOps.ok&&resLmb.ok&&resMon.ok){
        setDbxCfg((p:DropboxConfig)=>({...p,lastExcelBackup:fNow(),lastExcelBackupTs:Date.now()}));
      }
    };
    /* Check immediately on enable, then every hour */
    run24h(latestDataRef.current.ops||[], latestDataRef.current.lemburPegawai||[], latestDataRef.current.lemburData||{}, dbxCfg, latestDataRef.current.monitoringEntries||[], latestDataRef.current.monitoringCfg||{});
    dbxExcelAutoRef.current = setInterval(()=>{
      setDbxCfg(p=>{ run24h(latestDataRef.current.ops||[], latestDataRef.current.lemburPegawai||[], latestDataRef.current.lemburData||{}, p, latestDataRef.current.monitoringEntries||[], latestDataRef.current.monitoringCfg||{}); return p; });
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
      /* PENAMBAHAN #1 (micro-animation): transisi halus saat pindah tab —
         fade-in + sedikit slide-up, supaya perpindahan Home→Jadwal→Daftar
         dst terasa hidup, tidak "melompat" mendadak seperti sebelumnya. */
      /* PERKUAT #5 (Fade/Blur Content, mirip "FadeContent" ReactBits):
         ditambah blur() di awal transisi, bukan cuma opacity+translateY —
         kesan konten "menajam ke fokus" saat muncul, bukan sekadar fade. */
      @keyframes kbTabFadeIn{from{opacity:0;transform:translateY(6px);filter:blur(4px);}to{opacity:1;transform:translateY(0);filter:blur(0);}}
      .kb-content-inner{animation:kbTabFadeIn .28s ease-out;}
      /* Kartu (Card) sedikit terangkat saat disentuh/hover — memberi kesan
         interaktif tanpa mengubah layout/struktur data apa pun. */
      .kb-content-inner .kbCardHoverable{transition:transform .15s ease,box-shadow .15s ease;}
      .kb-content-inner .kbCardHoverable:active{transform:scale(.985);}

      /* PENAMBAHAN — efek visual "website profesional" (batch ke-2) */

      /* #3 Ripple effect: pakai trik CSS radial-gradient di ::after, tanpa
         perlu JS/tracking posisi klik — otomatis berlaku ke SEMUA <button>
         di seluruh app (global selector), tidak perlu edit tiap tombol. */
      button{position:relative;overflow:hidden;}
      button::after{content:"";position:absolute;inset:0;border-radius:inherit;
        background:radial-gradient(circle,rgba(255,255,255,.55) 0%,rgba(255,255,255,0) 70%);
        opacity:0;transform:scale(.4);transition:opacity .5s ease,transform .5s ease;pointer-events:none;}
      button:active::after{opacity:1;transform:scale(1.6);transition:opacity 0s,transform 0s;}

      /* #8 Input focus glow: transisi warna halus + shadow lembut saat
         fokus, berlaku global ke semua input/select/textarea. */
      input,select,textarea{transition:border-color .18s ease,box-shadow .18s ease;}
      input:focus,select:focus,textarea:focus{
        border-color:#3B82F6!important;
        box-shadow:0 0 0 3px rgba(59,130,246,.18)!important;
        outline:none;
      }

      /* #9 Loading bar tipis di atas layar (indeterminate progress),
         dipicu oleh state global topLoading — lihat App() untuk trigger. */
      @keyframes kbTopBarSlide{0%{left:-30%;width:30%;}50%{left:40%;width:45%;}100%{left:100%;width:30%;}}
      .kb-topbar{position:fixed;top:0;left:0;right:0;height:3px;z-index:99999;overflow:hidden;pointer-events:none;background:transparent;}
      .kb-topbar span{position:absolute;top:0;height:100%;background:linear-gradient(90deg,#38bdf8,#1d4ed8);
        border-radius:2px;animation:kbTopBarSlide 1.1s ease-in-out infinite;box-shadow:0 0 8px rgba(59,130,246,.6);}

      /* #7 Toast: masuk dengan sedikit "bounce" (spring-like), keluar
         fade+turun halus. */
      @keyframes kbToastIn{
        0%{opacity:0;transform:translate(-50%,16px) scale(.92);}
        60%{opacity:1;transform:translate(-50%,-4px) scale(1.02);}
        100%{opacity:1;transform:translate(-50%,0) scale(1);}
      }
      .kb-toast{animation:kbToastIn .38s cubic-bezier(.34,1.56,.64,1);}

      /* PENAMBAHAN — Modal transition: backdrop fade-in halus, konten
         modal scale+fade dari 96% ke 100% (bukan muncul instan). */
      @keyframes kbModalBackdropIn{from{opacity:0;}to{opacity:1;}}
      @keyframes kbModalContentIn{from{opacity:0;transform:scale(.96) translateY(8px);}to{opacity:1;transform:scale(1) translateY(0);}}
      .kb-modal-backdrop{animation:kbModalBackdropIn .2s ease-out;}
      .kb-modal-content{animation:kbModalContentIn .25s cubic-bezier(.16,1,.3,1);}

      /* #6 Gradient border animasi untuk elemen CYTO (urgent) — border
         "berjalan" mengelilingi card, menarik perhatian tanpa mengganggu
         konten di dalamnya. PERJELASAN: ini padanan/rekonstruksi dari
         komponen "Star Border" ReactBits (border berkilau berputar) —
         sudah diterapkan sejak batch efek visual sebelumnya, dipakai
         KHUSUS untuk item CYTO (operasi darurat) di Tab Jadwal karena
         butuh menarik perhatian secara visual, bukan dekorasi semata. */
      @keyframes kbGradientBorderSpin{to{--kb-angle:360deg;}}
      @property --kb-angle{syntax:'<angle>';initial-value:0deg;inherits:false;}
      .kb-cyto-glow{
        position:relative;border-radius:14px;
        background:linear-gradient(#fff,#fff) padding-box,
          conic-gradient(from var(--kb-angle),#B71C1C,#FF8A65,#FFC107,#B71C1C) border-box;
        border:2.5px solid transparent;
        animation:kbGradientBorderSpin 2.5s linear infinite;
      }

      /* #10 Noise/grain texture overlay — tekstur film grain SANGAT halus
         di seluruh app (mirip trend "grainy background" di produk modern
         2025-2026), memberi kesan "didesain" tanpa mengganggu keterbacaan.
         Dibuat via SVG data-uri (feTurbulence), TANPA file gambar eksternal. */
      .kb-noise-overlay{
        position:fixed;inset:0;z-index:1;pointer-events:none;opacity:.035;mix-blend-mode:overlay;
        background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='120' height='120'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.85' numOctaves='3' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23n)'/%3E%3C/svg%3E");
      }

      /* #2 Skeleton shimmer diperhalus: arah & timing lebih natural */
      .kb-skeleton-bar{animation-timing-function:cubic-bezier(.4,0,.6,1)!important;}

      /* PENAMBAHAN #5 (tipografi modern): kontras berat font yang lebih
         tegas antara judul (tebal, letter-spacing rapat) dan body text
         (lebih ringan, line-height lega) — ini yang membentuk hierarki
         visual "modern", bukan sekadar ukuran font yang beda-beda. */
      /* Font kustom (Manrope/Playfair Display) di-load lewat <link> yang
         disuntikkan via useEffect terpisah (lihat ensureCustomFontLink()) —
         BUKAN via @import di sini, karena @import HARUS jadi aturan
         PERTAMA dalam sebuah stylesheet atau browser mengabaikannya diam-
         diam; di tengah blok CSS besar seperti ini, @import tidak akan
         pernah berfungsi. */
      /* PENAMBAHAN — kesan "mewah": font kustom (bukan font sistem default
         yang jadi ciri khas tampilan "belum di-brand"). Manrope untuk body
         (rapi, modern), Playfair Display (serif elegan) khusus untuk
         judul-judul besar/hero — kombinasi umum di brand premium. */
      body,input,select,textarea,button{font-family:'Manrope',-apple-system,BlinkMacSystemFont,sans-serif!important;}
      .kb-font-display{font-family:'Playfair Display',Georgia,serif!important;}
      body{-webkit-font-smoothing:antialiased;text-rendering:optimizeLegibility;}
      /* Custom scrollbar — tipis, warna senada tema, rounded (scrollbar
         browser default terasa "murahan"/generik). */
      ::-webkit-scrollbar{width:8px;height:8px;}
      ::-webkit-scrollbar-track{background:transparent;}
      ::-webkit-scrollbar-thumb{background:rgba(11,18,32,.22);border-radius:8px;}
      ::-webkit-scrollbar-thumb:hover{background:rgba(11,18,32,.38);}
      *{scrollbar-width:thin;scrollbar-color:rgba(11,18,32,.22) transparent;}
      .kb-content-inner{line-height:1.5;}
      .kb-content-inner b, .kb-content-inner strong{letter-spacing:-.01em;}

      /* ── Mobile default (<640px) ── */
      body{background:#f8fafc;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;}
      .kbShell{width:100%;min-height:100vh;}
      .kb-header{padding:12px 16px 0;}
      .kb-header-inner{margin-bottom:10px;}
      .kb-tabs{display:flex;overflow-x:auto;-webkit-overflow-scrolling:touch;gap:0;padding-bottom:0;}
      .kb-tab-btn{padding:10px 14px;flex-shrink:0;white-space:nowrap;font-size:15px;font-weight:700;
        transition:transform .18s cubic-bezier(.34,1.56,.64,1);}
      .kb-tab-btn:active{transform:scale(1.18) translateY(-2px);}
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
          background:linear-gradient(180deg,#0B1220,#151F35 60%,#0B1220);
          padding:16px 10px;gap:4px;position:relative;
          border-right:1px solid rgba(232,206,143,.14);
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
        .kb-sidebar-btn.active{background:rgba(232,206,143,.14);color:#fff;font-weight:900;box-shadow:inset 3px 0 0 #E8CE8F;}
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

  /* REVISI: sebelumnya lemburPegawai HANYA di-load lewat loadAllFromSupa()
     yang baru jalan setelah pinOK (setelah login berhasil). Akibatnya saat
     PinScreen pertama kali dibuka (sebelum login), dropdown nama pegawai di
     alur "⏰ Input Lembur" selalu kosong dan memunculkan pesan keliru
     "Belum ada data pegawai" — padahal datanya memang ada di Supabase,
     hanya belum sempat di-fetch. Di sini kita fetch kb_lembur_pegawai saja
     (ringan, tanpa data operasi/pasien) SEKALI saat aplikasi pertama kali
     dimuat, sebelum login, supaya dropdown nama langsung terisi. */
  useEffect(()=>{
    if(pinOK) return; // sudah login → biarkan loadAllFromSupa yang urus
    (async ()=>{
      try{
        const {data} = await SUPA_CLIENT.from("kb_lembur_pegawai").select("data").order("updated_at",{ascending:false});
        if(data?.length) setLemburPegawai(data.map((x:any)=>x.data));
      } catch { /* offline — PinScreen tetap tampilkan pesan fallback apa adanya */ }
    })();
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[]);

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
    const applyDelta = (table: string, payload: RtPayload) => {
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
    } finally {
      /* PENAMBAHAN #3 (skeleton loading): SELALU diset false di sini —
         baik fetch sukses maupun gagal — supaya skeleton tidak pernah
         macet tampil selamanya kalau Supabase gagal/offline saat login. */
      setInitialLoading(false);
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

  /* PENAMBAHAN #8 (offline queue — flush otomatis): begitu browser
     mendeteksi koneksi kembali online, coba sinkronkan semua entri
     kb_operasi yang tertunda di localStorage (diisi dari saveOpFn /
     autoSaveOcrBatch saat offline — lihat enqueueOfflineOp). Item yang
     berhasil dihapus dari antrean; yang masih gagal (mis. koneksi putus
     lagi di tengah proses) tetap di antrean untuk dicoba lagi nanti. */
  const flushOfflineQueueRef = useRef(false);
  const flushOfflineQueue = useCallback(async () => {
    if (flushOfflineQueueRef.current) return; // cegah flush tumpang tindih
    const pending = readOfflineQueue();
    if (pending.length === 0) return;
    flushOfflineQueueRef.current = true;
    try {
      showToast(`🔄 Menyinkronkan ${pending.length} data yang tersimpan offline…`, C.tL);
      let successCount = 0;
      for (const item of pending) {
        const res = await upsertOneToSupa("kb_operasi", item.payload);
        if (res?.ok) { dequeueOfflineOp(item.id); successCount++; }
      }
      const remaining = countOfflineQueue();
      if (successCount > 0) {
        showToast(
          remaining > 0
            ? `✓ ${successCount} data tersinkron. ${remaining} masih tertunda (akan dicoba lagi).`
            : `✓ Semua ${successCount} data offline berhasil tersinkron ke database.`,
          C.s
        );
      }
    } finally {
      flushOfflineQueueRef.current = false;
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [upsertOneToSupa, showToast]);

  // Coba flush saat mount (kalau ada sisa antrean dari sesi sebelumnya
  // yang belum sempat sinkron) DAN setiap kali koneksi kembali online.
  useEffect(() => {
    if (navigator.onLine) flushOfflineQueue();
    window.addEventListener("online", flushOfflineQueue);
    return () => window.removeEventListener("online", flushOfflineQueue);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  /* Keep latestDataRef fresh for use in intervals (prevents stale closures) */
  useEffect(()=>{
    latestDataRef.current = {ops, archive, notifs, lemburPegawai, lemburData, monitoringEntries, monitoringCfg, staff, roster};
  },[ops, archive, notifs, lemburPegawai, lemburData, monitoringEntries, monitoringCfg, staff, roster]);

  const handlePinVerify = (newRole: "admin"|"perawat", newAdminPin?: string, newPerawatPin?: string, quickLembur?: boolean, pegawaiId?: string) => {
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
    /* Jika login lewat tombol "⏰ Input Lembur": kunci ke tab lembur saja.
       Login normal: pastikan mode ini ter-reset (jaga-jaga sesi sebelumnya). */
    setLemburOnlyMode(!!quickLembur);
    /* OPSI A: simpan pilihan pegawai dari PinScreen agar ViewLembur bisa
       auto-set selPeg tanpa user perlu pilih dropdown lagi */
    if(pegawaiId) setAutoSelPeg(pegawaiId);
    setTab(quickLembur ? "lembur" : "home");
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
    const h = ["No","Tanggal","Jam","Pasien","Usia","RM","Jenis","Diagnosa","Tindakan","Kamar","Operator","Anestesi","Asisten","Instrumen","P.Anestesi","Onloop","RR/Katim","Status","Dibuat",...KOLOM_TAMBAHAN_OPS];
    /* FIX AUDIT #33 (HIGH, lanjutan): sanitasi konsisten dengan export lain agar
       tidak jadi celah formula injection yang luput di path backup manual ini. */
    const rows = [...o].sort((a:any,b:any)=>a.date.localeCompare(b.date)||a.time.localeCompare(b.time)).map((op:any,i:number)=>[i+1,op.date,op.time,sanitizeCellForExcel(op.patient||""),op.age||"",op.rm||"",op.opType||"",sanitizeCellForExcel(op.diagnosis||""),sanitizeCellForExcel(op.procedure||""),op.room||"",sanitizeCellForExcel(op.surgeon||""),sanitizeCellForExcel(op.anesthesiologist||""),sanitizeCellForExcel(op.assistantNurse||""),sanitizeCellForExcel(op.circulatingNurse||""),sanitizeCellForExcel(op.anesthesiaNurse||""),sanitizeCellForExcel(op.onloopNurse||""),sanitizeCellForExcel(op.rrKatim||""),op.status||"",op.createdAt||"",...nilaiTambahanOps(op)]);
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
        /* UPDATE LEMBUR GPS+SELFIE: kolom lokasiKeluar & fotoKeluarUrl ditambahkan di ujung kanan */
        (rec.entries||[]).forEach((e:any)=>pegEntries.push([sanitizeCellForExcel(p.name),p.nik||"",k.replace(p.id+"_",""),e.tanggalAwal||"",e.tanggalAkhir||"",e.jamMasuk||"",e.jamKeluar||"",sanitizeCellForExcel(e.keperluanLembur||""),sanitizeCellForExcel(e.keterangan||""),sanitizeCellForExcel(e.ttd||""),e.lokasiKeluar||"—",e.fotoKeluarUrl||"—"]));
      });
      if(pegEntries.length){
        const ws=XLSX.utils.aoa_to_sheet([["Nama","NIK","Bulan","Tgl Awal","Tgl Akhir","Jam Masuk","Jam Keluar","Keperluan","Keterangan","TTD","Lokasi Keluar (GPS)","Tautan Foto Selfie"],...pegEntries]);
        ws["!cols"]=[{wch:22},{wch:14},{wch:10},{wch:14},{wch:14},{wch:10},{wch:10},{wch:28},{wch:20},{wch:14},{wch:24},{wch:32}];
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
  /* ── AUTO-SAVE MULTI-PASIEN DARI OCR ──────────────────────────────────
     Dipanggil saat 1 scan formulir menghasilkan ≥2 kandidat pasien.
     BERBEDA dari saveOpFn: TIDAK memvalidasi field WAJIB (nama/tanggal/dsb
     boleh kosong) karena tujuannya supaya perawat tidak perlu isi form
     satu-satu untuk tiap pasien — cukup scan, semua langsung masuk daftar,
     lalu perawat TINGGAL EDIT baris yang kurang lengkap. Setiap baris
     ditandai needsReview:true supaya tampil dengan border biru di daftar.

     FIX AUDIT #30 (CRITICAL — sebelumnya field semi-terstruktur seperti
     date/time/room/bloodType/age ditulis MENTAH tanpa validasi apa pun di
     jalur ini, berbeda dengan applyOcrResult yang sudah divalidasi. Kalau
     model salah baca opType jadi "cyto", banner darurat "EMERGENCY — SEGERA
     DITANGANI" langsung tampil ke semua staf TANPA verifikasi manusia.
     Sekarang: field semi-terstruktur divalidasi via validateOcrScalarFields
     (sama seperti applyOcrResult), teks bebas dipotong via capOcrText, DAN
     — kalau opType hasil scan adalah "cyto" — status TIDAK langsung dipasang
     "cyto" ke database; disimpan dulu sebagai "elektif"+needsReview dengan
     catatan dugaan CYTO di `keterangan`, supaya banner darurat/urutan-atas
     hanya muncul setelah perawat memverifikasi & mengubahnya secara sadar
     lewat form edit biasa. */
  /* FIX AUDIT #40 (HIGH — Tidak Ada Pengecekan Duplikat pada Auto-Save Batch):
     saveOpFn (jalur manual) sudah mengecek duplikat pasien+tanggal+jam
     sebelum menyimpan (lihat `dup`/`dupSameDay` di bawah), tapi
     autoSaveOcrBatch (jalur OCR ≥2 kandidat, tanpa review manusia)
     SEBELUMNYA tidak mengecek sama sekali. Kalau formulir yang sama
     ter-scan 2x (kamera lag, atau perawat mengira gagal lalu upload
     ulang), data pasien bisa masuk GANDA ke jadwal produksi tanpa
     peringatan apa pun. Sekarang dicek terhadap (a) jadwal yang sudah
     ada di `ops`, dan (b) duplikat DI DALAM hasil scan itu sendiri
     (kalau model vision membaca baris yang sama dua kali). Kandidat
     duplikat DILEWATI (tidak disimpan), dan perawat diberi tahu lewat
     toast supaya sadar ada entri yang sengaja tidak diproses. */
  const autoSaveOcrBatch = useCallback(async (results: OcrExtractResult[]) => {
    /* Guard tambahan (belt-and-suspenders): batasi jumlah entri yang diproses
       meskipun ScanFormulir sudah membatasi di sisi pemanggil. */
    const capped = results.length > MAX_OCR_BATCH ? results.slice(0, MAX_OCR_BATCH) : results;

    const dupKey = (patient: string, date: string, time: string) =>
      `${patient.toLowerCase().trim()}|${date}|${time}`;
    const existingKeys = new Set(
      ops.map((o: any) => dupKey(o.patient || "", o.date || "", o.time || ""))
    );
    const seenInBatch = new Set<string>();
    let skippedDup = 0;

    const newOps: Operation[] = [];
    /* SARAN PENINGKATAN #3: kumpulkan ringkasan field yang berasal dari OCR
       untuk tiap entri batch, dipakai untuk log audit trail per-entri di
       bawah — sama semangatnya dengan audit trail di jalur 1-kandidat
       (saveOpFn), meski di sini tidak ada tahap "edit sebelum simpan"
       karena batch memang didesain langsung tersimpan (dengan needsReview). */
    const ocrFieldSummary = new Map<string, {lowConfidence:boolean; lowConfidenceFields:string[]; hadCyto:boolean; hadBloodTypeOrAllergy:boolean}>();
    for (const rRaw of capped) {
      const r = rRaw;
      /* SARAN PENINGKATAN #10: pakai buildOcrSafeApply() yang SAMA dengan
         applyOcrResult (jalur 1-kandidat) — satu sumber kebenaran untuk
         validasi & aturan kehati-hatian, lihat komentar di definisinya. */
      const safe = buildOcrSafeApply(r);
      const v = safe.scalars;
      const suspectedCyto = safe.suspectedCyto;
      const patientClean = sanitize(capOcrText(r.patient || ""));
      /* Hanya dianggap "cukup sinyal untuk dedup" kalau ketiga kunci
         (nama, tanggal, jam) terisi — kalau salah satu kosong, terlalu
         berisiko false-positive (mis. dua pasien beda yang sama-sama
         belum ada jamnya kebaca "0000"), jadi tetap disimpan sebagai
         entri baru dan biarkan perawat yang memutuskan saat review. */
      const hasFullKey = !!(patientClean && v.date && v.time);
      const key = dupKey(patientClean, v.date, v.time);
      if (hasFullKey && (existingKeys.has(key) || seenInBatch.has(key))) {
        skippedDup++;
        continue;
      }
      if (hasFullKey) seenInBatch.add(key);
      /* SARAN PENINGKATAN #1: tandai butuh perhatian ekstra kalau model
         sendiri melaporkan confidence rendah untuk entri ini (mis. tulisan
         tangan kurang jelas) — sebelumnya sinyal ini dibuang begitu saja. */
      const lowConfNote = r.confidence==="low" ? "⚠ Keyakinan AI rendah untuk entri ini (kemungkinan tulisan tangan kurang jelas) — periksa lebih teliti dari entri lain." : "";
      const notesCombined = [safe.cautionNote, lowConfNote].filter(Boolean).join("\n") || undefined;
      const newId = gId();
      newOps.push({
        id: newId,
        patient: patientClean,
        age: v.age,
        ageMonths: v.ageMonths,
        rm: sanitize(capOcrText(r.rm || "", 50)),
        /* FIX AUDIT #30 (via buildOcrSafeApply): jangan pasang "cyto"
           otomatis tanpa verifikasi — default ke "elektif" dan catat
           dugaan di specialNeeds/needsReview. */
        opType: (suspectedCyto ? "elektif" : (r.opType || "elektif")) as "elektif" | "semi" | "cyto",
        status: "scheduled",
        date: v.date,
        time: v.time,
        room: v.room,
        surgeon: (() => {
          const rawName = sanitize(capOcrText(r.surgeon || "", 80));
          if (!rawName) return rawName;
          const candidates = staff.filter((s:any)=>s.type==="surgeon").map((s:any)=>s.name);
          return findBestStaffMatch(rawName, candidates) || rawName;
        })(),
        anesthesiologist: (() => {
          const rawName = sanitize(capOcrText(r.anesthesiologist || "", 80));
          if (!rawName || looksLikeNonDoctorAnesthesiaValue(rawName)) return rawName;
          const candidates = staff.filter((s:any)=>s.type==="anesthesiologist").map((s:any)=>s.name);
          return findBestStaffMatch(rawName, candidates) || rawName;
        })(),
        procedure: sanitize(capOcrText(r.procedure || "")),
        diagnosis: sanitize(capOcrText(r.diagnosis || "")),
        allergy: sanitize(capOcrText(r.allergy || "Tidak Ada", 120)),
        bloodType: v.bloodType,
        specialNeeds: notesCombined,
        reminders: [],
        requests: [],
        createdAt: fNow(),
        updated_at: new Date().toISOString(),
        needsReview: true,
      } as Operation);
      ocrFieldSummary.set(newId, {
        lowConfidence: r.confidence==="low",
        lowConfidenceFields: r.lowConfidenceFields||[],
        hadCyto: suspectedCyto,
        hadBloodTypeOrAllergy: !!(r.bloodType || r.allergy),
      });
    }

    /* FIX AUDIT #51 (CRITICAL — Data "hilang" dari Tab Jadwal setelah OCR
       batch): SEBELUMNYA upsertOneToSupa dipanggil fire-and-forget (tanpa
       await) di dalam forEach, lalu KODE DI BAWAHNYA (toast sukses +
       setTab("jadwal")) langsung jalan TANPA menunggu upsert selesai.
       Akibatnya: (a) berpindah ke Tab Jadwal SEBELUM setOps sempat
       terpanggil dari .then(), sehingga pasien baru sempat tidak terlihat
       sesaat — terlihat seperti "belum masuk"; (b) kalau salah satu upsert
       GAGAL (network drop dsb), toast UTAMA tetap bilang "✓ semua
       tersimpan" karena tidak pernah menunggu hasil sebenarnya — pasien itu
       tidak akan pernah muncul di Jadwal meski perawat sudah lihat toast
       sukses. Sekarang: SEMUA upsert ditunggu (Promise.allSettled) dulu,
       baru setOps/toast/pindah tab dijalankan — dan toast secara eksplisit
       membedakan berapa yang benar-benar tersimpan vs gagal. */
    /* PENAMBAHAN #8 (offline queue): cek status koneksi SEKALI di awal
       batch — kalau offline, tidak perlu buang waktu mencoba network yang
       sudah pasti gagal untuk tiap item; langsung masuk antrean lokal. */
    const wasOffline = !navigator.onLine;

    const upsertResults = await Promise.allSettled(
      newOps.map(async (newOp) => {
        logAudit("Tambah (OCR batch)", newOp);
        const summary = ocrFieldSummary.get(newOp.id);
        AmbilLogMedis(
          "DAFTAR_PASIEN",
          role === "admin" ? "Admin" : "Perawat",
          "TAMBAH_PASIEN",
          `Tambah pasien dari scan OCR (auto-batch): ${newOp.patient || "(nama belum terisi)"} (RM: ${newOp.rm || "-"})`,
          { ...newOp, ocrAuditTrail: { sumberBatch: true, confidenceRendah: !!summary?.lowConfidence, fieldRaguRagu: summary?.lowConfidenceFields||[], dugaanCytoTerdeteksi: !!summary?.hadCyto, adaGolonganDarahAtauAlergi: !!summary?.hadBloodTypeOrAllergy } }
        );
        if (newOp.specialNeeds?.includes("CYTO")) {
          AmbilLogMedis(
            "DAFTAR_PASIEN",
            role === "admin" ? "Admin" : "Perawat",
            "DUGAAN_CYTO_OCR",
            `Dugaan CYTO dari scan OCR (BELUM diverifikasi): NRM ${newOp.rm || "-"} — Pasien: ${newOp.patient || "-"}`,
            { ...newOp }
          );
        }
        if (wasOffline) {
          enqueueOfflineOp(newOp.id, newOp);
          return { newOp, res: { ok: false }, offlineQueued: true };
        }
        const res = await upsertOneToSupa("kb_operasi", newOp);
        if (!res?.ok && !navigator.onLine) {
          // Koneksi putus DI TENGAH request (bukan sejak awal) — tetap
          // masuk antrean supaya tidak hilang begitu saja.
          enqueueOfflineOp(newOp.id, newOp);
          return { newOp, res, offlineQueued: true };
        }
        return { newOp, res, offlineQueued: false };
      })
    );

    const savedOps: Operation[] = [];
    const failedOps: { newOp: Operation; error: string }[] = [];
    let queuedCount = 0;
    for (const r of upsertResults) {
      if (r.status === "fulfilled") {
        if (r.value.res?.ok) savedOps.push(r.value.newOp);
        else if (r.value.offlineQueued) { savedOps.push(r.value.newOp); queuedCount++; }
        else failedOps.push({ newOp: r.value.newOp, error: r.value.res?.error || "kesalahan tidak diketahui" });
      } else {
        // newOp tidak diketahui di sini karena reject terjadi sebelum return —
        // seharusnya sangat jarang karena upsertOneToSupa sendiri tidak pernah throw.
        failedOps.push({ newOp: null as any, error: String(r.reason) });
      }
    }

    if (savedOps.length > 0) {
      setOps((p) => {
        const existingIds = new Set(p.map((o: any) => o.id));
        const toAdd = savedOps.filter(o => !existingIds.has(o.id));
        return toAdd.length ? [...p, ...toAdd] : p;
      });
    }
    failedOps.forEach(({ newOp, error }) => {
      showToast(`⚠ Gagal menyimpan "${newOp?.patient || "(tanpa nama)"}" ke Supabase: ${error}`, C.d);
    });

    if (skippedDup > 0) {
      AmbilLogMedis(
        "DAFTAR_PASIEN",
        role === "admin" ? "Admin" : "Perawat",
        "OCR_DUPLIKAT_DILEWATI",
        `${skippedDup} entri hasil scan OCR dilewati karena sudah ada jadwal dengan pasien+tanggal+jam yang sama`,
        { skippedDup, totalDiproses: capped.length }
      );
    }

    const jumlahDugaanCyto = savedOps.filter(o=>o.specialNeeds?.includes("CYTO")).length;
    const bagianDup = skippedDup>0 ? ` (${skippedDup} entri dilewati karena sudah terdaftar sebelumnya)` : "";
    const bagianGagal = failedOps.length>0 ? ` — ⚠ ${failedOps.length} entri GAGAL tersimpan (lihat notifikasi merah), coba scan ulang untuk entri tsb.` : "";
    const bagianQueued = queuedCount>0 ? ` — 📴 ${queuedCount} entri disimpan LOKAL (offline), akan otomatis sinkron ke database saat koneksi kembali.` : "";
    if (newOps.length === 0 && skippedDup > 0) {
      showToast(`⚠ Semua ${skippedDup} entri hasil scan sudah terdaftar sebelumnya — tidak ada yang disimpan (cegah duplikat).`, C.w);
    } else if (savedOps.length === 0) {
      showToast(`⚠ Tidak ada pasien yang berhasil tersimpan.${bagianGagal || " Periksa koneksi lalu coba lagi."}`, C.d);
    } else {
      showToast(
        jumlahDugaanCyto>0
          ? `✓ ${savedOps.length} pasien otomatis tersimpan${bagianDup}${bagianGagal}${bagianQueued} — ⚠ ${jumlahDugaanCyto} di antaranya DIDUGA CYTO, segera verifikasi manual!`
          : `✓ ${savedOps.length} pasien otomatis tersimpan${bagianDup}${bagianGagal}${bagianQueued} — mohon lengkapi & periksa data yang belum lengkap`,
        (jumlahDugaanCyto>0 || failedOps.length>0) ? C.d : C.s
      );
    }
    /* Pindah ke Tab Jadwal HANYA setelah setOps di atas benar-benar
       dipanggil (kita sudah di titik ini setelah await Promise.allSettled),
       jadi data yang tersimpan dijamin sudah ada di state saat perawat
       melihat Tab Jadwal — tidak ada lagi jeda "belum muncul". */
    if(savedOps.length>0) setTab("jadwal");
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[showToast, upsertOneToSupa, logAudit, role, ops]);


  const saveOpFn = useCallback(({opForm, editingOp, setOpErrors, setDupWarn, resetOp, ocrOriginalValues}: SaveOpFnArgs) => {
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
    /* SARAN PENINGKATAN #3 (audit trail OCR): kalau entri ini berasal dari
       scan OCR (ocrOriginalValues terisi — lihat applyOcrResult di
       ViewDaftar), bandingkan nilai AKHIR yang benar-benar disimpan vs
       nilai ASLI yang dibaca AI. Dicatat lewat AmbilLogMedis (infrastruktur
       audit yang sudah ada, dipakai di 30+ titik lain di app ini) supaya
       kalau suatu saat ada laporan data pasien tidak akurat, bisa
       ditelusuri: apakah itu salah baca OCR yang luput dikoreksi perawat
       (field masuk daftar "dipakai apa adanya"), atau justru sudah
       dikoreksi tapi hasil koreksinya yang salah (field masuk daftar
       "dikoreksi user") — dua akar masalah yang butuh tindak lanjut beda. */
    if(ocrOriginalValues && Object.keys(ocrOriginalValues).length>0){
      const dipakaiApaAdanya: string[] = [];
      const dikoreksiUser: string[] = [];
      Object.entries(ocrOriginalValues).forEach(([k, origVal])=>{
        const finalVal = (clean as any)[k];
        if(String(finalVal ?? "").trim() === String(origVal ?? "").trim()) dipakaiApaAdanya.push(k);
        else dikoreksiUser.push(k);
      });
      AmbilLogMedis("DAFTAR_PASIEN", role==="admin"?"Admin":"Perawat", "OCR_AUDIT_TRAIL",
        `Simpan entri dari hasil scan OCR — ${dipakaiApaAdanya.length} field dipakai apa adanya, ${dikoreksiUser.length} field dikoreksi user sebelum simpan.`,
        { patient: clean.patient, rm: clean.rm, fieldDipakaiApaAdanya: dipakaiApaAdanya, fieldDikoreksiUser: dikoreksiUser, nilaiAsliHasilOcr: ocrOriginalValues }
      );
    }
    if(editingOp){
      /* FIX AUDIT #45 (MEDIUM/HIGH — Flag needsReview Tidak Pernah Hilang):
         Baris kb_operasi hasil auto-batch OCR (lihat autoSaveOcrBatch)
         ditandai needsReview:true supaya tampil dengan border+badge biru
         "🔍 Perlu Review (dari OCR)" di tab Jadwal — mengingatkan staf
         bahwa data itu belum diverifikasi manusia. SEBELUMNYA, saat
         perawat membuka baris tsb lewat Edit dan menekan Simpan (yaitu
         TINDAKAN VERIFIKASI itu sendiri), field needsReview TIDAK PERNAH
         disertakan di `clean` (hasil spread dari opForm, yang tidak
         punya field ini), sehingga `{...editingOp,...clean}` mewariskan
         needsReview:true dari editingOp APA ADANYA — flag itu menempel
         SELAMANYA walau datanya sudah diperiksa & dikonfirmasi berkali-
         kali. Lama-lama tab Jadwal penuh badge biru yang tidak berarti
         apa-apa lagi (sudah direview tapi masih ditandai "perlu review"),
         sehingga staf tidak bisa lagi memercayai badge itu untuk melihat
         entri MANA yang benar-benar belum diperiksa. Sekarang: proses
         Edit+Simpan dianggap sebagai tindakan verifikasi eksplisit —
         needsReview otomatis diset false begitu baris disimpan ulang. */
      const updatedOp = {...editingOp,...clean, needsReview:false, updated_at: new Date().toISOString()};
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
      if (!navigator.onLine) {
        /* PENAMBAHAN #8 (offline queue): jangan buang waktu coba network
           yang sudah pasti gagal — langsung antrekan, tetap tampilkan di
           UI secara optimistic (sudah di-setOps di atas), dengan pesan
           jujur bahwa ini baru tersimpan lokal. */
        enqueueOfflineOp(updatedOp.id, updatedOp);
        showToast("📴 Offline — perubahan tersimpan LOKAL, akan otomatis sinkron saat koneksi kembali.", "#E65100");
      } else {
        upsertOneToSupa("kb_operasi", updatedOp).then((res:any)=>{
          if(!res?.ok) {
            if (!navigator.onLine) {
              enqueueOfflineOp(updatedOp.id, updatedOp);
              showToast("📴 Koneksi terputus — perubahan tersimpan LOKAL, akan otomatis sinkron saat online.", "#E65100");
            } else {
              setOps(p=>p.map(o=>o.id===editingOp.id ? editingOp : o));
              showToast(`⚠ Gagal menyimpan ke Supabase: ${res?.error||"kesalahan tidak diketahui"}`,C.d);
            }
          } else {
            showToast("✓ Jadwal diperbarui & tersimpan ke Supabase",C.s);
          }
        });
      }
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
      if (!navigator.onLine) {
        enqueueOfflineOp(newOp.id, newOp);
        setOps(p=>{ if(p.some((o:any)=>o.id===newOp.id)) return p; return [...p, newOp]; });
        showToast("📴 Offline — pasien tersimpan LOKAL, akan otomatis sinkron saat koneksi kembali.", "#E65100");
      } else {
        upsertOneToSupa("kb_operasi", newOp).then((res:any)=>{
          if(!res?.ok) {
            if (!navigator.onLine) {
              enqueueOfflineOp(newOp.id, newOp);
              setOps(p=>{ if(p.some((o:any)=>o.id===newOp.id)) return p; return [...p, newOp]; });
              showToast("📴 Koneksi terputus — pasien tersimpan LOKAL, akan otomatis sinkron saat online.", "#E65100");
            } else {
              showToast(`⚠ Gagal menyimpan ke Supabase: ${res?.error||"kesalahan tidak diketahui"}`,C.d);
            }
          } else {
            setOps(p=>{ if(p.some((o:any)=>o.id===newOp.id)) return p; return [...p, newOp]; });
            showToast("✓ Jadwal tersimpan ke Supabase",C.s);
            fireConfettiBurst();
          }
        });
      }
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
    const notifId = gId();
    /* FIX AUDIT #24: cap 200 entri terbaru untuk endurance 24/7. */
    setNotifs(p=>[{id:notifId,type,label:(type==="H-1"?"H-1":"1 Jam")+" → "+op.surgeon,patient:op.patient,procedure:op.procedure,message:"",sentAt:fNow()},...p].slice(0,200));
    const updatedOp = {...op, reminders:[...(op.reminders||[]),type]};
    /* Optimistic update */
    setOps(p=>p.map((o:any)=>o.id===op.id ? updatedOp : o));
    upsertOneToSupa("kb_operasi", updatedOp).then(res=>{
      if(!res?.ok){
        setOps(p=>p.map((o:any)=>o.id===op.id ? op : o));
        /* FIX AUDIT #26 (LOW — Rollback Tidak Lengkap): sebelumnya kalau
           upsertOneToSupa gagal, ops di-rollback tapi entri notifs yang
           sudah ditambahkan optimistically TIDAK ikut dihapus — log UI
           tetap menampilkan "✓ berhasil" walau sebenarnya gagal tersimpan.
           Sekarang notifId dilacak agar entri log ikut dihapus saat rollback. */
        setNotifs(p=>p.filter((n:any)=>n.id!==notifId));
        showToast(`⚠ Gagal mencatat pengingat: ${res?.error||"kesalahan tidak diketahui"}`,C.d);
      }
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
      AmbilLogMedis("KEAMANAN", "Perawat", "AKSES_DITOLAK",
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
      AmbilLogMedis("KEAMANAN", "Perawat", "AKSES_DITOLAK",
        `Akses ditolak ke tab roster_gen via bottom nav (role: ${role})`, { tab: k, role });
      setTab("home");
      return;
    }
    AmbilLogMedis("NAVIGASI", role==="admin"?"Admin":"Perawat", "PINDAH_TAB",
      `Berpindah ke tab: ${k} (via bottom nav)`, { tab: k, role, source: "bottom_nav" });
    setTab(k);
  }, [role, showToast]);
  const TABS = useMemo(
    () => lemburOnlyMode
      ? ALL_TABS.filter(t => t.k === "lembur")
      : ALL_TABS.filter(t => t.roles.includes(role)),
    [role, lemburOnlyMode]
  );

  if(!pinLoaded) return (
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"linear-gradient(135deg,#0B1220,#151F35 55%,#0B1220)"}}>
      <div style={{textAlign:"center",color:"#fff"}}>
        <img src="/logo.jpeg" style={{width:80,height:80,borderRadius:"50%",objectFit:"cover",marginBottom:16,display:"block",margin:"0 auto 16px"}}/>
        <div style={{fontSize:16,fontWeight:700,marginBottom:8}}>SISTEM KOORDINASI KAMAR BEDAH</div>
        <div style={{fontSize:13,opacity:.8,marginBottom:20}}>Rumah Sakit Panti Rini</div>
        <div style={{fontSize:13,opacity:.7}}>⏳ Memuat konfigurasi...</div>
      </div>
    </div>
  );
  if(!pinOK) return <PinScreen onVerify={handlePinVerify} pinAdmin={pinAdmin} pinPerawat={pinPerawat} isFirstTime={isFirstTime} lemburPegawai={lemburPegawai}/>;

  const todayOpsCount = ops.filter((o:any)=>o.date===todayDate()).length;

  const content = (
    <ErrorBoundary>
      {tab==="home" && (
        <div style={{paddingBottom:32,position:"relative"}}>
          {/* #1 Aurora Background versi SUBTLE (opacity rendah) — benang
              merah visual dengan layar login, tanpa ganggu keterbacaan
              data dashboard. */}
          <AuroraBackground opacity={0.12}/>
          <div style={{position:"relative",zIndex:1}}>
          {/* Welcome Banner */}
          <ScrollReveal>
          <HomeParallaxBanner/>
          </ScrollReveal>

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
                  /* Pengaman: satu kamar tidak bisa menjalankan dua operasi sekaligus.
                     Kalau masih ada operasi lain BERLANGSUNG di kamar yang sama (biasanya
                     karena lupa menekan Selesai), tawarkan untuk menutupnya dulu. */
                  if(newStatus==="ongoing" && op.room){
                    const prev=ops.filter((o:any)=>o.id!==op.id && o.status==="ongoing" && o.room===op.room);
                    if(prev.length>0){
                      const nama=prev.map((o:any)=>maskName(o.patient,privacyMode)).join(", ");
                      if(!window.confirm(`Di ${op.room} masih tercatat BERLANGSUNG: ${nama}. Tandai operasi itu SELESAI lalu mulai operasi ini? (OK = ya, Batal = tidak ada perubahan)`)) return;
                      const nowIso=new Date().toISOString();
                      prev.forEach((o:any)=>{
                        const closed={...o,status:"done",updated_at:nowIso};
                        setOps((p:any[])=>p.map((x:any)=>x.id===o.id?closed:x));
                        AmbilLogMedis("JADWAL_OPERASI","Sistem","TUTUP_OTOMATIS_SEBELUM_MULAI","Operasi ditandai selesai karena operasi berikutnya dimulai di kamar yang sama",{id:o.id,room:o.room,next:op.id});
                        upsertOneToSupa("kb_operasi",closed).then((res:any)=>{ if(!res?.ok){ setOps((p:any[])=>p.map((x:any)=>x.id===o.id?o:x)); showToast(`⚠ Gagal menutup operasi sebelumnya: ${res?.error||"error"}`,C.d); } });
                      });
                    }
                  }
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

          {/* Quick Stats — Bento grid: 1 sel besar (Operasi Hari Ini) + 2 sel kecil bertumpuk */}
          <ScrollReveal delay={80}>
          <div style={{display:"grid",gridTemplateColumns:"1.3fr 1fr",gap:12,marginBottom:20}}>
            {/* Sel besar: Operasi Hari Ini — #2 SpotlightCard (cahaya
                mengikuti kursor saat hover) */}
            <GlareHover>
            <SpotlightCard style={{background:"linear-gradient(135deg,#EFF6FF,#DBEAFE)",borderRadius:18,padding:"20px 18px",boxShadow:"6px 6px 16px rgba(148,163,184,.16), -4px -4px 12px rgba(255,255,255,.7)",border:"1px solid #BFDBFE",display:"flex",flexDirection:"column",justifyContent:"center"}}>
              <div style={{fontSize:11,fontWeight:700,color:"#1565C0",letterSpacing:.5,textTransform:"uppercase",opacity:.8,marginBottom:6}}>
                {/* #8 DecryptedText: label "teracak lalu terbentuk" saat
                    pertama muncul — #7 ShinyText: highlight berkilau */}
                <ShinyText style={{color:"#1565C0"}}><DecryptedText text="Operasi"/></ShinyText>
              </div>
              <div style={{fontSize:44,fontWeight:900,color:"#1565C0",lineHeight:1}}><AnimatedNumber value={todayOpsCount}/></div>
              <div style={{fontSize:12,color:"#3B5A80",marginTop:6,fontWeight:600}}>Terjadwal Hari Ini</div>
            </SpotlightCard>
            </GlareHover>
            {/* 2 sel kecil bertumpuk — juga SpotlightCard */}
            <div style={{display:"flex",flexDirection:"column",gap:12}}>
              <SpotlightCard spotlightColor={rtStatus==="online"?"rgba(34,197,94,.25)":"rgba(250,204,21,.25)"} style={{background:rtStatus==="online"?"linear-gradient(135deg,#F1F8F4,#E8F5E9)":"linear-gradient(135deg,#FFF8F0,#FFF3E0)",borderRadius:14,padding:"12px 14px",boxShadow:"4px 4px 12px rgba(148,163,184,.14), -3px -3px 8px rgba(255,255,255,.7)",border:`1px solid ${rtStatus==="online"?"#C8E6C9":"#FFE0B2"}`,flex:1,display:"flex",alignItems:"center",gap:10}}>
                <div style={{fontSize:20}}>{rtStatus==="online"?"🟢":"🟡"}</div>
                <div>
                  <div style={{fontSize:11,fontWeight:700,color:rtStatus==="online"?C.s:C.w}}>{rtStatus==="online"?"Online":"Offline"}</div>
                  <div style={{fontSize:9,color:"#5C677D"}}>Cloud Sync</div>
                </div>
              </SpotlightCard>
              <SpotlightCard spotlightColor="rgba(168,85,247,.25)" style={{background:"linear-gradient(135deg,#F5F3FF,#EDE7F6)",borderRadius:14,padding:"12px 14px",boxShadow:"4px 4px 12px rgba(148,163,184,.14), -3px -3px 8px rgba(255,255,255,.7)",border:"1px solid #D1C4E9",flex:1,display:"flex",alignItems:"center",gap:10}}>
                <div style={{fontSize:20,fontWeight:900,color:"#512DA8"}}><AnimatedNumber value={staff.length}/></div>
                <div>
                  <div style={{fontSize:11,fontWeight:700,color:"#512DA8"}}>Staf</div>
                  <div style={{fontSize:9,color:"#5C677D"}}>Terdaftar</div>
                </div>
              </SpotlightCard>
            </div>
          </div>
          </ScrollReveal>


          {/* Statistik bulanan (admin + perawat) */}
          <ScrollReveal delay={120}>
            <StatistikBulanan ops={ops} archive={archive as any[]}/>
          </ScrollReveal>

          {/* Quick Links — #4 MagneticButton (tertarik halus ke kursor) +
              #9 TiltCard (kemiringan 3D halus) — SENGAJA hanya di sini
              (tombol navigasi), BUKAN di card data pasien, supaya kesan
              tetap serius untuk konteks kerja klinis. */}
          <ScrollReveal delay={140}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            <TiltCard maxTilt={4}>
              <MagneticButton onClick={()=>setTab("daftar")} strength={8} style={{width:"100%",background:"linear-gradient(135deg,#0284c7,#38bdf8)",color:"#fff",border:"none",borderRadius:14,padding:"18px 14px",fontSize:15,fontWeight:800,cursor:"pointer",boxShadow:"0 4px 16px rgba(2,132,199,.35)",textAlign:"center",fontFamily:"inherit"}}>
                📝 Daftar Operasi
              </MagneticButton>
            </TiltCard>
            <TiltCard maxTilt={4}>
              <MagneticButton onClick={()=>setTab("laporan")} strength={8} style={{width:"100%",background:"linear-gradient(135deg,#1d4ed8,#6366f1)",color:"#fff",border:"none",borderRadius:14,padding:"18px 14px",fontSize:15,fontWeight:800,cursor:"pointer",boxShadow:"0 4px 16px rgba(29,78,216,.30)",textAlign:"center",fontFamily:"inherit"}}>
                📄 Laporan Kepala Jaga
              </MagneticButton>
            </TiltCard>
            <TiltCard maxTilt={4}>
              <MagneticButton onClick={()=>setTab("jadwal")} strength={8} style={{width:"100%",background:"linear-gradient(135deg,#1e3a8a,#1d4ed8)",color:"#fff",border:"none",borderRadius:14,padding:"18px 14px",fontSize:15,fontWeight:800,cursor:"pointer",boxShadow:"0 4px 16px rgba(30,58,138,.30)",textAlign:"center",fontFamily:"inherit"}}>
                📋 Lihat Jadwal
              </MagneticButton>
            </TiltCard>
            {role==="admin" && (
              <TiltCard maxTilt={4}>
                <MagneticButton onClick={()=>setTab("roster_gen")} strength={8} style={{width:"100%",background:"linear-gradient(135deg,#0369a1,#0284c7)",color:"#fff",border:"none",borderRadius:14,padding:"18px 14px",fontSize:15,fontWeight:800,cursor:"pointer",boxShadow:"0 4px 16px rgba(3,105,161,.30)",textAlign:"center",fontFamily:"inherit"}}>
                  🗓 Generator Roster
                </MagneticButton>
              </TiltCard>
            )}
          </div>
          </ScrollReveal>
          </div>
        </div>
      )}
      {tab==="jadwal"     && <ViewJadwal ops={ops} setOps={setOps} startEditOp={startEditOp} deleteOp={deleteOp} sendReminder={sendReminder} reqOpId={reqOpId} setReqOpId={setReqOpId} reqText={reqText} setReqText={setReqText} addReq={addReq} delReq={delReq} getPhone={getPhone} setNotifs={setNotifs} showToast={showToast} privacyMode={privacyMode} role={role} upsertOneToSupa={upsertOneToSupa} initialLoading={initialLoading}/>}
      {tab==="daftar"     && <ViewDaftar pendingEditOp={pendingEditOp} clearPendingEditOp={()=>setPendingEditOp(null)} saveOpFn={saveOpFn} autoSaveOcrBatch={autoSaveOcrBatch} staff={staff} setTab={setTab} templates={templates} setTemplates={setTemplates} showToast={showToast}/>}
      {tab==="laporan"    && <ViewLaporan ops={ops} staff={staff} roster={roster} showToast={showToast} role={role} privacyMode={privacyMode}/>}
      {tab==="laporan_harian" && <ViewLaporanHarian ops={ops} staff={staff} showToast={showToast} role={role} privacyMode={privacyMode} upsertOneToSupa={upsertOneToSupa} deleteFromSupa={deleteFromSupa} dbxCfg={dbxCfg}/>}
      {tab==="jadwal_ok" && <ViewJadwalOK ops={ops} showToast={showToast} role={role} privacyMode={privacyMode} upsertOneToSupa={upsertOneToSupa} deleteFromSupa={deleteFromSupa} dbxCfg={dbxCfg}/>}
      {tab==="wa"         && <ViewKirimWA ops={ops} staff={staff} setNotifs={setNotifs} showToast={showToast}/>}
      {tab==="statistik"  && <ViewStatistik ops={ops} archive={archive}/>}
      {tab==="staf"       && <ViewStaf staff={staff} setStaff={setStaff} roster={roster} setRoster={setRoster} showToast={showToast} upsertOneToSupa={upsertOneToSupa} deleteFromSupa={deleteFromSupa} upsertBulkToSupa={upsertBulkToSupa}/>}
      {tab==="lembur"     && <ViewLembur lemburPegawai={lemburPegawai} setLemburPegawai={setLemburPegawai} lemburData={lemburData} setLemburData={setLemburData} showToast={showToast} supaCfg={supaCfg} dbxCfg={dbxCfg} role={role} upsertOneToSupa={upsertOneToSupa} deleteFromSupa={deleteFromSupa} autoSelPeg={autoSelPeg} onAutoSelPegUsed={()=>setAutoSelPeg("")}/>}
      {tab==="monitoring" && <ViewMonitoring monitoringEntries={monitoringEntries} setMonitoringEntries={setMonitoringEntries} monitoringCfg={monitoringCfg} setMonitoringCfg={setMonitoringCfg} showToast={showToast} supaCfg={supaCfg} dbxCfg={dbxCfg} role={role} upsertOneToSupa={upsertOneToSupa} deleteFromSupa={deleteFromSupa}/>}
      {tab==="roster_gen" && role==="admin" && <ViewRosterGenerator showToast={showToast} upsertOneToSupa={upsertOneToSupa} deleteFromSupa={deleteFromSupa} dbxCfg={dbxCfg}/>}
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
      <div className="kb-noise-overlay"/>
      <InstallPrompt />
      {/* PENAMBAHAN #9 (loading bar profesional): garis progress tipis di
          paling atas layar selama data awal dimuat dari Supabase — pola
          umum di web app modern (GitHub/YouTube), tidak mengganggu layout
          karena position:fixed & tinggi cuma 3px. */}
      {initialLoading && (
        <div className="kb-topbar"><span/></div>
      )}

      {/* ── Sticky header (mobile + tablet: full header; desktop: top bar only) ── */}
      <div className="kb-header" style={{background:"linear-gradient(135deg,rgba(11,18,32,.92),rgba(21,31,53,.90) 55%,rgba(11,18,32,.92))",backdropFilter:"blur(14px) saturate(160%)",WebkitBackdropFilter:"blur(14px) saturate(160%)",position:"sticky",top:0,zIndex:100,boxShadow:"0 4px 24px rgba(11,18,32,.35)",borderBottom:"1px solid rgba(232,206,143,.18)"}}>
        <div className="kb-header-inner" style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
          <div style={{display:"flex",alignItems:"center",gap:12}}>
            <img src="/logo.jpeg" alt="Logo KB" style={{width:50,height:50,borderRadius:"50%",objectFit:"cover",border:"2px solid rgba(255,255,255,.5)",animation:"lgFloat 3s ease-in-out infinite",flexShrink:0}}/>
            <div>
              <div style={{fontSize:9,fontWeight:700,color:"rgba(255,255,255,.6)",letterSpacing:2,textTransform:"uppercase",marginBottom:1}}>Rumah Sakit Panti Rini</div>
              <div className="kb-title kb-font-display" style={{fontWeight:700,color:"#fff",letterSpacing:.3,lineHeight:1.15}}>SISTEM KOORDINASI<br/>KAMAR BEDAH</div>
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
            {ops.filter((o:any)=>o.status==="ongoing"&&o.date===toLocalISODate(new Date())).length>0 && (
              <span style={{background:"#F44336",color:"#fff",fontSize:11,fontWeight:700,padding:"3px 9px",borderRadius:20}}>
                ● {ops.filter((o:any)=>o.status==="ongoing"&&o.date===toLocalISODate(new Date())).length} berlangsung
              </span>
            )}
            {role==="admin" && <button onClick={()=>setShowPinMgmt(true)} style={{background:"rgba(255,255,255,.15)",border:"none",borderRadius:8,padding:"6px 10px",color:"#fff",fontSize:11,cursor:"pointer",fontFamily:"inherit",fontWeight:700}}>⚙ PIN</button>}
            <span style={{background:role==="admin"?"rgba(255,255,255,.25)":"rgba(255,255,255,.12)",border:"1px solid rgba(255,255,255,.3)",borderRadius:8,padding:"4px 9px",color:"#fff",fontSize:10,fontWeight:700}}>{role==="admin"?"🔐 Admin":"👤 Perawat"}</span>
            <button onClick={()=>{setPinOK(false);setRole("perawat");setLemburOnlyMode(false);}} style={{background:"rgba(255,255,255,.15)",border:"none",borderRadius:8,padding:"6px 12px",color:"#fff",fontSize:11,cursor:"pointer",fontFamily:"inherit",fontWeight:700}}>🔒 Kunci</button>
          </div>
        </div>
        {/* Mobile/tablet tabs — hidden on desktop via CSS */}
        <div className="kb-tabs">
          {TABS.map(t=>(
            <button key={t.k} data-tab={t.k} onClick={handleTabClickEvt} className="kb-tab-btn"
              style={{border:"none",background:tab===t.k?"rgba(255,255,255,.20)":"none",color:tab===t.k?"#fff":"rgba(255,255,255,.60)",fontWeight:tab===t.k?900:600,cursor:"pointer",borderBottom:tab===t.k?"3px solid #38bdf8":"3px solid transparent",borderRadius:"8px 8px 0 0",transition:"all .15s",fontFamily:"inherit",display:"flex",alignItems:"center",gap:6}}>
              <TabIcon k={t.k}/>{t.l}
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
              <span style={{flexShrink:0,display:"flex"}}><TabIcon k={t.k}/></span>
              <span>{t.l}</span>
              {t.k==="jadwal" && ops.filter((o:any)=>o.status==="ongoing"&&o.date===toLocalISODate(new Date())).length>0 && (
                <span style={{marginLeft:"auto",background:"#F44336",color:"#fff",fontSize:9,fontWeight:700,padding:"2px 6px",borderRadius:10}}>{ops.filter((o:any)=>o.status==="ongoing"&&o.date===toLocalISODate(new Date())).length}</span>
              )}
            </button>
          ))}
          <div style={{flex:1}}/>
          {TABS.filter(t=>t.k==="arsip").map(t=>(
            <button key={t.k} data-tab={t.k} onClick={handleTabClickEvt}
              className={`kb-sidebar-btn${tab===t.k?" active":""}`}>
              <span style={{flexShrink:0,display:"flex"}}><TabIcon k={t.k}/></span>
              <span>{t.l}</span>
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
          <div key={tab} className="kb-content-inner">
            {content}
          </div>
        </div>
      </div>

      {/* PIN Management Modal (Admin only) */}
      {showPinMgmt && role==="admin" && (
        <div className="kb-modal-backdrop" style={{position:"fixed",inset:0,background:"rgba(15,23,42,.45)",backdropFilter:"blur(6px)",WebkitBackdropFilter:"blur(6px)",zIndex:9000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
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
        <div className="kb-toast" style={{position:"fixed",bottom:24,left:"50%",transform:"translateX(-50%)",background:toast.color,color:"#fff",padding:"11px 22px",borderRadius:14,fontSize:13,fontWeight:700,zIndex:9999,boxShadow:"0 4px 20px rgba(0,0,0,.2)",whiteSpace:"nowrap",maxWidth:"90vw",textAlign:"center"}}>
          {toast.msg}
        </div>
      )}

      {/* Asisten Virtual — bubble melayang, global di semua tab setelah login */}
      <FloatingAssistant autoSaveOcrBatch={autoSaveOcrBatch} showToast={showToast}/>
    </div>
  );
}
