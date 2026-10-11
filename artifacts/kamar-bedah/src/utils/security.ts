/* src/utils/security.ts
   Fungsi utilitas murni untuk keamanan & sanitasi — hashing PIN,
   perbandingan constant-time, ID acak, parsing JSON aman, dan sanitasi
   teks/HTML/nama/RM. Dipindah dari App.tsx (langkah 3 refactor — pisah
   utils, batch 1: security & sanitasi), TIDAK ADA perubahan logika,
   termasuk seluruh FIX AUDIT & FIX BUG yang sudah ada di dalamnya. */

/** fWIB — format Date ke string WIB: "YYYY-MM-DD HH:MM:SS" */
export function fWIB(d: Date = new Date()): string {
  const wib = new Date(d.getTime() + 7 * 60 * 60 * 1000);
  return wib.toISOString().replace("T"," ").slice(0, 19);
}

/** gHex16 — 16-char hex string sebagai ID audit unik (crypto.randomUUID fallback) */
export function gHex16(): string {
  try { return crypto.randomUUID().replace(/-/g,"").slice(0,16); }
  catch { return Math.random().toString(36).slice(2,10) + Math.random().toString(36).slice(2,10); }
}

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
export const safeJSONParse = (data: any, fallback: any = {}): any => {
  const guard = (v: any) => (v !== null && typeof v === "object") ? v : fallback;
  if (typeof data !== "string") return guard(data ?? fallback);
  try { return guard(JSON.parse(data)); } catch (e) { console.error("[safeJSONParse] Data JSON korup:", e); return fallback; }
};

/* ─── PIN HASHING (FIX AUDIT #4 — CRITICAL: PIN Plaintext) ──────────────
   PIN sebelumnya disimpan plaintext di Supabase (kamar_bedah_config).
   Sekarang di-hash dengan PBKDF2-SHA256 (100k iterasi) + salt acak per-PIN
   menggunakan Web Crypto API sebelum dikirim/disimpan. Format string yang
   disimpan: "<salt_hex>:<hash_hex>". Verifikasi PIN dilakukan dengan
   me-re-derive hash dari salt yang tersimpan lalu membandingkan hasilnya —
   PIN asli tidak pernah disimpan maupun dikirim dalam bentuk plaintext
   ke Supabase. */
export async function hashPin(pin: string, saltHex?: string): Promise<string> {
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
export function constantTimeEqual(a: string, b: string): boolean {
  const maxLen = Math.max(a.length, b.length);
  let result = a.length === b.length ? 0 : 1;
  for (let i = 0; i < maxLen; i++) {
    const ca = i < a.length ? a.charCodeAt(i) : 0;
    const cb = i < b.length ? b.charCodeAt(i) : 0;
    result |= ca ^ cb;
  }
  return result === 0;
}

/* verifyPin: bandingkan PIN polos yang diketik user terhadap string
   "<salt>:<hash>" yang tersimpan. Mendukung mode legacy (plaintext lama,
   tanpa format salt:hash) untuk migrasi mulus dari data existing. */
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
export async function verifyPin(pin: string, stored: string | undefined | null): Promise<boolean> {
  if (!stored) return false;
  const isLegacyPlaintext = !stored.includes(":");
  /* Selalu derive PBKDF2 (pakai salt acak baru bila legacy) supaya biaya
     komputasi seragam antara kedua cabang — hasil dummy ini TIDAK dipakai
     untuk verifikasi legacy, hanya untuk menyamakan waktu eksekusi. */
  const dummyDeriveForTiming = isLegacyPlaintext ? hashPin(pin) : Promise.resolve("");
  if (isLegacyPlaintext) {
    const [legacyMatch] = await Promise.all([
      Promise.resolve(constantTimeEqual(pin, stored)),
      dummyDeriveForTiming
    ]);
    return legacyMatch; // legacy plaintext fallback (auto-migrasi di pemanggil)
  }
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
export const gId = (): string => {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") return crypto.randomUUID();
  const bytes = new Uint8Array(16);
  if (typeof crypto !== "undefined" && typeof crypto.getRandomValues === "function") crypto.getRandomValues(bytes);
  else for (let i = 0; i < 16; i++) bytes[i] = Math.floor(Math.random() * 256);
  bytes[6] = (bytes[6] & 0x0f) | 0x40; // version 4
  bytes[8] = (bytes[8] & 0x3f) | 0x80; // variant 10
  const hex = Array.from(bytes, b => b.toString(16).padStart(2, "0")).join("");
  return `${hex.slice(0,8)}-${hex.slice(8,12)}-${hex.slice(12,16)}-${hex.slice(16,20)}-${hex.slice(20)}`;
};

export const isValidUUID = (s: any): boolean => typeof s === "string" && /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(s);

/* FIX BUG #34 (audit lanjutan): strip-tag SATU KALI (`replace(/<[^>]*>/g,"")`
   sekali jalan) secara teoretis bisa lolos untuk pola tag bersarang yang
   dikonstruksi khusus untuk "reformasi" tag setelah 1 pass dihapus. Sekarang
   diulang (loop) sampai tidak ada lagi perubahan, menutup celah itu. Catatan:
   ini pertahanan lapis KEDUA untuk teks bebas (nama, keterangan) — untuk
   HTML yang benar-benar dirender (export Word), pertahanan UTAMA tetap
   escHtml() yang sudah proper-escape sejak FIX AUDIT #3, bukan strip-tag. */
export const sanitize  = (s: any) => {
  let out = String(s||"");
  let prev: string;
  do { prev = out; out = out.replace(/<[^>]*>/g,""); } while (out !== prev);
  return out.replace(/['"`;]/g,"").trim();
};
export const maskName  = (n: string,a: boolean) => !a||!n ? n??"" : String(n).replace(/\S+/g,(w:string)=>w[0]+"*".repeat(Math.max(1,w.length-1)));
/* FIX BUG #33 (audit lanjutan): SEBELUMNYA slice(0,2)+"****"+slice(-2) untuk
   RM pendek (mis. 2-3 digit) menghasilkan tampilan yang JUSTRU MEMBOCORKAN
   SELURUH nomor RM dua kali (mis. RM "12" → "12****12") — bukan cuma salah
   tampilan, tapi gagal fungsi utamanya (privacy mode). Sekarang RM dengan
   panjang <=4 digit di-mask TOTAL (semua jadi bintang), karena tidak cukup
   panjang untuk disamarkan sebagian tanpa membocorkan seluruh isinya. */
export const maskRM    = (r: string,a: boolean) => {
  if (!a || !r) return r ?? "";
  const s = String(r);
  return s.length <= 4 ? "*".repeat(s.length) : s.slice(0,2)+"****"+s.slice(-2);
};

/* FIX AUDIT #3 (CRITICAL — XSS pada Export Word): Semua teks dinamis
   (judul, header, isi sel) WAJIB di-escape sebelum disisipkan ke string
   HTML. Tanpa ini, data pasien yang mengandung "<script>" atau atribut
   "onerror=" bisa tereksekusi saat file .doc dibuka di Word/LibreOffice
   (stored XSS). */
export const escHtml = (s: any): string =>
  String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
