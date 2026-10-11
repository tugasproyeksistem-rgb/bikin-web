/* src/utils/storage.ts
   Fungsi utilitas murni untuk penyimpanan browser (sessionStorage/
   localStorage) — state lockout percobaan PIN, dan antrean offline
   kb_operasi. Dipindah dari App.tsx (langkah 3 refactor — pisah utils,
   batch 5), TIDAK ADA perubahan logika, termasuk FIX AUDIT #31 dan
   catatan desain lainnya. */

import type { OfflineQueueItem } from "../types";
import { PIN_LOCK_STORAGE_KEY, LS_OFFLINE_QUEUE_KEY } from "../constants";

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
export function readPinLockState(): {attempts:number; lockedUntil:number|null} {
  try {
    const raw = sessionStorage.getItem(PIN_LOCK_STORAGE_KEY);
    if(!raw) return {attempts:0, lockedUntil:null};
    const parsed = JSON.parse(raw);
    /* buang lockout yang sudah lewat waktunya supaya tidak nyangkut permanen */
    if(parsed.lockedUntil && Date.now() >= parsed.lockedUntil) return {attempts:0, lockedUntil:null};
    return { attempts: Number(parsed.attempts)||0, lockedUntil: parsed.lockedUntil||null };
  } catch { return {attempts:0, lockedUntil:null}; }
}
export function writePinLockState(attempts:number, lockedUntil:number|null){
  try { sessionStorage.setItem(PIN_LOCK_STORAGE_KEY, JSON.stringify({attempts,lockedUntil})); } catch {}
}

/* ─── ANTREAN OFFLINE (kb_operasi) ───────────────────────────────────
   Setiap upsert kb_operasi yang gagal (offline atau error network) otomatis
   masuk antrean localStorage yang PERSISTEN (bertahan meski tab ditutup),
   dan dicoba lagi otomatis begitu browser mendeteksi koneksi kembali
   online — pola yang sama persis semangatnya dengan yang sudah dipakai di
   roster generator. */
export function readOfflineQueue(): OfflineQueueItem[] {
  try {
    const raw = localStorage.getItem(LS_OFFLINE_QUEUE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch { return []; }
}

export function writeOfflineQueue(items: OfflineQueueItem[]): void {
  try { localStorage.setItem(LS_OFFLINE_QUEUE_KEY, JSON.stringify(items)); }
  catch { /* localStorage penuh/disabled — abaikan senyap, tidak fatal */ }
}

/** Tambahkan satu entri pasien ke antrean offline (dedup by id — kalau id
 *  sudah ada di antrean, timpa payload lama dengan yang terbaru). */
export function enqueueOfflineOp(id: string, payload: any): void {
  const q = readOfflineQueue().filter(item => item.id !== id);
  q.push({ id, payload, queuedAt: Date.now() });
  writeOfflineQueue(q);
}

/** Hapus satu entri dari antrean offline (dipanggil setelah upsert ke
 *  Supabase benar-benar berhasil). */
export function dequeueOfflineOp(id: string): void {
  const q = readOfflineQueue().filter(item => item.id !== id);
  writeOfflineQueue(q);
}

/** Jumlah entri yang masih menunggu sinkron — dipakai untuk indikator
 *  kecil di UI ("N data belum tersinkron") supaya perawat tahu ada data
 *  yang masih perlu koneksi untuk benar-benar tersimpan permanen. */
export function countOfflineQueue(): number {
  return readOfflineQueue().length;
}
