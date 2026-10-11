/* src/services/supabase.ts
   Inisialisasi Supabase client & fungsi audit log terpusat (AmbilLogMedis).
   Dipindah dari App.tsx (langkah 4 refactor — pisah services), TIDAK ADA
   perubahan logika. Ini SATU-SATUNYA tempat createClient() dipanggil —
   semua bagian app lain (termasuk services/dropbox.ts) mengimpor
   SUPA_CLIENT dari sini, bukan membuat instance baru. */

import { createClient } from "@supabase/supabase-js";
import type { SupabaseConfig } from "../types";
import { gHex16, fWIB } from "../utils";

/* ─── KONFIGURASI SUPABASE (via environment variable, BUKAN hardcode) ──
   VITE_SUPABASE_URL & VITE_SUPABASE_ANON_KEY WAJIB diset di .env.local
   (development) atau environment variable platform hosting (production).
   anon key ini memang didesain untuk boleh terlihat di client oleh
   Supabase (akses sesungguhnya dikontrol via RLS), namun tetap tidak boleh
   di-hardcode di source: hardcode membuat rotasi key sulit, membuat key
   bocor ke setiap fork/clone repo, dan menyamarkan kebutuhan RLS yang benar.
   ─────────────────────────────────────────────────────────────────── */
export const SUPA_URL    = (import.meta as any).env?.VITE_SUPABASE_URL as string | undefined;
export const SUPA_ANON   = (import.meta as any).env?.VITE_SUPABASE_ANON_KEY as string | undefined;
if (!SUPA_URL || !SUPA_ANON) {
  // eslint-disable-next-line no-console
  console.error(
    "[Konfigurasi] VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY belum diset di .env.local. " +
    "Aplikasi tidak dapat terhubung ke Supabase sampai env var ini dikonfigurasi."
  );
}
export const SUPA_CLIENT = createClient(SUPA_URL || "", SUPA_ANON || "");

export const defaultSupaCfg: SupabaseConfig = { url:"", anonKey:"", autoBackup:true, backupInterval:60, lastBackup:null, realtimeBackup:true, autoExcelBackup:true, lastExcelBackup:null, lastExcelBackupTs:null };

/**
 * AmbilLogMedis — audit log terpusat untuk seluruh aksi medis/administratif
 * di aplikasi (login, tambah/edit/hapus pasien, backup, dll).
 *
 * @param kategori  'LOGIN' | 'JADWAL_OPERASI' | 'REMINDER_WA' | 'BACKUP' |
 *                  'ROSTER_STAF' | 'LAPORAN_BACKUP' | 'MANAJEMEN_STAF'
 * @param operator  'Admin' | 'Perawat' | 'Sistem' | 'Unknown'
 * @param aksi      Nama tindakan spesifik: 'LOGIN_SUKSES', 'TAMBAH_PASIEN', dll.
 * @param detail    Deskripsi teks ringkas
 * @param metadata  Payload JSON forensik (before/after data, optional)
 *
 * PENTING: Fungsi ini TIDAK menggunakan await — eksekusi Turso berjalan
 * asinkronus di background, TIDAK memblokir event loop React.
 */
export const AmbilLogMedis = (
  kategori: string,
  operator: string,
  aksi:     string,
  detail:   string,
  metadata: object = {}
): void => {
  /* FIX AUDIT #29: tidak lagi bicara ke Turso langsung dari client.
     Payload dikirim ke Edge Function "turso-audit-proxy"; kredensial
     Turso tidak pernah menyentuh browser. Tetap fire-and-forget agar
     tidak memblokir UI, dan gagal-senyap (best-effort logging) — audit
     log yang gagal terkirim TIDAK boleh menghalangi alur bisnis utama. */
  try {
    SUPA_CLIENT.functions.invoke("turso-audit-proxy", {
      body: {
        id: gHex16(),
        waktu: fWIB(),
        kategori,
        operator,
        aksi,
        detail,
        metadata,
      },
    }).catch((err: unknown) => console.error("[TURSO-PROXY] audit log error:", err));
  } catch (err) {
    console.error("[TURSO-PROXY] audit log invoke error:", err);
  }
};
