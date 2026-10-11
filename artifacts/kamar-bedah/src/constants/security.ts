/* src/constants/security.ts
   Konstanta keamanan: proteksi formula-injection Excel (FORMULA_TRIGGER_CHARS,
   lihat FIX BUG #39) dan pengaturan lockout percobaan PIN. Dipindah dari
   App.tsx, tidak ada perubahan nilai. */

/* FIX BUG #39 (audit lanjutan): "|" (pipe, pemicu DDE formula di sebagian
   versi Excel lama) ditambahkan supaya konsisten dengan sanitizeForCSV(). */
export const FORMULA_TRIGGER_CHARS = ["=", "+", "-", "@", "\t", "\r", "|"];

export const PIN_MAX_ATTEMPTS = 5;
export const PIN_LOCKOUT_MS   = 60_000;
/* FIX AUDIT #31 (HIGH — Bypass Lockout via Refresh): state attempts/lockedUntil
   di-persist ke sessionStorage (key ini) agar bertahan lewat refresh (F5) —
   mitigasi client-side. Perbaikan PERMANEN tetap perlu rate-limit di server. */
export const PIN_LOCK_STORAGE_KEY = "kb_pin_lock_state_v1";
