/* src/types/callbacks.ts
   Tipe fungsi callback yang dipakai berulang lintas komponen — toast,
   upsert/delete Supabase, dan payload realtime. Dipindah dari App.tsx
   (langkah 1 refactor — pisah types), tidak ada perubahan logika. */

/** Fungsi umum untuk menampilkan toast */
export type ShowToastFn = (msg: string, color?: string) => void;

/** Upsert satu record ke Supabase */
export type UpsertOneFn = (table: string, data: Record<string, unknown>) => Promise<{ok: boolean; error?: string}>;

/** Hapus satu record dari Supabase */
export type DeleteFromSupaFn = (table: string, id: string) => Promise<{ok: boolean; error?: string}>;

/** Upsert banyak record ke Supabase sekaligus */
export type UpsertBulkFn = (table: string, rows: Record<string, unknown>[]) => Promise<{ok: boolean; error?: string}>;

/* FIX AUDIT #42 (LOW — Type Safety pada Payload Realtime): sebelumnya
   applyDelta(table, payload:any) menerima payload Supabase Realtime tanpa
   tipe sama sekali, sehingga kalau skema kolom berubah (mis. nama kolom
   `data`/`updated_at` diganti saat migrasi DB), TypeScript tidak akan
   memperingatkan apa pun saat build — error baru akan muncul diam-diam di
   runtime production. RtPayload mendeskripsikan bentuk minimal payload
   `postgres_changes` dari Supabase (baris disimpan sebagai jsonb di kolom
   `data`, sesuai pola tabel kb_* di app ini). */
export interface RtPayload {
  eventType: "INSERT" | "UPDATE" | "DELETE";
  new: { id: string; data?: any } | null;
  old: { id: string } | null;
}
