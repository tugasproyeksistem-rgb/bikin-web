/* src/types/roster.ts
   Tipe untuk fitur roster generator & master data perawat. Dipindah dari
   App.tsx, tidak ada perubahan logika. Bergantung ke callbacks.ts (fungsi
   Supabase) dan config.ts (DropboxConfig). */

import type { ShowToastFn, UpsertOneFn, DeleteFromSupaFn } from "./callbacks";
import type { DropboxConfig } from "./config";

export type NurseType = "karu" | "senior" | "junior";

export interface RosterNurse {
  id: string;
  nik: string;
  pk: string;
  name: string;
  sisaCuti: number;
  tipe: NurseType;
}

export interface RosterGenProps {
  showToast: ShowToastFn;
  upsertOneToSupa: UpsertOneFn;
  deleteFromSupa: DeleteFromSupaFn;
  dbxCfg: DropboxConfig;
}

/* Master Perawat — interface (terpisah dari RosterNurse).
   FIX (Audit Sentralisasi DB): localStorage DIHAPUS TOTAL sebagai sumber
   data. Tabel Supabase "kb_roster_master" adalah satu-satunya sumber
   kebenaran (1 row per perawat: id, data jsonb, updated_at), konsisten
   dengan pola tabel kb_* lain di file ini (kb_operasi, kb_staf, dst). */
export interface NurseMaster {
  id: string;
  nik: string;
  name: string;
  pk: string;
}
