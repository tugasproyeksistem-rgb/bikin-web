/* src/types/views.ts
   Tipe props untuk komponen-komponen View utama tiap tab (Jadwal, Daftar,
   Laporan, Kirim WA, Staf, Arsip) serta ScanFormulirProps. Dipindah dari
   App.tsx, tidak ada perubahan logika. Bergantung ke operation.ts,
   callbacks.ts, ocr.ts, monitoring.ts, config.ts. */

import type { Operation, StaffMember, RosterEntry, Notif } from "./operation";
import type { ShowToastFn, UpsertOneFn, DeleteFromSupaFn, UpsertBulkFn } from "./callbacks";
import type { OcrExtractResult } from "./ocr";
import type { MonitoringEntry, MonitoringCfg } from "./monitoring";
import type { DropboxConfig } from "./config";

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
  /* PENAMBAHAN #3 (skeleton loading): true selama fetch data AWAL dari
     Supabase (loadAllFromSupa) belum selesai. Dipakai supaya empty-state
     "Tidak ada jadwal" TIDAK ditampilkan keliru saat data sebenarnya
     masih dalam proses dimuat — sebelumnya `ops.length===0` dianggap
     "memang kosong" tanpa membedakan dari "belum sempat dimuat". */
  initialLoading: boolean;
}

export interface SaveOpFnArgs {
  opForm: Partial<Operation>;
  editingOp: Operation | null;
  setOpErrors: React.Dispatch<React.SetStateAction<Record<string, string>>>;
  setDupWarn: React.Dispatch<React.SetStateAction<boolean>>;
  resetOp: () => void;
  /* SARAN PENINGKATAN #3 (audit trail OCR): nilai ASLI hasil OCR untuk
     field yang sempat di-auto-fill, dikumpulkan ViewDaftar (lihat
     applyOcrResult). Opsional — undefined kalau entri ini tidak berasal
     dari scan OCR sama sekali (input manual biasa), dalam hal ini
     saveOpFn tidak mencatat log audit OCR apa pun. */
  ocrOriginalValues?: Record<string,string>;
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
  autoSaveOcrBatch: (results: OcrExtractResult[]) => void;
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

export interface ScanFormulirProps {
  onResults: (results: OcrExtractResult[]) => void;
  showToast: ShowToastFn;
}
