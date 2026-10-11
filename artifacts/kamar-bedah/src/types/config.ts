/* src/types/config.ts
   Tipe konfigurasi backup. SENGAJA hanya interface-nya yang dipindah ke
   sini — nilai default (defaultSupaCfg/defaultDbxCfg) dan fungsi proxy
   (callDbxProxy, dropboxUpload, dst) masih di App.tsx karena itu tergolong
   "services" (langkah 4 refactor), bukan "types" (langkah 1). Memisah
   keduanya sekarang akan mencampur 2 kategori berbeda dalam 1 langkah. */

export interface SupabaseConfig {
  url: string;
  anonKey: string;
  autoBackup: boolean;
  backupInterval: number;
  lastBackup: string | null;
  realtimeBackup: boolean;
  autoExcelBackup: boolean;
  lastExcelBackup: string | null;
  lastExcelBackupTs: number | null;
}

export interface DropboxConfig {
  path: string;
  autoBackup: boolean;
  backupInterval: number;
  realtimeBackup: boolean;
  lastBackup: string | null;
  autoExcelBackup: boolean;
  lastExcelBackup: string | null;
  lastExcelBackupTs: number | null;
}
