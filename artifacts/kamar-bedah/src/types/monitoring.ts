/* src/types/monitoring.ts
   Tipe untuk fitur monitoring suhu/kelembaban ruang kamar bedah. Dipindah
   dari App.tsx, tidak ada perubahan logika. */

export interface MonitoringEntry {
  id: string;
  ruang: string;        // nama ruangan (hardcoded dari MON_ROOMS)
  tanggal: string;
  jam: string;
  suhu: number;
  kelembaban: number;
  petugas: string;
  createdAt: string;
}

export interface MonitoringCfg {
  suhuMin: number;
  suhuMax: number;
  rhMin: number;
  rhMax: number;
  lokasiRuang: string;
  kepalaKamarBedah: string;
}
