/* src/constants/monitoring.ts
   Konstanta untuk fitur monitoring suhu/kelembaban kamar bedah: jam
   pengecekan, daftar ruangan, warna tema, dan warna per-ruangan untuk
   grafik. Dipindah dari App.tsx, tidak ada perubahan nilai. */

export const MON_JAMS  = ["07:00","14:00","21:00"];
/* MON_ROOMS: 3 ruangan hardcoded sesuai spesifikasi. Urutan ini dipakai
   di seluruh UI, Excel, dan ID deterministik — JANGAN diubah urutannya
   tanpa migrasi data. */
export const MON_ROOMS = ["Kamar Bedah 1","Kamar Bedah 2","Ruang Instrumen"];
export const MON_MONTHS_ID = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
export const MC = { pri:"#0369A1", priL:"#0284C7", priBg:"#EFF6FF", ok:"#16A34A", okBg:"#DCFCE7", err:"#DC2626", errBg:"#FEE2E2", warn:"#D97706" };
export const MON_ROOM_COLORS = ["#0369A1","#7C3AED","#B45309"];
