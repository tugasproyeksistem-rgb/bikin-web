/* src/constants/time.ts
   Konstanta terkait waktu & tanggal: batas pagi/siang/sore/malam (lihat
   FIX BUG #35), basis serial-date Excel, dan nama bulan Indonesia dalam
   2 varian kapitalisasi (dipakai di tempat berbeda). Dipindah dari
   App.tsx, tidak ada perubahan nilai. */

/* FIX BUG #35 (audit lanjutan): magic number 660/900/1110 diganti konstanta
   bernama supaya batas pagi/siang/sore/malam langsung terbaca. */
export const BATAS_SIANG_MENIT = 11*60;      // 11:00
export const BATAS_SORE_MENIT  = 15*60;      // 15:00
export const BATAS_MALAM_MENIT = 18*60+30;   // 18:30

export const EXCEL_EPOCH = Date.UTC(1899, 11, 30); // basis serial date Excel (termasuk leap-year bug 1900)

export const BULAN_ID = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
/* Varian UPPERCASE — dipakai di header Excel roster (SUMMARY_COLS dkk),
   terpisah dari BULAN_ID (title-case) yang dipakai di label UI biasa. */
export const BULAN_NAMA_ID = ["JANUARI","FEBRUARI","MARET","APRIL","MEI","JUNI","JULI","AGUSTUS","SEPTEMBER","OKTOBER","NOVEMBER","DESEMBER"];
