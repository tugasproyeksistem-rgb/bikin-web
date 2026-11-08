/* src/constants/nav.ts
   Daftar tab navigasi utama beserta role yang boleh mengaksesnya. Dipindah
   dari App.tsx, tidak ada perubahan nilai.
   FIX AUDIT #12 (asli): ALL_TABS sebagai module-level constant (stabil
   antar render, bukan didefinisikan ulang tiap render di dalam komponen). */
/* Tab MUTU kamar bedah: admin bisa melihat sekarang; perawat otomatis mulai 1 Januari 2027 pukul 00.00 WIB (dibaca saat halaman dimuat). */
const MUTU_AKTIF = Date.now() >= Date.parse("2027-01-01T00:00:00+07:00");
export const ALL_TABS: Array<{k:string; l:string; roles:Array<"admin"|"perawat">}> = [
  {k:"home",      l:"Home",       roles:["admin","perawat"]},
  {k:"jadwal",    l:"Jadwal",     roles:["admin","perawat"]},
  {k:"daftar",    l:"Daftar",     roles:["admin","perawat"]},
  {k:"laporan",   l:"Laporan Kepala Jaga",    roles:["admin","perawat"]},
  {k:"laporan_harian", l:"Laporan Harian OK", roles:["admin","perawat"]},
  {k:"beset_jadwal", l:"Beset Jadwal", roles:["admin","perawat"]},
  {k:"pengumuman", l:"Pengumuman", roles:["admin"]},
  {k:"jadwal_ok", l:"Jadwal & Peserta", roles:["admin","perawat"]},
  {k:"mutu_ok", l:"Pemantauan Kamar Bedah", roles: MUTU_AKTIF ? ["admin","perawat"] : ["admin"]},
  {k:"implant_tulang", l:"IMPLANT TULANG", roles: MUTU_AKTIF ? ["admin","perawat"] : ["admin"]},
  {k:"wa",        l:"Kirim WA",   roles:["admin","perawat"]},
  {k:"statistik", l:"Statistik",  roles:["admin","perawat"]},
  {k:"staf",      l:"Staf",       roles:["admin"]},
  {k:"lembur",    l:"Lembur",     roles:["admin","perawat"]},
  {k:"monitoring",l:"Monitoring", roles:["admin","perawat"]},
  {k:"roster_gen", l:"Generator",   roles:["admin"]},
  {k:"arsip",     l:"Arsip",      roles:["admin"]},
];
