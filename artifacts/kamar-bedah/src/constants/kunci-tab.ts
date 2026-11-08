/* src/constants/kunci-tab.ts
   Tab grup MUTU terlihat oleh perawat tetapi terkunci (tidak bisa diklik) sampai 1 Januari 2027 pukul 00.00 WIB.
   Admin tidak terkunci. Dibaca setiap kali menu/tab dirender, jadi terbuka otomatis pada tanggal itu. */
export const MUTU_MULAI = Date.parse("2027-01-01T00:00:00+07:00");
export const TAB_MUTU: string[] = ["mutu_ok", "implant_tulang", "implant_mata", "implant_hernia"];
export const tabTerkunci = (k: string, role: string | null | undefined): boolean =>
  role !== "admin" && TAB_MUTU.indexOf(k) >= 0 && Date.now() < MUTU_MULAI;
