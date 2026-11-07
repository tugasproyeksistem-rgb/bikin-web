/* src/constants/fitur-tanggal.ts
   Fitur yang baru aktif pada tanggal tertentu. Dibaca setiap kali form dirender, jadi aktif otomatis
   mulai 1 Januari 2027 pukul 00.00 WIB. Uji developer: localStorage.setItem("kb_uji_implant","1") di Console. */
export const IMPLANT_MULAI = Date.parse("2027-01-01T00:00:00+07:00");
export const implantAktif = (): boolean => {
  if (Date.now() >= IMPLANT_MULAI) return true;
  try { return typeof localStorage !== "undefined" && localStorage.getItem("kb_uji_implant") === "1"; } catch { return false; }
};
