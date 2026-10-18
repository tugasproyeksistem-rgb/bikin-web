/* Kolom tambahan Excel Operasi: field yang diinput tapi sebelumnya tidak tertulis.
   Dipakai bersama oleh semua pembuat Excel Operasi supaya kolomnya selalu sama. */

export const KOLOM_TAMBAHAN_OPS: string[] = [
  "Umur (bulan)", "Alergi", "Gol. Darah", "Ruang Asal", "Kebutuhan Khusus", "Alasan Batal", "Permintaan",
];

const aman = (v: unknown): string => {
  const s = String(v ?? "");
  return /^([=+@\t\r]|-.)/.test(s) ? "'" + s : s;
};

export function nilaiTambahanOps(o: any): string[] {
  const req = Array.isArray(o?.requests)
    ? o.requests.map((r: any) => `${r?.time || ""}: ${r?.text || ""}`).join("; ")
    : "";
  return [o?.ageMonths, o?.allergy, o?.bloodType, o?.ruangAsal, o?.specialNeeds, o?.cancelReason, req].map(aman);
}
