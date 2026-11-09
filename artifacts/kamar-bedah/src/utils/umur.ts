/* src/utils/umur.ts
   Hitung umur (tahun + bulan) dari tanggal lahir (YYYY-MM-DD) pada tanggal acuan (tanggal operasi; kalau kosong, hari ini). */
const ISO = /^(\d{4})-(\d{2})-(\d{2})$/;
export function umurDariLahir(lahir: string | undefined | null, acuan?: string | null): { tahun: number; bulan: number } | null {
  const m = ISO.exec(String(lahir || ""));
  if (!m) return null;
  const ly = +m[1], lm = +m[2], ld = +m[3];
  const dt = new Date(ly, lm - 1, ld);
  if (dt.getFullYear() !== ly || dt.getMonth() !== lm - 1 || dt.getDate() !== ld) return null;
  let ay: number, am: number, ad: number;
  const a = ISO.exec(String(acuan || ""));
  if (a) { ay = +a[1]; am = +a[2]; ad = +a[3]; }
  else { const n = new Date(); ay = n.getFullYear(); am = n.getMonth() + 1; ad = n.getDate(); }
  let bln = (ay - ly) * 12 + (am - lm);
  if (ad < ld) bln -= 1;
  if (bln < 0) return null;
  return { tahun: Math.floor(bln / 12), bulan: bln % 12 };
}
