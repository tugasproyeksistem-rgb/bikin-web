/* src/utils/laporan-harian.ts
   Logika murni tab "Laporan Harian OK": aturan shift, penyusunan baris
   jadwal/rencana/kandidat dari daftar operasi, teks preview, dan Excel.
   Aturan shift DIPUSATKAN di sini (SHIFT_BATAS + shiftDariJam) supaya
   mudah diubah tanpa menyentuh komponen. */

import * as XLSX from "xlsx";

export type ShiftKey = "pagi" | "siang" | "malam";
export const SHIFT_URUT: ShiftKey[] = ["pagi", "siang", "malam"];
export const SHIFT_LABEL: Record<ShiftKey, string> = { pagi: "Shift Pagi", siang: "Shift Siang", malam: "Shift Malam" };
/* Batas shift dalam menit sejak 00:00: pagi 07:00, siang 14:00, malam 20:30 */
export const SHIFT_BATAS = { pagiMulai: 2 * 60, siangMulai: 14 * 60, malamMulai: 20 * 60 + 30 };
export const STOK_OPSI: string[] = ["Tersedia", "Kurang"];
export const STATUS_OPSI: string[] = ["Terjadwal", "Proses", "Selesai", "Tunda", "Batal"];
export const ALAT_BUTIR: { k: string; l: string }[] = [
  { k: "lampu", l: "Lampu Operasi" },
  { k: "meja", l: "Meja Operasi" },
  { k: "mesinAnestesi", l: "Mesin Anestesi" },
  { k: "monitor", l: "Monitor" },
  { k: "autoclave", l: "Autoclave / Sterilisasi" },
  { k: "suction", l: "Suction" },
];
export const NAMA_HARI = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
export const NAMA_BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

export interface JadwalRow { opId: string; jam: string; pasien: string; jenis: string; operator: string; anestesi: string; status: string; catatan: string; cyto: boolean; }
export interface RencanaRow { opId: string; tanggal: string; jam: string; operator: string; anestesi: string; pasien: string; umur: string; kamar: string; jenis: string; alkes: string; keterangan: string; }
export interface KandidatRow { opId: string; tanggal: string; jam: string; dokter: string; pasien: string; umur: string; jenis: string; keterangan: string; }

export interface LaporanHarian {
  id: string; tanggal: string; hari: string;
  shift: ShiftKey; dari: ShiftKey; ke: ShiftKey; tanggalRencana: string;
  jumlahOperasi: string; adaCyto: "ya" | "tidak"; totalPasien: string;
  jadwal: JadwalRow[]; rencana: RencanaRow[]; kandidat: KandidatRow[]; tanggalKandidat?: string;
  alat: Record<string, string>; alatLain: string; alatCatatan: string;
  instrumen: string; implan: string; alkesDipakai: string; alkesPersediaan: string; pa: string; stokLain: string; stokCatatan: string;
  kejadian: "ya" | "tidak"; uraian: string;
  lama: string; baru: string; kepala: string;
  operator: string; createdAt: string; updated_at: string;
}

/* ── Tanggal (waktu LOKAL, bukan UTC) ── */
export const pad2 = (n: number): string => String(n).padStart(2, "0");
export function isoLokal(d: Date): string { return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`; }
function pecahIso(iso: string): [number, number, number] {
  const p = String(iso || "").split("-").map(Number);
  return [p[0] || 1970, p[1] || 1, p[2] || 1];
}
export function geserHari(iso: string, n: number): string { const [y, m, d] = pecahIso(iso); return isoLokal(new Date(y, m - 1, d + n)); }
export function indeksHari(iso: string): number { const [y, m, d] = pecahIso(iso); return new Date(y, m - 1, d).getDay(); }
export function namaHari(iso: string): string { return NAMA_HARI[indeksHari(iso)] || ""; }
export function fmtTgl(iso: string): string { const [y, m, d] = pecahIso(iso); return `${pad2(d)}-${pad2(m)}-${y}`; }
export function tanggalPanjang(iso: string): string { return `${namaHari(iso)}, ${fmtTgl(iso)}`; }

/* ── Aturan shift ── */
export function menitDariJam(t: unknown): number | null {
  const m = String(t ?? "").trim().match(/^(\d{1,2})[.:](\d{2})/);
  if (!m) return null;
  const h = Number(m[1]); const mm = Number(m[2]);
  if (h > 23 || mm > 59) return null;
  return h * 60 + mm;
}
export function shiftDariJam(t: unknown, malamAktif: boolean): ShiftKey {
  const m = menitDariJam(t);
  if (m === null) return "siang";                                   /* jam kosong → Siang */
  if (m < (malamAktif ? 7 * 60 : SHIFT_BATAS.pagiMulai)) return malamAktif ? "malam" : "siang"; /* dini hari */
  if (m < SHIFT_BATAS.siangMulai) return "pagi";
  if (m < SHIFT_BATAS.malamMulai) return "siang";
  return malamAktif ? "malam" : "siang";
}
export function shiftAktif(malamAktif: boolean): ShiftKey[] { return malamAktif ? ["pagi", "siang", "malam"] : ["pagi", "siang"]; }
export function shiftBerikut(s: ShiftKey, malamAktif: boolean): ShiftKey {
  const list = shiftAktif(malamAktif);
  const i = list.indexOf(s);
  if (i < 0) return list[0];
  return list[(i + 1) % list.length];
}
/* Tanggal shift tujuan: kalau urutannya "mundur" (siang→pagi) berarti besok. */
export function tanggalTujuan(tanggal: string, dari: ShiftKey, ke: ShiftKey): string {
  return SHIFT_URUT.indexOf(ke) > SHIFT_URUT.indexOf(dari) ? tanggal : geserHari(tanggal, 1);
}

/* ── Penyusunan baris dari daftar operasi ── */
export function statusLabel(s: unknown): string {
  switch (s) { case "done": return "Selesai"; case "ongoing": return "Proses"; case "batal": return "Batal"; default: return "Terjadwal"; }
}
export function teksUmur(o: any): string {
  const a = String(o?.age ?? "").trim(); const b = String(o?.ageMonths ?? "").trim();
  const p: string[] = [];
  if (a) p.push(a + " th");
  if (b) p.push(b + " bln");
  return p.join(" ");
}
const jamUrut = (a: any, b: any): number => (menitDariJam(a?.time) ?? 9999) - (menitDariJam(b?.time) ?? 9999);

export function susunJadwal(ops: any[], tanggal: string, shift: ShiftKey, malamAktif: boolean): JadwalRow[] {
  return (ops || [])
    .filter((o: any) => o && o.date === tanggal && shiftYangMencakup(o.time, malamAktif).includes(shift))
    .sort(jamUrut)
    .map((o: any): JadwalRow => ({
      opId: String(o.id), jam: o.time || "", pasien: o.patient || "", jenis: o.procedure || "",
      operator: o.surgeon || "", anestesi: o.anesthesiologist || "", status: statusLabel(o.status),
      catatan: "", cyto: o.opType === "cyto",
    }));
}
export function susunRencana(ops: any[], tanggal: string, shift: ShiftKey, malamAktif: boolean): RencanaRow[] {
  return (ops || [])
    .filter((o: any) => o && o.date === tanggal && o.status !== "batal" && shiftYangMencakup(o.time, malamAktif).includes(shift))
    .sort(jamUrut)
    .map((o: any): RencanaRow => ({
      opId: String(o.id), tanggal: o.date || "", jam: o.time || "", operator: o.surgeon || "", anestesi: o.anesthesiologist || "",
      pasien: o.patient || "", umur: teksUmur(o), kamar: o.room || "", jenis: o.procedure || "", alkes: "", keterangan: "",
    }));
}
/* Kandidat = operasi besok (hari Sabtu: Minggu + Senin) */
export function susunKandidat(ops: any[], tanggal: string): KandidatRow[] {
  const tgl = [geserHari(tanggal, 1)];
  if (indeksHari(tanggal) === 6) tgl.push(geserHari(tanggal, 2));
  return (ops || [])
    .filter((o: any) => o && tgl.includes(o.date) && o.status !== "batal")
    .sort((a: any, b: any) => String(a.date).localeCompare(String(b.date)) || jamUrut(a, b))
    .map((o: any): KandidatRow => ({
      opId: String(o.id), tanggal: o.date || "", jam: o.time || "", dokter: o.surgeon || "", pasien: o.patient || "",
      umur: teksUmur(o), jenis: o.procedure || "", keterangan: "",
    }));
}

/* ── Teks preview (bisa dibacakan / disalin) ── */
const kondisiAlat = (v: string): string => (v === "baik" ? "Baik" : v === "tidak" ? "Tidak baik" : "-");
export function buatTeksLaporan(l: LaporanHarian, mask: (s: string) => string): string {
  const L: string[] = [];
  L.push("LAPORAN HARIAN KAMAR BEDAH");
  L.push(`${tanggalPanjang(l.tanggal)} - ${SHIFT_LABEL[l.shift]}`);
  L.push(`Serah terima: ${SHIFT_LABEL[l.dari]} ke ${SHIFT_LABEL[l.ke]}`);
  L.push(`Jumlah operasi: ${l.jumlahOperasi} | Cyto/darurat: ${l.adaCyto === "ya" ? "Ya" : "Tidak"} | Total pasien: ${l.totalPasien}`);
  L.push("");
  L.push("JADWAL OPERASI");
  if (!l.jadwal.length) { L.push("- Tidak ada acara operasi"); }
  l.jadwal.forEach((r, i) => {
    L.push(`${i + 1}. ${r.jam || "-"} ${mask(r.pasien)} - ${r.jenis}. Operator: ${r.operator || "-"}. Anestesi: ${r.anestesi || "-"}. Status: ${r.status}.${r.catatan ? " Catatan: " + r.catatan + "." : ""}`);
  });
  L.push("");
  L.push(`RENCANA OPERASI (${rencanaHariPenuh(l.dari, l.ke) ? "Semua operasi besok" : SHIFT_LABEL[l.ke]}, ${tanggalPanjang(l.tanggalRencana)})`);
  if (!l.rencana.length) { L.push("- Tidak ada acara operasi"); }
  l.rencana.forEach((r, i) => {
    L.push(`${i + 1}. ${r.jam || "-"} ${mask(r.pasien)}${r.umur ? " (" + r.umur + ")" : ""} - ${r.jenis}. Kamar: ${r.kamar || "-"}. Operator: ${r.operator || "-"}. Anestesi: ${r.anestesi || "-"}.${r.alkes ? " Alkes: " + r.alkes + "." : ""}${r.keterangan ? " Ket: " + r.keterangan + "." : ""}`);
  });
  L.push("");
  L.push(judulKandidat(l));
  if (!l.kandidat.length) { L.push("- Tidak ada acara operasi"); }
  l.kandidat.forEach((r, i) => {
    L.push(`${i + 1}. ${fmtTgl(r.tanggal)} ${r.jam || "-"} ${mask(r.pasien)}${r.umur ? " (" + r.umur + ")" : ""} - ${r.jenis}. Dokter: ${r.dokter || "-"}.`);
  });
  L.push("");
  L.push("KONDISI PERALATAN & KAMAR");
  ALAT_BUTIR.forEach((b) => L.push(`${b.l}: ${kondisiAlat(l.alat[b.k])}`));
  if (l.alatLain) { L.push(`Lain-lain: ${l.alatLain}`); }
  if (l.alatCatatan) { L.push(`Catatan: ${l.alatCatatan}`); }
  L.push("");
  L.push("STOK & BAHAN");
  L.push(`Instrumen: ${l.instrumen || "-"} | Implan/Bahan habis pakai: ${l.implan || "-"}`);
  if (l.alkesDipakai) { L.push(`Alkes dipakai: ${l.alkesDipakai}`); }
  if (l.alkesPersediaan) { L.push(`Alkes persediaan: ${l.alkesPersediaan}`); }
  if (l.pa) { L.push(`PA: ${l.pa}`); }
  if (l.stokLain) { L.push(`Lain-lain: ${l.stokLain}`); }
  if (l.stokCatatan) { L.push(`Catatan: ${l.stokCatatan}`); }
  L.push("");
  L.push("KEJADIAN / INSIDEN / KENDALA");
  L.push(l.kejadian === "ya" ? `Ada kejadian: ${l.uraian}` : "Tidak ada kejadian.");
  L.push("");
  L.push(`Petugas shift lama: ${l.lama} | Petugas shift baru: ${l.baru} | Mengetahui: ${l.kepala || "-"}`);
  return L.join("\n");
}

/* ── Excel ── */
function sx(v: any): any { return typeof v === "string" && /^[=+\-@\t\r]/.test(v) ? "'" + v : v; }
export const urutLaporanTurun = (a: LaporanHarian, b: LaporanHarian): number =>
  b.tanggal.localeCompare(a.tanggal) || SHIFT_URUT.indexOf(b.shift) - SHIFT_URUT.indexOf(a.shift);
export function filterLaporan(list: LaporanHarian[], mode: "tanggal" | "bulan" | "tahun", nilai: string): LaporanHarian[] {
  if (!nilai) return [];
  return list.filter((l) => (mode === "tanggal" ? l.tanggal === nilai : l.tanggal.startsWith(nilai)));
}
export function buatWorkbookLaporan(list: LaporanHarian[]): any {
  const wb = XLSX.utils.book_new();
  const urut = [...list].sort((a, b) => -urutLaporanTurun(a, b));
  if (!urut.length) {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["Belum ada laporan harian"]]), "Laporan");
    return wb;
  }
  const h1 = ["Tanggal", "Hari", "Shift", "Dari Shift", "Ke Shift", "Jumlah Operasi", "Ada Cyto", "Total Pasien",
    ...ALAT_BUTIR.map((b) => b.l), "Alat Lain-lain", "Catatan Alat", "Instrumen", "Implan/BHP", "Alkes Dipakai", "Alkes Persediaan", "PA",
    "Stok Lain-lain", "Catatan Stok", "Ada Kejadian", "Uraian Kejadian", "Petugas Shift Lama", "Petugas Shift Baru", "Kepala Kamar Bedah", "Disimpan"];
  const r1 = urut.map((l) => [fmtTgl(l.tanggal), l.hari, SHIFT_LABEL[l.shift], SHIFT_LABEL[l.dari], SHIFT_LABEL[l.ke], l.jumlahOperasi, l.adaCyto === "ya" ? "Ya" : "Tidak", l.totalPasien,
    ...ALAT_BUTIR.map((b) => kondisiAlat(l.alat[b.k])), sx(l.alatLain), sx(l.alatCatatan), l.instrumen, l.implan, sx(l.alkesDipakai), sx(l.alkesPersediaan), sx(l.pa),
    sx(l.stokLain), sx(l.stokCatatan), l.kejadian === "ya" ? "Ya" : "Tidak", sx(l.uraian), sx(l.lama), sx(l.baru), sx(l.kepala), l.updated_at]);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([h1, ...r1]), "Laporan");

  const r2: any[][] = [];
  urut.forEach((l) => l.jadwal.forEach((r) => r2.push([fmtTgl(l.tanggal), SHIFT_LABEL[l.shift], r.jam, sx(r.pasien), sx(r.jenis), sx(r.operator), sx(r.anestesi), r.status, sx(r.catatan)])));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["Tanggal", "Shift", "Jam", "Nama Pasien", "Jenis Operasi", "Dokter Operator", "Dokter Anestesi", "Status", "Catatan"], ...r2]), "Jadwal Operasi");

  const r3: any[][] = [];
  urut.forEach((l) => l.rencana.forEach((r) => r3.push([fmtTgl(l.tanggal), SHIFT_LABEL[l.shift], fmtTgl(r.tanggal), r.jam, sx(r.operator), sx(r.anestesi), sx(r.pasien), r.umur, sx(r.kamar), sx(r.jenis), sx(r.alkes), sx(r.keterangan)])));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["Tanggal Laporan", "Shift", "Tanggal", "Jam", "Dokter Operator", "Dokter Anestesi", "Nama Pasien", "Umur", "Kamar", "Jenis Operasi", "Alkes", "Keterangan"], ...r3]), "Rencana Operasi");
  return wb;
}

/* ── PDF minimal tanpa library: tiap halaman = satu gambar JPEG ── */
export function buatPdfDariJpeg(halaman: { jpeg: Uint8Array; w: number; h: number }[], pw = 595.28, ph = 841.89): Blob {
  const enc = new TextEncoder();
  const bagian: Uint8Array[] = [];
  let pos = 0;
  const tambah = (u: Uint8Array) => { bagian.push(u); pos += u.length; };
  const teks = (s: string) => tambah(enc.encode(s));
  const offs: number[] = [];
  const n = halaman.length;
  const total = 2 + 3 * n;
  teks("%PDF-1.4\n");
  offs[1] = pos; teks("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
  const kids = halaman.map((_, i) => `${3 + 3 * i} 0 R`).join(" ");
  offs[2] = pos; teks(`2 0 obj\n<< /Type /Pages /Kids [${kids}] /Count ${n} >>\nendobj\n`);
  halaman.forEach((p, i) => {
    const po = 3 + 3 * i; const co = po + 1; const io = po + 2;
    const hPt = (p.h * pw) / p.w; const y = ph - hPt;
    offs[po] = pos;
    teks(`${po} 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${pw.toFixed(2)} ${ph.toFixed(2)}] /Resources << /XObject << /Im0 ${io} 0 R >> >> /Contents ${co} 0 R >>\nendobj\n`);
    const isi = `q ${pw.toFixed(2)} 0 0 ${hPt.toFixed(2)} 0 ${y.toFixed(2)} cm /Im0 Do Q`;
    offs[co] = pos;
    teks(`${co} 0 obj\n<< /Length ${isi.length} >>\nstream\n${isi}\nendstream\nendobj\n`);
    offs[io] = pos;
    teks(`${io} 0 obj\n<< /Type /XObject /Subtype /Image /Width ${p.w} /Height ${p.h} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${p.jpeg.length} >>\nstream\n`);
    tambah(p.jpeg);
    teks("\nendstream\nendobj\n");
  });
  const xref = pos;
  let x = `xref\n0 ${total + 1}\n0000000000 65535 f \n`;
  for (let i = 1; i <= total; i++) x += `${String(offs[i]).padStart(10, "0")} 00000 n \n`;
  teks(x + `trailer\n<< /Size ${total + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF`);
  return new Blob(bagian as unknown as BlobPart[], { type: "application/pdf" });
}

/* Jam dinas: pagi 07:00-14:00, siang 13:30-20:30 (13:30-13:59 tumpang tindih = jam serah terima,
   operasi di rentang itu muncul di KEDUA shift), malam 20:30-07:00 (kalau malam mati, masuk siang).
   shiftDariJam di atas hanya dipakai untuk menebak shift awal dari jam sekarang. */
export const SHIFT_JAM = {
  pagiMulai: 2 * 60, pagiSelesai: 14 * 60,
  siangMulai: 13 * 60 + 30, siangSelesai: 20 * 60 + 30,
  malamMulai: 20 * 60 + 30, malamSelesai: 2 * 60,
};
export function shiftYangMencakup(t: unknown, malamAktif: boolean): ShiftKey[] {
  const m = menitDariJam(t);
  if (m === null) return ["siang"];
  const hasil: ShiftKey[] = [];
  const pagiMulai = malamAktif ? 7 * 60 : SHIFT_JAM.pagiMulai;
  const malamSelesai = malamAktif ? 7 * 60 : SHIFT_JAM.malamSelesai;
  if (m >= pagiMulai && m < SHIFT_JAM.pagiSelesai) hasil.push("pagi");
  if (m >= SHIFT_JAM.siangMulai && m < SHIFT_JAM.siangSelesai) hasil.push("siang");
  if (m >= SHIFT_JAM.malamMulai || m < malamSelesai) hasil.push(malamAktif ? "malam" : "siang");
  return hasil;
}

/* Siang -> Pagi: Rencana Operasi berisi SEMUA operasi besok (satu hari penuh), Kandidat dikosongkan. */
export function rencanaHariPenuh(dari: ShiftKey, ke: ShiftKey): boolean { return dari === "siang" && ke === "pagi"; }
export function susunRencanaHariPenuh(ops: any[], tanggal: string): RencanaRow[] {
  return (ops || [])
    .filter((o: any) => o && o.date === tanggal && o.status !== "batal")
    .sort(jamUrut)
    .map((o: any): RencanaRow => ({
      opId: String(o.id), tanggal: o.date || "", jam: o.time || "", operator: o.surgeon || "", anestesi: o.anesthesiologist || "",
      pasien: o.patient || "", umur: teksUmur(o), kamar: o.room || "", jenis: o.procedure || "", alkes: "", keterangan: "",
    }));
}

/* Kandidat lusa: dinas Siang, atau Shift Malam aktif */
export const KOSONG_OPERASI = "Tidak ada acara operasi";
export function pakaiKandidatLusa(shift: ShiftKey, malamAktif: boolean): boolean { return shift === "siang" || malamAktif; }
export function tglLusa(tanggal: string): string { return geserHari(tanggal, 2); }
export function susunKandidatLusa(ops: any[], tanggal: string): KandidatRow[] {
  const tgl = tglLusa(tanggal);
  return (ops || [])
    .filter((o: any) => o && o.date === tgl && o.status !== "batal")
    .sort(jamUrut)
    .map((o: any): KandidatRow => ({
      opId: String(o.id), tanggal: o.date || "", jam: o.time || "", dokter: o.surgeon || "", pasien: o.patient || "",
      umur: teksUmur(o), jenis: o.procedure || "", keterangan: "",
    }));
}
export function judulKandidat(l: LaporanHarian): string {
  return l.tanggalKandidat ? `KANDIDAT (OPERASI LUSA, ${tanggalPanjang(l.tanggalKandidat)})` : "KANDIDAT (OPERASI BESOK)";
}
