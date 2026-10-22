/* src/utils/jadwal-ok.ts
   Logika tab "Jadwal Operasi dan Peserta Kamar Bedah": susun baris dari
   daftar operasi, ambil perawat jaga dari Generator (kb_roster_gen),
   dan buat Excel (Times New Roman 14) lewat ExcelJS. */

import ExcelJS from "exceljs";
import { fmtTgl, namaHari, teksUmur, menitDariJam } from "./laporan-harian";

export interface BarisOK {
  id: string; jam: string; bgsl: string; ok: string; dokter: string; pasien: string; umur: string;
  jenis: string; asisten1: string; asisten2: string; onloop: string; rr: string;
}
export interface PerawatJaga { pagi: string; pg: string; sore: string; sg: string; cuti: string; }
export interface JadwalOK {
  id: string; tanggal: string; hari: string; perawat: PerawatJaga; baris: BarisOK[]; dihapus?: string[]; operator: string; updated_at: string;
}

export function barisKosong(): BarisOK {
  return { id: "m_" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), jam: "", bgsl: "", ok: "", dokter: "", pasien: "", umur: "", jenis: "", asisten1: "", asisten2: "", onloop: "", rr: "" };
}

/* Baris awal dari jadwal operasi hari itu (tanpa yang batal, urut jam) */
export function susunBarisDariOps(ops: any[], tanggal: string): BarisOK[] {
  return (ops || [])
    .filter((o: any) => o && o.date === tanggal && o.status !== "batal")
    .sort((a: any, b: any) => (menitDariJam(a?.time) ?? 9999) - (menitDariJam(b?.time) ?? 9999))
    .map((o: any): BarisOK => ({
      id: "op_" + String(o.id), jam: o.time || "", bgsl: "", ok: o.room || "",
      dokter: [o.surgeon, o.anesthesiologist].filter(Boolean).join(" / "),
      pasien: o.patient || "", umur: teksUmur(o), jenis: o.procedure || "",
      asisten1: o.assistantNurse || "", asisten2: o.circulatingNurse || "",
      onloop: o.onloopNurse || "", rr: o.rrKatim || "",
    }));
}

/* Perawat jaga per tanggal dari baris Generator bulan itu.
   nurses = JSON daftar perawat; jadwal = kode per hari (hari dipisah ",", perawat dipisah "|"). */
export function perawatDariGenerator(gen: any, tanggal: string): PerawatJaga | null {
  if (!gen) return null;
  let nurses: any = gen.nurses;
  try { if (typeof nurses === "string") nurses = JSON.parse(nurses); } catch { return null; }
  if (!Array.isArray(nurses) || typeof gen.jadwal !== "string") return null;
  const jadwal: string[] = gen.jadwal.split("|");
  const hari = Number(tanggal.slice(8, 10));
  if (!hari) return null;
  const grup: Record<string, string[]> = { P: [], PG: [], S: [], SG: [], T: [] };
  nurses.forEach((n: any, i: number) => {
    const kode = String((jadwal[i] || "").split(",")[hari - 1] || "").trim();
    const nama = String(n?.name || "").trim();
    if (nama && Object.prototype.hasOwnProperty.call(grup, kode)) grup[kode].push(nama);
  });
  return { pagi: grup.P.join(", "), pg: grup.PG.join(", "), sore: grup.S.join(", "), sg: grup.SG.join(", "), cuti: grup.T.join(", ") };
}

export function filterJadwalOK(list: JadwalOK[], mode: "tanggal" | "bulan" | "tahun", nilai: string): JadwalOK[] {
  if (!nilai) return [];
  return list
    .filter((l) => (mode === "tanggal" ? l.tanggal === nilai : l.tanggal.startsWith(nilai)))
    .sort((a, b) => a.tanggal.localeCompare(b.tanggal));
}

/* ── Excel ── */
const FONT_NAMA = "Times New Roman";
const FONT_UKURAN = 14;
const sx = (v: string): string => (/^[=+\-@\t\r]/.test(v) ? "'" + v : v);
const TIPIS: Partial<ExcelJS.Borders> = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
const KOLOM_EXCEL: { l: string; w: number }[] = [
  { l: "TANGGAL", w: 13 }, { l: "JAM", w: 8 }, { l: "BGSL", w: 12 }, { l: "OK", w: 16 }, { l: "DOKTER/ANESTESI", w: 30 },
  { l: "NAMA PASIEN", w: 26 }, { l: "UMUR", w: 9 }, { l: "JENIS OPERASI", w: 28 }, { l: "ASISTEN I", w: 20 },
  { l: "ASISTEN II", w: 20 }, { l: "ONLOOP", w: 20 }, { l: "RR", w: 16 },
];

function barisJudul(ws: ExcelJS.Worksheet, teks: string, tebal: boolean, tengah: boolean, tinggiMin = 22): void {
  const r = ws.addRow([teks]);
  ws.mergeCells(r.number, 1, r.number, KOLOM_EXCEL.length);
  const c = r.getCell(1);
  c.font = { name: FONT_NAMA, size: FONT_UKURAN, bold: tebal };
  c.alignment = { horizontal: tengah ? "center" : "left", vertical: "middle", wrapText: true };
  const lebarTotal = KOLOM_EXCEL.reduce((a, k) => a + k.w, 0);
  r.height = Math.max(tinggiMin, 21 * Math.ceil(teks.length / (lebarTotal * 0.8)));
}

function tambahSheet(wb: ExcelJS.Workbook, l: JadwalOK): void {
  const ws = wb.addWorksheet(fmtTgl(l.tanggal), {
    pageSetup: { orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });
  ws.columns = KOLOM_EXCEL.map((k) => ({ width: k.w }));
  barisJudul(ws, "JADWAL OPERASI KAMAR BEDAH", true, true, 26);
  barisJudul(ws, `${namaHari(l.tanggal)}, ${fmtTgl(l.tanggal)}`, true, true);
  const p = l.perawat;
  barisJudul(ws, `PERAWAT JAGA PAGI : ${sx(p.pagi || "")}`, false, false);
  barisJudul(ws, `PG : ${sx(p.pg || "")}`, false, false);
  barisJudul(ws, `SORE : ${sx(p.sore || "")}`, false, false);
  barisJudul(ws, `SG : ${sx(p.sg || "")}`, false, false);
  barisJudul(ws, `CUTI (T) : ${sx(p.cuti || "")}`, false, false);
  ws.addRow([]);

  const head = ws.addRow(KOLOM_EXCEL.map((k) => k.l));
  head.height = 24;
  head.eachCell({ includeEmpty: true }, (c) => {
    c.font = { name: FONT_NAMA, size: FONT_UKURAN, bold: true };
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    c.border = TIPIS;
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEEEEEE" } };
  });

  const total = Math.max(10, l.baris.length);
  for (let i = 0; i < total; i++) {
    const b = l.baris[i];
    const nilai: string[] = b
      ? [fmtTgl(l.tanggal), b.jam, b.bgsl, b.ok, b.dokter, b.pasien, b.umur, b.jenis, b.asisten1, b.asisten2, b.onloop, b.rr].map((v) => sx(String(v ?? "")))
      : KOLOM_EXCEL.map(() => "");
    const r = ws.addRow(nilai);
    let baris = 1;
    nilai.forEach((v, j) => { baris = Math.max(baris, Math.ceil(v.length / (KOLOM_EXCEL[j].w * 0.75)) || 1); });
    r.height = 22 * baris;
    r.eachCell({ includeEmpty: true }, (c) => {
      c.font = { name: FONT_NAMA, size: FONT_UKURAN };
      c.alignment = { vertical: "middle", wrapText: true };
      c.border = TIPIS;
    });
  }
}

/* Satu sheet per tanggal */
export function buatWorkbookJadwalOK(list: JadwalOK[]): ExcelJS.Workbook {
  const wb = new ExcelJS.Workbook();
  list.forEach((l) => tambahSheet(wb, l));
  return wb;
}

export async function workbookKeBlob(wb: ExcelJS.Workbook): Promise<Blob> {
  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf as unknown as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

export async function workbookKeBase64(wb: ExcelJS.Workbook): Promise<string> {
  const buf = await wb.xlsx.writeBuffer();
  const u8 = new Uint8Array(buf as unknown as ArrayBuffer);
  const CH = 8192;
  let bin = "";
  for (let i = 0; i < u8.length; i += CH) bin += String.fromCharCode(...u8.subarray(i, i + CH));
  return btoa(bin);
}
