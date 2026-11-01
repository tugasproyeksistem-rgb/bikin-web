/* Excel Laporan Kepala Jaga (Versi 2): satu sheet per laporan, Times New Roman 14. */
import ExcelJS from "exceljs";
import { ALAT_BUTIR, SHIFT_LABEL, fmtTgl, tanggalPanjang, KOSONG_OPERASI } from "./laporan-harian";
import type { LaporanHarian } from "./laporan-harian";

const NAMA = "Times New Roman";
const W = [12, 9, 24, 26, 22, 22, 34];
const N = W.length;
const LEBAR_NILAI = W.slice(3).reduce((a, b) => a + b, 0);
const TIPIS: Partial<ExcelJS.Borders> = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
const ABU: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEEEEEE" } };
const sx = (v: string): string => (/^[=+\-@\t\r]/.test(v) ? "'" + v : v);
const t = (v: unknown): string => sx(String(v ?? ""));
const font = (bold = false, size = 14): Partial<ExcelJS.Font> => ({ name: NAMA, size, bold });
const kond = (v: string): string => (v === "baik" ? "Baik" : v === "tidak" ? "Tidak baik" : "-");

function tinggi(teks: string[], lebar: number[]): number {
  let baris = 1;
  teks.forEach((s, i) => {
    const per = Math.max(4, Math.floor((lebar[i] || 10) * 0.85));
    const n = s.split("\n").reduce((a, p) => a + Math.max(1, Math.ceil(p.length / per)), 0);
    baris = Math.max(baris, n);
  });
  return Math.max(22, baris * 19);
}

function judul(ws: ExcelJS.Worksheet, teks: string, besar = false): void {
  const r = ws.addRow([teks]);
  ws.mergeCells(r.number, 1, r.number, N);
  const c = r.getCell(1);
  c.font = font(true, besar ? 16 : 14);
  c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  r.height = 24;
}

function bagian(ws: ExcelJS.Worksheet, teks: string): void {
  ws.addRow([]);
  const r = ws.addRow([teks]);
  ws.mergeCells(r.number, 1, r.number, N);
  for (let i = 1; i <= N; i++) { r.getCell(i).border = TIPIS; r.getCell(i).fill = ABU; }
  const c = r.getCell(1);
  c.font = font(true);
  c.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
  r.height = 24;
}

function kv(ws: ExcelJS.Worksheet, label: string, nilai: unknown): void {
  const isi = t(nilai);
  const r = ws.addRow([label]);
  ws.mergeCells(r.number, 1, r.number, 3);
  ws.mergeCells(r.number, 4, r.number, N);
  r.getCell(4).value = isi;
  for (let i = 1; i <= N; i++) r.getCell(i).border = TIPIS;
  r.getCell(1).font = font(true);
  r.getCell(4).font = font();
  r.getCell(1).alignment = { vertical: "middle", wrapText: true };
  r.getCell(4).alignment = { vertical: "middle", wrapText: true };
  r.height = tinggi([isi], [LEBAR_NILAI]);
}

function tabel(ws: ExcelJS.Worksheet, head: string[], rows: unknown[][], minimal: number): void {
  const h = ws.addRow(head);
  h.height = 24;
  h.eachCell({ includeEmpty: true }, (c) => {
    c.font = font(true);
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    c.border = TIPIS;
    c.fill = ABU;
  });
  if (!rows.length) {
    const k = ws.addRow([KOSONG_OPERASI]);
    ws.mergeCells(k.number, 1, k.number, N);
    for (let i = 1; i <= N; i++) k.getCell(i).border = TIPIS;
    k.getCell(1).font = { name: NAMA, size: 14, italic: true };
    k.getCell(1).alignment = { horizontal: "center", vertical: "middle" };
    k.height = 24;
    return;
  }
  const total = Math.max(minimal, rows.length);
  for (let i = 0; i < total; i++) {
    const nilai: string[] = rows[i] ? rows[i].map(t) : head.map(() => "");
    const r = ws.addRow(nilai);
    r.height = tinggi(nilai, W);
    r.eachCell({ includeEmpty: true }, (c) => {
      c.font = font();
      c.alignment = { vertical: "middle", wrapText: true };
      c.border = TIPIS;
    });
  }
}

export interface InfoExcelKJ { cytoIds?: Set<string>; }

/* Tandai baris cyto: teks "OPERASI CYTO" huruf besar, tebal, merah di kolom yang diberikan. */
function tandaiCyto(ws: ExcelJS.Worksheet, barisAwal: number, flags: boolean[], kolom: number): void {
  flags.forEach((cyto, i) => {
    if (!cyto) return;
    const c = ws.getCell(barisAwal + i, kolom);
    const teks = String(c.value ?? "");
    c.value = { richText: [
      ...(teks ? [{ text: teks + "\n", font: font() }] : []),
      { text: "OPERASI CYTO", font: { name: NAMA, size: 14, bold: true, color: { argb: "FFDC2626" } } },
    ] };
    const r = ws.getRow(barisAwal + i);
    r.height = Math.max(Number(r.height) || 22, 22) + 19;
  });
}

function tambahSheet(wb: ExcelJS.Workbook, d: any, nama: string, info?: InfoExcelKJ): void {
  const L: LaporanHarian | null = d.v1 || null;
  const ada = (id: unknown) => !!(info && info.cytoIds && info.cytoIds.has(String(id)));
  const x = (v: unknown) => (L ? v : "-");
  const ws = wb.addWorksheet(nama, { pageSetup: { orientation: "portrait", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
  ws.columns = W.map((w) => ({ width: w }));
  judul(ws, "LAPORAN KEPALA JAGA KAMAR BEDAH", true);
  judul(ws, "RS Panti Rini");
  ws.addRow([]);
  kv(ws, "Tanggal", tanggalPanjang(d.tanggal));
  kv(ws, "Shift", d.shiftTeks);
  kv(ws, "Serah terima", L ? `${SHIFT_LABEL[L.dari]} ke ${SHIFT_LABEL[L.ke]}` : "-");
  kv(ws, "Jumlah operasi", d.jumlah);
  kv(ws, "Operasi darurat / cyto", d.adaCyto ? "Ya" : "Tidak");
  kv(ws, "Total pasien", d.jumlah);
  kv(ws, "Siaga anestesi", (d.anest || []).join(", ") || "-");
  kv(ws, "Siaga anestesi cyto", (d.cyto || []).join(", ") || "-");
  kv(ws, "Perawat siaga", (d.nurses || []).join(", ") || "-");
  kv(ws, "Pembawa HP", d.pembawaHP || "-");

  const kepala = ["No", "Jam", "Nama Pasien", "Jenis Operasi", "Dokter Operator", "Dokter Anestesi", "Perawat"];
  const barisOp = (rows: any[]) => rows.map((r: any, i: number) => [String(i + 1), r.jam, r.pasien, r.jenis, r.operator, r.anestesi, (r.perawat || []).join("\n")]);
  bagian(ws, "OPERASI PAGI");
  const awalP = ws.rowCount;
  tabel(ws, kepala, barisOp(d.pagi || []), 3);
  tandaiCyto(ws, awalP + 2, (d.pagi || []).map((r: any) => ada(r.opId)), 3);
  bagian(ws, "OPERASI SIANG");
  const awalS = ws.rowCount;
  tabel(ws, kepala, barisOp(d.siang || []), 3);
  tandaiCyto(ws, awalS + 2, (d.siang || []).map((r: any) => ada(r.opId)), 3);
  bagian(ws, "OPERASI BESOK");
  const awalB = ws.rowCount;
  tabel(ws, ["Tanggal", "Jam", "Dokter", "Nama", "Umur", "Jenis Operasi", "Keterangan"],
    (d.besok || []).map((r: any) => [fmtTgl(r.tanggal), r.jam, r.dokter, r.pasien, r.umur, r.jenis, r.keterangan]), 3);
  tandaiCyto(ws, awalB + 2, (d.besok || []).map((r: any) => ada(r.opId)), 7);

  bagian(ws, "KONDISI PERALATAN & KAMAR");
  ALAT_BUTIR.forEach((b) => kv(ws, b.l, L ? kond((L.alat || {})[b.k]) : "-"));
  kv(ws, "Lain-lain", x(L?.alatLain));
  kv(ws, "Catatan", x(L?.alatCatatan));

  bagian(ws, "STOK & BAHAN");
  kv(ws, "Instrumen", x(L?.instrumen));
  kv(ws, "Implan / Bahan Habis Pakai", x(L?.implan));
  kv(ws, "Alkes yang dipakai", x(L?.alkesDipakai));
  kv(ws, "Alkes persediaan", x(L?.alkesPersediaan));
  kv(ws, "PA", x(L?.pa));
  kv(ws, "Lain-lain", x(L?.stokLain));
  kv(ws, "Catatan", x(L?.stokCatatan));

  bagian(ws, "KEJADIAN / INSIDEN / KENDALA");
  kv(ws, "Ada kejadian?", L ? (L.kejadian === "ya" ? "Ya" : "Tidak") : "-");
  kv(ws, "Jika ada, jelaskan", L && L.kejadian === "ya" ? L.uraian : "-");

  ws.addRow([]);
  const r1 = ws.addRow([]);
  const r2 = ws.addRow([]);
  const r3 = ws.addRow([]);
  r2.height = 56;
  const blok: [string, string, number, number][] = [
    ["Petugas Shift Lama", L?.lama, 1, 3], ["Petugas Shift Baru", L?.baru, 4, 5], ["Mengetahui / Kepala Kamar Bedah", L?.kepala, 6, 7],
  ] as any;
  blok.forEach(([jdl, nm, c1, c2]) => {
    [r1, r2, r3].forEach((r) => ws.mergeCells(r.number, c1, r.number, c2));
    const a = r1.getCell(c1);
    a.value = jdl; a.font = font(true); a.alignment = { horizontal: "center", vertical: "middle" };
    const c = r3.getCell(c1);
    c.value = "Nama Jelas: " + (t(nm) || "........................");
    c.font = font(true); c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    c.border = { top: { style: "thin" } };
  });
}

export function buatWorkbookKepalaJaga(list: any[], info?: InfoExcelKJ): ExcelJS.Workbook {
  const wb = new ExcelJS.Workbook();
  list.forEach((d) => tambahSheet(wb, d, String(d.tanggal), info));
  return wb;
}
