/* Excel Laporan Harian OK mengikuti preview: satu sheet per laporan, Times New Roman 14. */
import ExcelJS from "exceljs";
import { ALAT_BUTIR, SHIFT_LABEL, fmtTgl, tanggalPanjang, rencanaHariPenuh, urutLaporanTurun } from "./laporan-harian";
import type { LaporanHarian } from "./laporan-harian";

const NAMA = "Times New Roman";
const W = [14, 9, 24, 24, 24, 10, 14, 30, 20, 26];
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

function tambahSheet(wb: ExcelJS.Workbook, l: LaporanHarian, nama: string): void {
  const ws = wb.addWorksheet(nama, { pageSetup: { orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
  ws.columns = W.map((w) => ({ width: w }));
  judul(ws, "LAPORAN HARIAN KAMAR BEDAH", true);
  judul(ws, "RS Panti Rini");
  ws.addRow([]);
  kv(ws, "Tanggal", tanggalPanjang(l.tanggal));
  kv(ws, "Shift", SHIFT_LABEL[l.shift]);
  kv(ws, "Serah terima", `${SHIFT_LABEL[l.dari]} ke ${SHIFT_LABEL[l.ke]}`);
  kv(ws, "Jumlah operasi", l.jumlahOperasi);
  kv(ws, "Operasi darurat / cyto", l.adaCyto === "ya" ? "Ya" : "Tidak");
  kv(ws, "Total pasien", l.totalPasien);

  bagian(ws, "JADWAL OPERASI");
  tabel(ws, ["No", "Jam", "Nama Pasien", "Jenis Operasi", "Dokter Operator", "Dokter Anestesi", "Status", "Catatan"],
    l.jadwal.map((r, i) => [String(i + 1), r.jam, r.pasien, r.jenis, r.operator, r.anestesi, r.status, r.catatan]), 10);

  bagian(ws, `RENCANA OPERASI — ${rencanaHariPenuh(l.dari, l.ke) ? "Semua operasi besok" : SHIFT_LABEL[l.ke]}, ${tanggalPanjang(l.tanggalRencana)}`);
  tabel(ws, ["Tanggal", "Jam", "Dokter Operator", "Dokter Anestesi", "Nama", "Umur", "Kamar", "Jenis Operasi", "Alkes", "Keterangan"],
    l.rencana.map((r) => [fmtTgl(r.tanggal), r.jam, r.operator, r.anestesi, r.pasien, r.umur, r.kamar, r.jenis, r.alkes, r.keterangan]), 4);

  bagian(ws, "KANDIDAT (OPERASI BESOK)");
  tabel(ws, ["Tanggal", "Jam", "Dokter", "Nama", "Umur", "Jenis Operasi", "Keterangan"],
    l.kandidat.map((r) => [fmtTgl(r.tanggal), r.jam, r.dokter, r.pasien, r.umur, r.jenis, r.keterangan]), 3);

  bagian(ws, "KONDISI PERALATAN & KAMAR");
  ALAT_BUTIR.forEach((b) => kv(ws, b.l, kond(l.alat[b.k])));
  kv(ws, "Lain-lain", l.alatLain);
  kv(ws, "Catatan", l.alatCatatan);

  bagian(ws, "STOK & BAHAN");
  kv(ws, "Instrumen", l.instrumen);
  kv(ws, "Implan / Bahan Habis Pakai", l.implan);
  kv(ws, "Alkes yang dipakai", l.alkesDipakai);
  kv(ws, "Alkes persediaan", l.alkesPersediaan);
  kv(ws, "PA", l.pa);
  kv(ws, "Lain-lain", l.stokLain);
  kv(ws, "Catatan", l.stokCatatan);

  bagian(ws, "KEJADIAN / INSIDEN / KENDALA");
  kv(ws, "Ada kejadian?", l.kejadian === "ya" ? "Ya" : "Tidak");
  kv(ws, "Jika ada, jelaskan", l.kejadian === "ya" ? l.uraian : "-");

  ws.addRow([]);
  const r1 = ws.addRow([]);
  const r2 = ws.addRow([]);
  const r3 = ws.addRow([]);
  r2.height = 56;
  const blok: [string, string, number, number][] = [
    ["Petugas Shift Lama", l.lama, 1, 3], ["Petugas Shift Baru", l.baru, 4, 7], ["Mengetahui / Kepala Kamar Bedah", l.kepala, 8, 10],
  ];
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

export function buatWorkbookLaporanLengkap(list: LaporanHarian[]): ExcelJS.Workbook {
  const wb = new ExcelJS.Workbook();
  const urut = [...list].sort((a, b) => -urutLaporanTurun(a, b));
  if (!urut.length) {
    const ws = wb.addWorksheet("Laporan");
    ws.addRow(["Belum ada laporan harian"]).getCell(1).font = font(true);
    return wb;
  }
  const dipakai = new Set<string>();
  urut.forEach((l) => {
    const dasar = `${fmtTgl(l.tanggal)} ${SHIFT_LABEL[l.shift]}`;
    let nama = dasar.slice(0, 31);
    let n = 2;
    while (dipakai.has(nama.toLowerCase())) nama = `${dasar} (${n++})`.slice(0, 31);
    dipakai.add(nama.toLowerCase());
    tambahSheet(wb, l, nama);
  });
  return wb;
}
