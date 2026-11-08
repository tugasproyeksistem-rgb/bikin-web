/* src/utils/implan-ok.ts
   Data & Excel untuk grup MUTU: tab Implant (Tulang; Mata dan Hernia memakai komponen yang sama).
   Pasien diambil otomatis dari daftar operasi (tab Daftar/Jadwal): implant = "ya" dan kelompokImplant sesuai tab.
   Isian perawat (register, tanggal lahir, implant yang dipakai, jumlah, keterangan) disimpan di tabel kb_implan_ok.
   Excel dibuat dari template form_tulang.xlsx. */
import ExcelJS from "exceljs";
import { IMPLAN_TEMPLATE_TULANG_B64 } from "./implan-template";
import { KATALOG_TULANG } from "../constants/kode-implant";
import type { KodeImplan } from "../constants/kode-implant";
import { NAMA_BULAN, idDariOp, rmItem } from "./mutu-ok";

export type Kelompok = "tulang" | "mata" | "hernia";
export interface ItemImplan { kode: string; nama: string; jumlah: number; }
export interface ImplanRow {
  id: string; opId: string; tanggal: string;
  register: string; tglLahir: string; noTelp: string; alamat: string;
  items: ItemImplan[]; ket: string; updated_at?: string;
}
export interface ImplanItem { op: any; row: ImplanRow | null; tanggal: string; jam: string; }

export interface KonfigImplan { judul: string; berkas: string; katalog: KodeImplan[]; templateB64: string; }
/* Mata dan Hernia ditambahkan di sini nanti (katalog + template Excel masing-masing). */
export const KONFIG_IMPLAN: Partial<Record<Kelompok, KonfigImplan>> = {
  tulang: { judul: "PENGGUNAAN IMPLANT TULANG", berkas: "ImplantTulang", katalog: KATALOG_TULANG, templateB64: IMPLAN_TEMPLATE_TULANG_B64 },
};

export const labelKatalog = (k: KodeImplan): string => (k.p ? `${k.n} no ${k.p}` : k.n);
export const idImplan = (opId: string): string => idDariOp("implan-" + String(opId));
export const baruImplan = (op: any): ImplanRow => ({ id: idImplan(op.id), opId: String(op.id), tanggal: String(op.date || ""), register: "", tglLahir: "", noTelp: "", alamat: "", items: [], ket: "" });

export const adaIsiImplan = (row: ImplanRow | null): boolean =>
  !!row && (row.items.length > 0 || [row.register, row.tglLahir, row.noTelp, row.alamat, row.ket].some((v) => String(v || "").trim() !== ""));

const menit = (j: string): number => { const m = /^(\d{1,2})[:.](\d{2})/.exec(String(j || "")); return m ? +m[1] * 60 + +m[2] : 9999; };

export function gabungImplan(ops: any[], list: ImplanRow[], kelompok: Kelompok): ImplanItem[] {
  const peta = new Map<string, ImplanRow>();
  list.forEach((r) => { if (r.opId) peta.set(String(r.opId), r); });
  return (ops || [])
    .filter((o) => o && o.date && o.status !== "batal" && o.implant === "ya" && o.kelompokImplant === kelompok)
    .map((o) => ({ op: o, row: peta.get(String(o.id)) || null, tanggal: String(o.date), jam: String(o.time || "") }))
    .sort((a, b) => (a.tanggal < b.tanggal ? -1 : a.tanggal > b.tanggal ? 1 : menit(a.jam) - menit(b.jam)));
}

/* nomor telepon & alamat: dari data pasien kalau sudah tersimpan, kalau belum dari isian manual perawat */
export const telpItem = (it: ImplanItem): string => String(it.op?.noHp ?? "").trim() || String(it.row?.noTelp ?? "").trim();
export const alamatItem = (it: ImplanItem): string => String(it.op?.alamatPasien ?? "").trim() || String(it.row?.alamat ?? "").trim();

export interface BarisModel {
  jenis: "pasien" | "lanjut" | "kosong";
  no?: number; tgl?: number; rmReg?: string; nama?: string; lahir?: string; telp?: string; alamat?: string;
  impl?: string; jml?: number | null; ket?: string;
}
/* satu pasien: baris pertama berisi identitas + implant pertama, baris berikutnya hanya implant, lalu satu baris kosong */
export function susunBaris(items: ImplanItem[]): BarisModel[] {
  const out: BarisModel[] = [];
  items.forEach((it, i) => {
    const r = it.row; const list = r?.items || [];
    const rmReg = [rmItem(it as any), String(r?.register || "").trim()].filter((x) => x !== "").join(" ");
    out.push({
      jenis: "pasien", no: i + 1, tgl: parseInt(it.tanggal.slice(8, 10), 10), rmReg, nama: String(it.op?.patient ?? "").trim(),
      lahir: String(r?.tglLahir || ""), telp: telpItem(it), alamat: alamatItem(it),
      impl: list[0] ? list[0].nama : "", jml: list[0] ? list[0].jumlah : null, ket: String(r?.ket || "").trim(),
    });
    list.slice(1).forEach((x) => out.push({ jenis: "lanjut", impl: x.nama, jml: x.jumlah }));
    out.push({ jenis: "kosong" });
  });
  return out;
}

const tglDariIso = (s: string): Date | null => {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ""));
  return m ? new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])) : null;
};
export const fmtLahir = (s: string): string => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || "")); return m ? `${m[3]}/${m[2]}/${m[1]}` : String(s || ""); };

/* Excel dari template: judul, lebar kolom, font, dan susunan baris sama dengan form_tulang.xlsx */
export async function buatWorkbookImplan(items: ImplanItem[], bulan: string, kelompok: Kelompok): Promise<ExcelJS.Workbook> {
  const cfg = KONFIG_IMPLAN[kelompok];
  if (!cfg) throw new Error("Template Excel untuk kelompok ini belum tersedia");
  const bin = atob(cfg.templateB64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes.buffer as any);
  const ws = wb.worksheets[0];
  const [thn, bln] = bulan.split("-");
  ws.getCell("F2").value = `BULAN ${NAMA_BULAN[Math.max(0, Math.min(11, parseInt(bln, 10) - 1))]} ${thn}`;

  const KOLOM = 10;
  const donor = (baris: number): any[] => { const a: any[] = []; for (let c = 1; c <= KOLOM; c++) a[c] = JSON.parse(JSON.stringify(ws.getCell(baris, c).style || {})); return a; };
  const gayaPasien = donor(5), gayaKosong = donor(6), gayaLanjut = donor(8);
  const sisa = ws.rowCount - 4;
  if (sisa > 0) ws.spliceRows(5, sisa);

  susunBaris(items).forEach((b, i) => {
    const r = 5 + i;
    const gaya = b.jenis === "pasien" ? gayaPasien : b.jenis === "lanjut" ? gayaLanjut : gayaKosong;
    for (let c = 1; c <= KOLOM; c++) ws.getCell(r, c).style = JSON.parse(JSON.stringify(gaya[c]));
    if (b.jenis === "pasien") {
      const lahirTgl = tglDariIso(b.lahir || "");
      const nilai: (string | number | Date | null)[] = [b.no ?? null, b.tgl ?? null, b.rmReg || null, b.nama || null, lahirTgl || (b.lahir || null), b.telp || null, b.alamat || null, b.impl || null, b.jml ?? null, b.ket || null];
      nilai.forEach((v, k) => { ws.getCell(r, k + 1).value = v as any; });
      [4, 7, 10].forEach((c) => { const cell = ws.getCell(r, c); cell.alignment = { ...(cell.alignment || {}), wrapText: true, vertical: "middle" }; });
      ws.getRow(r).height = (b.alamat || "").length > 38 || (b.nama || "").length > 22 || (b.ket || "").length > 28 ? 31 : 15.5;
    } else if (b.jenis === "lanjut") {
      ws.getCell(r, 8).value = (b.impl || null) as any;
      ws.getCell(r, 9).value = (b.jml ?? null) as any;
      ws.getRow(r).height = 15.5;
    } else ws.getRow(r).height = 15.5;
  });
  return wb;
}
