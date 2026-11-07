/* src/utils/mutu-ok.ts
   Data & Excel untuk tab "Pemantauan Kamar Bedah" (MUTU kamar bedah).
   Satu pasien/operasi = satu baris. Baris diturunkan otomatis dari daftar operasi (tab Jadwal),
   isian perawat disimpan di tabel kb_mutu_ok. Excel dibuat dari template mutu.xlsx. */
import ExcelJS from "exceljs";
import { MUTU_TEMPLATE_B64 } from "./mutu-template";

export const NAMA_BULAN = ["JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"];

export interface MutuRow {
  id: string;
  opId: string;
  tanggal: string;
  rm: string;
  f: Record<string, string>;
  updated_at?: string;
}
export interface MutuItem { op: any; row: MutuRow | null; tanggal: string; jam: string; }

export type TipeField = "pilih" | "teks" | "angka" | "jam";
export interface FieldDef { k: string; l: string; t: TipeField; }
export interface GrupForm { id: string; l: string; f: FieldDef[]; auto?: "anest" | "bedah"; }

const P = (k: string, l: string): FieldDef => ({ k, l, t: "pilih" });
export const GRUP_FORM: GrupForm[] = [
  { id: "kat", l: "KATEGORI", f: [P("kat_k", "K"), P("kat_s", "S"), P("kat_b", "B"), P("kat_kh", "KH")] },
  { id: "skp", l: "SKP - Kepatuhan Visite Dokter dan Pengisian Form", auto: "anest", f: [P("skp_ya", "YA"), P("skp_tdk", "TDK")] },
  { id: "bius", l: "PEMBIUSAN", f: [P("konv_ya", "Konversi YA"), P("konv_tdk", "Konversi TDK"), P("sat_ya", "Saturasi <92% YA"), P("sat_tdk", "Saturasi <92% TDK")] },
  { id: "to", l: "Pelaksanaan Time Out Verbal", f: [P("to_ya", "YA"), P("to_tdk", "TDK"), { k: "to_nama", l: "Nama", t: "teks" }] },
  { id: "nas", l: "NASIONAL - Penundaan Operasi Elektif >1 jam", f: [P("tunda_ya", "YA"), P("tunda_tdk", "TDK"), { k: "tunda_mnt", l: "Waktu (menit)", t: "angka" }] },
  { id: "bedah", l: "PEMBEDAHAN", auto: "bedah", f: [P("site_ya", "Site Marking YA"), P("site_tdk", "Site Marking TDK"), P("ases_ya", "Assesmen Prabedah YA"), P("ases_tdk", "Assesmen Prabedah TDK"), P("dx_ya", "Salah Diagnosa YA"), P("dx_tdk", "Salah Diagnosa TDK")] },
  { id: "ponek", l: "PONEK - Waktu Tanggap SC < 30 menit", f: [P("sc_ya", "YA"), P("sc_tdk", "TDK"), { k: "sc_daftar", l: "Waktu Daftar", t: "jam" }, { k: "sc_insisi", l: "Waktu Insisi", t: "jam" }] },
  { id: "ppra", l: "PPRA - Operasi Bersih tanpa AB profilaksis pada operasi FAM", f: [P("ppra_ya", "YA"), P("ppra_tdk", "TDK")] },
  { id: "unit", l: "UNIT - Kepatuhan pengisian surgical safety checklist", f: [P("unit_ya", "YA"), P("unit_tdk", "TDK")] },
];

export const adaIsi = (row: MutuRow | null, g: GrupForm): boolean =>
  !!row && g.f.some((fd) => String(row.f?.[fd.k] ?? "").trim() !== "");
export const adaIsiApapun = (row: MutuRow | null): boolean =>
  !!row && (String(row.rm || "").trim() !== "" || Object.values(row.f || {}).some((v) => String(v ?? "").trim() !== ""));

/* ID baris mutu diturunkan dari id operasi supaya tidak pernah dobel antar perangkat. */
const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
export function idDariOp(opId: string): string {
  const s = String(opId);
  if (UUID_RE.test(s)) return s.toLowerCase();
  const h = [0x811c9dc5, 0x9e3779b9, 0x85ebca6b, 0xc2b2ae35];
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    h[0] = Math.imul(h[0] ^ c, 2654435761); h[1] = Math.imul(h[1] ^ c, 1597334677);
    h[2] = Math.imul(h[2] ^ c, 2246822519); h[3] = Math.imul(h[3] ^ c, 3266489917);
    h[0] ^= h[3] >>> 13; h[1] ^= h[0] >>> 11; h[2] ^= h[1] >>> 7; h[3] ^= h[2] >>> 17;
  }
  const hex = h.map((x) => (x >>> 0).toString(16).padStart(8, "0")).join("");
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-4${hex.slice(13, 16)}-a${hex.slice(17, 20)}-${hex.slice(20, 32)}`;
}

export function baruRow(op: any): MutuRow {
  return { id: idDariOp(op.id), opId: String(op.id), tanggal: String(op.date || ""), rm: "", f: {} };
}

const menit = (j: string): number => { const m = /^(\d{1,2})[:.](\d{2})/.exec(String(j || "")); return m ? +m[1] * 60 + +m[2] : 9999; };

/* Gabung daftar operasi (tab Jadwal) dengan isian mutu. Operasi batal tidak ditampilkan. */
export function gabung(ops: any[], list: MutuRow[]): MutuItem[] {
  const peta = new Map<string, MutuRow>();
  list.forEach((r) => { if (r.opId) peta.set(String(r.opId), r); });
  return (ops || [])
    .filter((o) => o && o.date && o.status !== "batal")
    .map((o) => ({ op: o, row: peta.get(String(o.id)) || null, tanggal: String(o.date), jam: String(o.time || "") }))
    .sort((a, b) => (a.tanggal < b.tanggal ? -1 : a.tanggal > b.tanggal ? 1 : menit(a.jam) - menit(b.jam)));
}

export const rmItem = (it: MutuItem): string => String(it.op?.rm ?? "").trim() || String(it.row?.rm ?? "").trim();

const angkaAtauTeks = (v: string): string | number | null => {
  const s = String(v ?? "").trim();
  if (s === "") return null;
  return /^-?\d+(\.\d+)?$/.test(s) && !/^0\d/.test(s) ? Number(s) : s;
};
const pilihan = (row: MutuRow | null, k: string): number | null => {
  const v = row?.f?.[k];
  return v === "1" ? 1 : v === "0" ? 0 : null;
};
const teks = (row: MutuRow | null, k: string): string | null => {
  const v = String(row?.f?.[k] ?? "").trim();
  return v === "" ? null : v;
};

/* 35 kolom A..AI, urutan persis sama dengan mutu.xlsx */
export function nilaiBaris(it: MutuItem, no: number): (string | number | null)[] {
  const r = it.row;
  const rm = rmItem(it);
  return [
    no, parseInt(it.tanggal.slice(8, 10), 10),
    pilihan(r, "kat_k"), pilihan(r, "kat_s"), pilihan(r, "kat_b"), pilihan(r, "kat_kh"),
    rm === "" ? null : angkaAtauTeks(rm),
    pilihan(r, "skp_ya"), pilihan(r, "skp_tdk"), String(it.op?.anesthesiologist ?? "").trim() || null,
    pilihan(r, "konv_ya"), pilihan(r, "konv_tdk"), pilihan(r, "sat_ya"), pilihan(r, "sat_tdk"),
    pilihan(r, "to_ya"), pilihan(r, "to_tdk"), teks(r, "to_nama"),
    pilihan(r, "tunda_ya"), pilihan(r, "tunda_tdk"), angkaAtauTeks(r?.f?.tunda_mnt ?? ""),
    pilihan(r, "site_ya"), pilihan(r, "site_tdk"), pilihan(r, "ases_ya"), pilihan(r, "ases_tdk"),
    String(it.op?.surgeon ?? "").trim() || null,
    pilihan(r, "dx_ya"), pilihan(r, "dx_tdk"),
    pilihan(r, "sc_ya"), pilihan(r, "sc_tdk"), teks(r, "sc_daftar"), teks(r, "sc_insisi"),
    pilihan(r, "ppra_ya"), pilihan(r, "ppra_tdk"),
    pilihan(r, "unit_ya"), pilihan(r, "unit_tdk"),
  ];
}

const KOLOM_TOTAL = 35, BARIS_AWAL = 7, BARIS_CONTOH = 3;

/* Excel dibuat dari template mutu.xlsx: judul, gabungan sel, font, garis tepi tetap sama. */
export async function buatWorkbookMutu(items: MutuItem[], bulan: string): Promise<ExcelJS.Workbook> {
  const bin = atob(MUTU_TEMPLATE_B64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes.buffer as any);
  const ws = wb.worksheets[0];
  const [thn, bln] = bulan.split("-");
  ws.getCell("A2").value = `Bulan : ${NAMA_BULAN[Math.max(0, Math.min(11, parseInt(bln, 10) - 1))]} ${thn}`;

  const gaya: any[] = [];
  for (let c = 1; c <= KOLOM_TOTAL; c++) gaya[c] = JSON.parse(JSON.stringify(ws.getCell(BARIS_AWAL, c).style || {}));
  const tinggi = ws.getRow(BARIS_AWAL).height || 15.5;
  for (let r = BARIS_AWAL; r < BARIS_AWAL + BARIS_CONTOH; r++) for (let c = 1; c <= KOLOM_TOTAL; c++) ws.getCell(r, c).value = null;

  items.forEach((it, i) => {
    const r = BARIS_AWAL + i;
    const nilai = nilaiBaris(it, i + 1);
    for (let c = 1; c <= KOLOM_TOTAL; c++) {
      const cell = ws.getCell(r, c);
      cell.style = JSON.parse(JSON.stringify(gaya[c]));
      cell.value = nilai[c - 1];
    }
    ws.getRow(r).height = tinggi;
  });
  return wb;
}
