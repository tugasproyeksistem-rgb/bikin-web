/* src/utils/jadwal-perawat.ts
   Baca file jadwal perawat 1 bulan (format ekspor Generator: kolom NAMA, lalu
   tanggal 1..31 berisi kode P/PG/S/SG/T/L/LG) dan ambil perawat jaga per tanggal. */

import * as XLSX from "xlsx";
import type { PerawatJaga } from "./jadwal-ok";

export interface PerawatBulan { nama: string; kode: string[]; } /* kode[hari-1] */
export interface JadwalPerawatBulan { id: string; tahun: number; bulan: number; perawat: PerawatBulan[]; updated_at: string; }
export interface HasilBaca { tahun: number | null; bulan: number | null; perawat: PerawatBulan[]; diabaikan: number; }

const BULAN_3 = ["jan", "feb", "mar", "apr", "mei", "jun", "jul", "agu", "sep", "okt", "nov", "des"];
const KODE_DIKENAL = new Set(["P", "PG", "S", "SG", "T", "L", "LG"]);

function hari(v: any): number | null {
  if (typeof v === "number" && Number.isInteger(v) && v >= 1 && v <= 31) return v;
  const m = String(v ?? "").match(/^\s*(\d{1,2})(?!\d)/);
  if (!m) return null;
  const n = Number(m[1]);
  return n >= 1 && n <= 31 ? n : null;
}

function parseBaris(rows: any[][]): HasilBaca {
  let tahun: number | null = null;
  let bulan: number | null = null;
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    for (const c of rows[i] || []) {
      const m = String(c ?? "").match(/BULAN\s*:?\s*([A-Za-z]+)\s+(\d{4})/i);
      if (m) {
        const idx = BULAN_3.indexOf(m[1].slice(0, 3).toLowerCase());
        if (idx >= 0) { bulan = idx + 1; tahun = Number(m[2]); }
      }
    }
  }
  let hi = -1;
  let kn = -1;
  for (let i = 0; i < Math.min(rows.length, 15) && hi < 0; i++) {
    const j = (rows[i] || []).findIndex((c: any) => String(c ?? "").trim().toUpperCase() === "NAMA");
    if (j >= 0) { hi = i; kn = j; }
  }
  if (hi < 0) return { tahun, bulan, perawat: [], diabaikan: 0 };

  const kolomHari = (r: any[]): Map<number, number> => {
    const m = new Map<number, number>();
    r.forEach((c: any, j: number) => {
      if (j <= kn) return;
      const d = hari(c);
      if (d !== null && !m.has(d)) m.set(d, j);
    });
    return m;
  };
  const a = kolomHari(rows[hi] || []);
  const b = kolomHari(rows[hi + 1] || []);
  const pakaiB = b.size > a.size;
  const peta = pakaiB ? b : a;
  const mulai = (pakaiB ? hi + 1 : hi) + 1;

  const perawat: PerawatBulan[] = [];
  let diabaikan = 0;
  for (let i = mulai; i < rows.length; i++) {
    const r = rows[i] || [];
    const nama = String(r[kn] ?? "").trim().replace(/\s+/g, " ");
    if (!nama) continue;
    const kode: string[] = Array(31).fill("");
    let dikenal = 0;
    peta.forEach((j, d) => {
      const k = String(r[j] ?? "").trim().toUpperCase();
      kode[d - 1] = k;
      if (KODE_DIKENAL.has(k)) dikenal++;
    });
    if (dikenal < 3) { diabaikan++; continue; }
    perawat.push({ nama, kode });
  }
  return { tahun, bulan, perawat, diabaikan };
}

export async function bacaJadwalPerawat(file: File): Promise<HasilBaca> {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(new Uint8Array(buf), { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true }) as any[][];
  return parseBaris(rows);
}

/* Perawat jaga per tanggal dari baris kb_jadwal_perawat */
export function perawatDariUpload(data: any, tanggal: string): PerawatJaga | null {
  const daftar = data?.perawat;
  if (!Array.isArray(daftar)) return null;
  const h = Number(tanggal.slice(8, 10));
  if (!h) return null;
  const grup: Record<string, string[]> = { P: [], PG: [], S: [], SG: [], T: [] };
  daftar.forEach((p: any) => {
    const k = String(p?.kode?.[h - 1] ?? "").trim().toUpperCase();
    const nama = String(p?.nama ?? "").trim();
    if (nama && Object.prototype.hasOwnProperty.call(grup, k)) grup[k].push(nama);
  });
  return { pagi: grup.P.join(", "), pg: grup.PG.join(", "), sore: grup.S.join(", "), sg: grup.SG.join(", "), cuti: grup.T.join(", ") };
}
