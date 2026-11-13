/* src/components/view-rekening.tsx
   Tab Rekening (grup Operasi): formulir "Laporan Tindakan Pembedahan Kamar Bedah RS. Panti Rini".
   Pasien dipilih dari daftar operasi di tab Jadwal (nama, RM, umur, operator, dll terisi otomatis).
   Isian tersimpan ke tabel baru kb_rekening_ok: 1 baris per pasien + 1 baris daftar obat/barang.
   Bisa Print (pilih ukuran kertas, hasil 1 lembar), unduh Word, dan unduh Excel.
   Komponen ini berdiri sendiri: tidak mengubah tabel atau data yang sudah ada. */
import { useState, useEffect, useMemo, useRef } from "react";
import type { CSSProperties } from "react";
import * as XLSX from "xlsx";
import { C } from "../constants";
import { maskName } from "../utils";
import { SUPA_CLIENT } from "../services/supabase";
import { escHtml } from "../utils/security";
import { downloadElementAsWord } from "../utils/word-export";
import { labelKatalog } from "../utils/implan-ok";
import { KATALOG_TULANG } from "../constants/kode-implant";
import type { KodeImplan } from "../constants/kode-implant";

const TABEL = "kb_rekening_ok";

/* ───────── tipe ───────── */
export interface Baris { id: string; n: string; v: string; u: string; t?: "yt" | "pilih"; o?: string[]; p?: string; x?: string; }
export interface ItemImp { id: string; nama: string; jumlah: string; }
export interface Rek {
  id: string; opId: string; tanggal: string; updated_at?: string;
  ov: Record<string, string>;
  pa: string; kamar: string;
  anest: Baris[]; tengah: Baris[]; lain: Baris[]; benang: Baris[]; implant: ItemImp[];
}
type Sek = "anest" | "tengah" | "lain";
interface Master { id: string; updated_at?: string; anest: Baris[]; tengah: Baris[]; lain: Baris[]; }

/* ───────── daftar bawaan (sesuai formulir) ───────── */
const bb = (id: string, n: string, u = "", extra: Partial<Baris> = {}): Baris => ({ id, n, v: "", u, ...extra });
export const DEF_ANEST: Baris[] = [
  bb("a1", "Recofol", "mg"), bb("a2", "Ketalar", "mg"), bb("a3", "Succinyl", "mg"), bb("a4", "Dormicum", "mg"), bb("a5", "Pethidin", "mg"),
  bb("a6", "Hallothane", "% / menit"), bb("a7", "Flouthane", "% / menit"), bb("a8", "Ethrane", "% / menit"), bb("a9", "Fentanyl", "cc"),
  bb("a10", "N2O", "liter / menit"), bb("a11", "O2 + jasa dan canul oksigen", "liter / menit"),
];
export const DEF_TENGAH: Baris[] = [
  bb("t1", "Dextran L"), bb("t2", "Ringer Dextrose 5%"), bb("t3", "Asering"), bb("t4", "KA-EN"), bb("t5", "Ringer Lactated"), bb("t6", "Haes"),
  bb("t7", "NaCl", "", { t: "pilih", o: ["1 liter", "500 cc", "25 cc"], p: "" }),
  bb("t8", "Set Infus"), bb("t9", "Set Tranfusi"), bb("t10", "Spuit 20 cc"), bb("t11", "Spuit 10 cc"), bb("t12", "Spuit 5 cc"), bb("t13", "Spuit 3 cc"), bb("t14", "Spuit 1 cc"),
  bb("t15", "Wing Needle"), bb("t16", "Aquadest"), bb("t17", "Gaas"), bb("t18", "Suntik"), bb("t19", "Handschoon"), bb("t20", "Plester"), bb("t21", "Masker"), bb("t22", "Gurita"),
  bb("t23", "Mess", "", { t: "pilih", o: ["11", "15", "10"], p: "", x: "No." }),
];
export const DEF_LAIN: Baris[] = [
  bb("l1", "Sewa boor", "", { t: "yt" }), bb("l2", "Surgical gown"), bb("l3", "Chest electrode"), bb("l4", "Mayo"), bb("l5", "Dettol pekat"),
  bb("l6", "Breathing system"), bb("l7", "Lina pen"), bb("l8", "PD plate"), bb("l9", "Bactigras"), bb("l10", "Ondansetron"),
  bb("l11", "Ketorolac"), bb("l12", "Paracetamol"), bb("l13", "Underpad"), bb("l14", "Braunoderm", "cc"), bb("l15", "Povidon", "cc"), bb("l16", "Perhidrol", "cc"),
];

/* ───────── ukuran kertas ───────── */
export type KunciKertas = "A3" | "A4" | "A5" | "F4" | "Letter" | "Legal";
interface Kertas { k: KunciKertas; l: string; w: number; h: number; }
export const LIST_KERTAS: Kertas[] = [
  { k: "A3", l: "A3", w: 297, h: 420 }, { k: "A4", l: "A4", w: 210, h: 297 }, { k: "A5", l: "A5", w: 148, h: 210 },
  { k: "F4", l: "F4 (Folio)", w: 215, h: 330 }, { k: "Letter", l: "Letter", w: 216, h: 279 }, { k: "Legal", l: "Legal", w: 216, h: 356 },
];
const cariKertas = (k: string): Kertas => LIST_KERTAS.find((x) => x.k === k) || LIST_KERTAS[1] || { k: "A4", l: "A4", w: 210, h: 297 };
/* skala supaya seluruh lembar muat di 1 halaman kertas terpilih (margin 6 mm) */
export const skalaMuat = (w: number, h: number, lebar: number, tinggi: number): number => {
  const mm = 96 / 25.4;
  const pw = (w - 12) * mm, ph = (h - 12) * mm;
  return Math.max(0.2, Math.min(pw / lebar, ph / tinggi) * 0.98);
};

/* ───────── util ───────── */
const norm = (s: string): string => s.toLowerCase().replace(/,/g, ".").replace(/\s+/g, " ").trim();
const idBaru = (): string => "n" + Math.random().toString(36).slice(2, 8) + Date.now().toString(36).slice(-3);
const hariIni = (): string => { const d = new Date(); return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); };
const fmtTgl = (s: string): string => (/^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(8, 10) + "-" + s.slice(5, 7) + "-" + s.slice(0, 4) : s);

function hash128(str: string): number[] {
  let h1 = 1779033703, h2 = 3144134277, h3 = 1013904242, h4 = 2773480762;
  for (let i = 0; i < str.length; i++) {
    const k = str.charCodeAt(i);
    h1 = h2 ^ Math.imul(h1 ^ k, 597399067);
    h2 = h3 ^ Math.imul(h2 ^ k, 2869860233);
    h3 = h4 ^ Math.imul(h3 ^ k, 951274213);
    h4 = h1 ^ Math.imul(h4 ^ k, 2716044179);
  }
  h1 = Math.imul(h3 ^ (h1 >>> 18), 597399067);
  h2 = Math.imul(h4 ^ (h2 >>> 22), 2869860233);
  h3 = Math.imul(h1 ^ (h3 >>> 17), 951274213);
  h4 = Math.imul(h2 ^ (h4 >>> 19), 2716044179);
  return [(h1 ^ h2 ^ h3 ^ h4) >>> 0, (h2 ^ h1) >>> 0, (h3 ^ h1) >>> 0, (h4 ^ h1) >>> 0];
}
/* UUID tetap (sama setiap kali) untuk satu id operasi, supaya satu pasien = satu baris rekening */
export const uuidDari = (s: string): string => {
  const h = hash128("rekening-" + s).map((n) => n.toString(16).padStart(8, "0")).join("");
  return h.slice(0, 8) + "-" + h.slice(8, 12) + "-4" + h.slice(13, 16) + "-" + "89ab".charAt(parseInt(h.charAt(16), 16) & 3) + h.slice(17, 20) + "-" + h.slice(20, 32);
};
const MASTER_ID = uuidDari("master-daftar");

const salin = (a: Baris[]): Baris[] => a.map((r) => ({ ...r, v: "", p: r.t === "pilih" ? "" : r.p }));
const rekBaru = (op: any, m: Master): Rek => ({
  id: uuidDari(String(op.id)), opId: String(op.id), tanggal: String(op.date || ""), ov: {}, pa: "", kamar: "",
  anest: salin(m.anest), tengah: salin(m.tengah), lain: salin(m.lain), benang: [], implant: [],
});
const masterBawaan = (): Master => ({ id: MASTER_ID, anest: DEF_ANEST, tengah: DEF_TENGAH, lain: DEF_LAIN });

const umurOp = (op: any): string => {
  const a = String(op.age || "").trim(), m = String(op.ageMonths || "").trim();
  if (a) return a + " th" + (m ? " " + m + " bln" : "");
  return m ? m + " bln" : "";
};
const dasarOp = (op: any): Record<string, string> => ({
  nama: String(op.patient || ""), rm: String(op.rm || ""), tanggal: String(op.date || ""), umur: umurOp(op),
  ruang: String(op.ruangAsal || ""), alamat: String(op.alamatPasien || ""), operator: String(op.surgeon || ""),
  anestesi: String(op.anesthesiologist || ""), asisten: String(op.assistantNurse || ""), diagnosa: String(op.diagnosis || ""),
  pembedahan: String(op.procedure || ""),
  rn: "", kelas: "", konsulent: "", jamA: "", jamB: "", postOp: "",
});
/* nilai akhir tiap kolom data pasien: isian manual (ov) menimpa data dari Jadwal */
const nilaiField = (rek: Rek, op: any): Record<string, string> => {
  const out = dasarOp(op);
  Object.keys(rek.ov).forEach((k) => { out[k] = String(rek.ov[k] ?? ""); });
  return out;
};

/* ───────── teks satu baris (dipakai Print, Word, Excel) ───────── */
const labelBaris = (r: Baris): string => r.n + (r.t === "pilih" && r.p ? " " + (r.x ? r.x + " " : "") + r.p : "");
const nilaiBaris = (r: Baris): string => (r.t === "yt" ? (r.v === "ya" ? "Ya" : r.v === "tidak" ? "Tidak" : "") : r.v);
const kePrint = (i: ItemImp): Baris => ({ id: i.id, n: i.nama, v: i.jumlah, u: "" });

/* ───────── HTML lembar (Print dan Word) ───────── */
const e = (s: string): string => escHtml(String(s || ""));
const TD = "font-family:Arial,Helvetica,sans-serif;font-size:12px;line-height:1.25;padding:2px 3px;vertical-align:bottom;color:#000;";
const TITIK = "border-bottom:1px dotted #000;";
const kotak = "border:1px solid #000;";

const kolomBaris = (rows: Baris[], kosong: number): string => {
  let h = "";
  rows.forEach((r) => {
    const v = nilaiBaris(r);
    h += "<tr><td style=\"" + TD + "white-space:nowrap\">" + e(labelBaris(r)) + "</td><td style=\"" + TD + "width:8px\">:</td><td style=\"" + TD + TITIK + "min-width:60px\">" + (v ? e(v) : "&nbsp;") + "</td><td style=\"" + TD + "white-space:nowrap\">" + e(r.u) + "</td></tr>";
  });
  for (let i = 0; i < kosong; i++) {
    h += "<tr><td style=\"" + TD + TITIK + "\">&nbsp;</td><td style=\"" + TD + "\">:</td><td style=\"" + TD + TITIK + "\">&nbsp;</td><td style=\"" + TD + "\">&nbsp;</td></tr>";
  }
  return "<table style=\"width:100%;border-collapse:collapse\">" + h + "</table>";
};
const barisKosongSaja = (n: number): string => {
  let h = "";
  for (let i = 0; i < n; i++) h += "<tr><td style=\"" + TD + TITIK + "\">&nbsp;</td></tr>";
  return "<table style=\"width:100%;border-collapse:collapse\">" + h + "</table>";
};
const judul = (t: string): string => "<div style=\"" + TD + "font-weight:bold;text-decoration:underline;padding-top:4px\">" + e(t) + "</div>";
const kolomTeks = (lab: string, v: string): string =>
  "<table style=\"width:100%;border-collapse:collapse\"><tr><td style=\"" + TD + "white-space:nowrap;padding-left:0\">" + e(lab) + " :</td><td style=\"" + TD + TITIK + "width:100%\">" + (v ? e(v) : "&nbsp;") + "</td></tr></table>";
const cek = (aktif: boolean): string => (aktif ? "[ X ]" : "[&nbsp;&nbsp;&nbsp;]");

/* isi lembar (tanpa pembungkus lebar): dipakai langsung oleh Print dan Word */
export const htmlLembar = (f: Record<string, string>, rek: Rek): string => {
  const s = (k: string): string => f[k] || "";
  const atas =
    "<table style=\"width:100%;border-collapse:collapse\"><tr><td style=\"" + TD + "text-align:center;font-size:17px;font-weight:bold;font-family:Georgia,'Times New Roman',serif\">LAPORAN TINDAKAN PEMBEDAHAN<br/>KAMAR BEDAH RS. PANTI RINI</td></tr></table>";
  const baris1 =
    "<table style=\"width:100%;border-collapse:collapse;margin-top:6px\"><tr>" +
    "<td style=\"" + TD + "width:34%\">" + kolomTeks("No. RM", s("rm")) + "</td>" +
    "<td style=\"" + TD + "width:33%\">" + kolomTeks("No. RN", s("rn")) + "</td>" +
    "<td style=\"" + TD + "width:33%\">" + kolomTeks("Tanggal", fmtTgl(s("tanggal"))) + "</td></tr></table>";
  const kotakPasien =
    "<table style=\"width:100%;border-collapse:collapse;" + kotak + "\">" +
    "<tr><td style=\"" + TD + "width:50%\">" + kolomTeks("NAMA", s("nama")) + "</td><td style=\"" + TD + "width:17%\">" + kolomTeks("Umur", s("umur")) + "</td><td style=\"" + TD + "width:33%\">" + kolomTeks("Ruang / Klas", s("ruang") + (s("kelas") ? " / " + s("kelas") : "")) + "</td></tr>" +
    "<tr><td colspan=\"3\" style=\"" + TD + "\">" + kolomTeks("Alamat", s("alamat")) + "</td></tr></table>";
  const kotakOperator =
    "<table style=\"width:100%;border-collapse:collapse;" + kotak + "border-top:0\"><tr>" +
    "<td style=\"" + TD + "width:38%;border-right:1px solid #000;vertical-align:top\">" +
    kolomTeks("Operator", s("operator")) + kolomTeks("Anaesthesi", s("anestesi")) + kolomTeks("Asisten", s("asisten")) + kolomTeks("Konsulent", s("konsulent")) + "</td>" +
    "<td style=\"" + TD + "width:34%;border-right:1px solid #000;text-align:center;vertical-align:top\">Lama Anaesthesi<br/>" +
    "<span style=\"display:inline-block;min-width:60px;" + TITIK + "\">" + (s("jamA") ? e(s("jamA")) : "&nbsp;") + "</span> s/d <span style=\"display:inline-block;min-width:60px;" + TITIK + "\">" + (s("jamB") ? e(s("jamB")) : "&nbsp;") + "</span> Jam</td>" +
    "<td style=\"" + TD + "width:28%;text-align:center;vertical-align:top\"><b>PA</b><br/>" + cek(rek.pa === "ya") + " Ya &nbsp;&nbsp; " + cek(rek.pa === "tidak") + " Tidak</td></tr></table>";
  const kotakDiagnosa =
    "<table style=\"width:100%;border-collapse:collapse;" + kotak + "border-top:0\">" +
    "<tr><td style=\"" + TD + "\">" + kolomTeks("Diagnosa pre op", s("diagnosa")) + "</td></tr>" +
    "<tr><td style=\"" + TD + "\">" + kolomTeks("Diagnosa post op", s("postOp")) + "</td></tr></table>";
  const kotakBedah =
    "<table style=\"width:100%;border-collapse:collapse;" + kotak + "border-top:0\">" +
    "<tr><td style=\"" + TD + "\">" + kolomTeks("Pembedahan", s("pembedahan")) + "</td></tr>" +
    "<tr><td style=\"" + TD + "\"><b>Kamar</b> : " + cek(rek.kamar === "Kecil") + " Kecil &nbsp; " + cek(rek.kamar === "Sedang") + " Sedang &nbsp; " + cek(rek.kamar === "Besar") + " Besar &nbsp; " + cek(rek.kamar === "Khusus") + " Khusus</td></tr></table>";
  const tiga =
    "<table style=\"width:100%;border-collapse:collapse;" + kotak + "border-top:0\"><tr>" +
    "<td style=\"width:36%;vertical-align:top;padding:2px 4px;border-right:1px solid #000\">" + judul("Obat-obat Anaesthesi :") + kolomBaris(rek.anest, 4) + judul("Lain - Lain :") + kolomBaris(rek.lain, 3) + "</td>" +
    "<td style=\"width:32%;vertical-align:top;padding:2px 4px;border-right:1px solid #000\">" + kolomBaris(rek.tengah, 4) + "</td>" +
    "<td style=\"width:32%;vertical-align:top;padding:2px 4px\">" + judul("Implant :") + kolomBaris(rek.implant.map(kePrint), 3) + judul("Benang :") + kolomBaris(rek.benang, 3) + judul("Lain lain :") + barisKosongSaja(6) + "</td></tr></table>";
  return "<div style=\"font-family:Arial,Helvetica,sans-serif;color:#000\">" + atas + baris1 + kotakPasien + kotakOperator + kotakDiagnosa + kotakBedah + tiga + "</div>";
};

/* ───────── gaya layar ───────── */
const KEL = "#374151";
const kartu: CSSProperties = { background: "#fff", border: "1px solid #e5e7eb", borderRadius: 14, padding: 14 };
const inp: CSSProperties = { padding: "8px 10px", borderRadius: 10, border: "1px solid #d1d5db", fontSize: 14, fontFamily: "inherit", background: "#fff", boxSizing: "border-box", minWidth: 0 };
const tbl = (warna: string, isi = false): CSSProperties => ({ padding: "8px 12px", borderRadius: 10, border: isi ? "none" : "1px solid " + warna, background: isi ? warna : "#fff", color: isi ? "#fff" : warna, fontWeight: 700, fontSize: 13, cursor: "pointer", fontFamily: "inherit" });
const label: CSSProperties = { fontSize: 12, fontWeight: 700, color: KEL, marginBottom: 3 };

/* satu bagian daftar baris (obat, cairan/alat, barang pengganti Lain-lain, benang) */
function Seksi(props: { judul: string; rows: Baris[]; ubah: (r: Baris[]) => void; master?: Baris[]; simpan?: (r: Baris) => void; simpanPilihan?: (r: Baris) => void }) {
  const { rows, ubah, master, simpan, simpanPilihan } = props;
  const set = (i: number, patch: Partial<Baris>) => ubah(rows.map((r, k) => (k === i ? { ...r, ...patch } : r)));
  const dmId = "dm-" + props.judul.replace(/[^a-zA-Z0-9]/g, "");
  /* nama cocok dengan item master: ganti baris dengan item itu (jumlah yang sudah diisi tetap) */
  const gantiNama = (i: number, nama: string) => {
    const m = master ? master.find((x) => norm(x.n) === norm(nama)) : undefined;
    const lama = rows[i];
    if (m && lama && norm(lama.n) !== norm(nama)) set(i, { n: m.n, u: m.u, t: m.t, o: m.o, p: m.t === "pilih" ? "" : m.p, x: m.x });
    else set(i, { n: nama });
  };
  return (
    <div style={{ marginTop: 14 }}>
      <div style={{ fontWeight: 800, color: KEL, fontSize: 14, marginBottom: 8 }}>{props.judul}</div>
      {master && <datalist id={dmId}>{master.map((m) => <option key={m.id} value={m.n} />)}</datalist>}
      {rows.map((r, i) => {
        const belumAda = !!master && r.n.trim() !== "" && !master.some((m) => norm(m.n) === norm(r.n));
        const pilihanBaru = r.t === "pilih" && !!r.p && !(r.o || []).includes(r.p);
        return (
          <div key={r.id} style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 6, flexWrap: "wrap" }}>
            <input value={r.n} list={dmId} placeholder="Nama" autoComplete="off" onChange={(ev) => gantiNama(i, ev.target.value)} style={{ ...inp, flex: "2 1 150px" }} />
            {master && master.length > 0 && (
              <select value="" aria-label="Ganti dari daftar" title="Ganti dengan item dari daftar" onChange={(ev) => { if (ev.target.value) gantiNama(i, ev.target.value); }} style={{ ...inp, flex: "0 0 46px", padding: "6px 4px" }}>
                <option value="">▾</option>
                {master.map((m) => <option key={m.id} value={m.n}>{m.n}</option>)}
              </select>
            )}
            {r.t === "yt" ? (
              <select value={r.v} onChange={(ev) => set(i, { v: ev.target.value })} style={{ ...inp, flex: "1 1 90px" }}>
                <option value="">-</option><option value="ya">Ya</option><option value="tidak">Tidak</option>
              </select>
            ) : r.t === "pilih" ? (
              <>
                <input list={"dl-" + r.id} value={r.p || ""} placeholder="Pilih / ketik" onChange={(ev) => set(i, { p: ev.target.value })} style={{ ...inp, flex: "1 1 100px" }} />
                <datalist id={"dl-" + r.id}>{(r.o || []).map((o) => <option key={o} value={o} />)}</datalist>
                <input value={r.v} placeholder="Jumlah" onChange={(ev) => set(i, { v: ev.target.value })} style={{ ...inp, flex: "1 1 70px" }} />
              </>
            ) : (
              <input value={r.v} placeholder="Isi" onChange={(ev) => set(i, { v: ev.target.value })} style={{ ...inp, flex: "1 1 90px" }} />
            )}
            <input value={r.u} placeholder="satuan" onChange={(ev) => set(i, { u: ev.target.value })} style={{ ...inp, flex: "0 1 90px" }} />
            {pilihanBaru && simpanPilihan && (
              <button type="button" style={tbl("#4f46e5")} onClick={() => { const baru = { ...r, o: [...(r.o || []), r.p || ""] }; set(i, { o: baru.o }); simpanPilihan(baru); }}>+ pilihan</button>
            )}
            {belumAda && simpan && <button type="button" style={tbl("#059669")} onClick={() => simpan(r)}>Simpan ke master</button>}
            <button type="button" aria-label={"Hapus " + (r.n || "baris")} title="Hapus dari pasien ini" style={tbl("#dc2626")} onClick={() => ubah(rows.filter((_, k) => k !== i))}>✕</button>
          </div>
        );
      })}
      <button type="button" style={tbl("#4f46e5")} onClick={() => ubah([...rows, { id: idBaru(), n: "", v: "", u: "" }])}>+ Tambah baris</button>
    </div>
  );
}

/* implant: ketik sebagian nama/kode dari katalog tulang, pilih dari daftar, atau tulis sendiri */
function TambahImplant({ onTambah }: { onTambah: (nama: string) => void }) {
  const [q, setQ] = useState("");
  const [fokus, setFokus] = useState(false);
  const daftar = useMemo(() => KATALOG_TULANG.map((k: KodeImplan) => { const lb = labelKatalog(k); return { lb, kode: k.k, cari: norm(lb + " " + k.k) }; }), []);
  const hasil = useMemo(() => {
    const t = norm(q).split(" ").filter(Boolean);
    if (!t.length) return [];
    return daftar.filter((d) => t.every((x) => d.cari.includes(x))).slice(0, 8);
  }, [q, daftar]);
  const pilih = (nama: string) => { if (nama.trim()) { onTambah(nama.trim()); setQ(""); } };
  return (
    <div style={{ position: "relative" }}>
      <input value={q} placeholder="Ketik nama atau kode implant" autoComplete="off" spellCheck={false}
        onChange={(ev) => setQ(ev.target.value)} onFocus={() => setFokus(true)} onBlur={() => setTimeout(() => setFokus(false), 150)}
        onKeyDown={(ev) => { if (ev.key === "Enter") { ev.preventDefault(); const h = hasil[0]; pilih(h ? h.lb : q); } if (ev.key === "Escape") setQ(""); }}
        style={{ ...inp, width: "100%" }} />
      {fokus && q.trim() !== "" && (
        <div style={{ position: "absolute", left: 0, right: 0, top: "100%", zIndex: 20, marginTop: 4, background: "#fff", border: "1px solid #d1d5db", borderRadius: 10, boxShadow: "0 8px 24px rgba(0,0,0,0.12)", maxHeight: 260, overflowY: "auto" }}>
          {hasil.map((d, i) => (
            <div key={i} onMouseDown={(ev) => { ev.preventDefault(); pilih(d.lb); }} style={{ padding: "8px 10px", cursor: "pointer", borderBottom: "1px solid #f3f4f6", display: "flex", justifyContent: "space-between", gap: 8, fontSize: 13 }}>
              <span style={{ fontWeight: 600, color: KEL }}>{d.lb}</span>
              <span style={{ color: "#9ca3af", fontSize: 11, whiteSpace: "nowrap" }}>{d.kode}</span>
            </div>
          ))}
          <div onMouseDown={(ev) => { ev.preventDefault(); pilih(q); }} style={{ padding: "8px 10px", cursor: "pointer", fontSize: 13, color: "#4f46e5", fontWeight: 700 }}>+ Tambah tulisan sendiri: {q.trim()}</div>
        </div>
      )}
    </div>
  );
}

const FIELD_ATAS: Array<[string, string]> = [
  ["nama", "Nama"], ["rm", "No. RM"], ["rn", "No. RN"], ["umur", "Umur"], ["ruang", "Ruang"], ["kelas", "Klas"],
  ["alamat", "Alamat"], ["operator", "Operator"], ["anestesi", "Anaesthesi"], ["asisten", "Asisten"], ["konsulent", "Konsulent"],
  ["jamA", "Lama anaesthesi: dari jam"], ["jamB", "sampai jam"], ["diagnosa", "Diagnosa pre op"], ["postOp", "Diagnosa post op"], ["pembedahan", "Pembedahan (nama tindakan / prosedur)"],
];

export function ViewRekening({ ops, showToast, privacyMode, upsertOneToSupa, role }: any) {
  const admin = role === "admin";
  const [tanggal, setTanggal] = useState<string>(hariIni());
  const [semua, setSemua] = useState(false);
  const [q, setQ] = useState("");
  const [pilihId, setPilihId] = useState<string | null>(null);
  const [rek, setRek] = useState<Rek | null>(null);
  const [dirty, setDirty] = useState(false);
  const [muat, setMuat] = useState(false);
  const [simpanBusy, setSimpanBusy] = useState(false);
  const [galat, setGalat] = useState("");
  const [kertas, setKertas] = useState<KunciKertas>("A4");
  const [master, setMaster] = useState<Master>(masterBawaan());
  const masterRef = useRef<Master>(master);
  const token = useRef(0);

  const daftarOps: any[] = Array.isArray(ops) ? ops : [];
  const op: any = pilihId ? daftarOps.find((o: any) => String(o.id) === pilihId) || null : null;

  /* daftar obat/barang bersama (1 baris di tabel) */
  useEffect(() => {
    let batal = false;
    (async () => {
      try {
        const r: any = await SUPA_CLIENT.from(TABEL).select("data").eq("id", MASTER_ID);
        const d = r?.data?.[0]?.data;
        if (!batal && d && Array.isArray(d.anest) && Array.isArray(d.tengah) && Array.isArray(d.lain)) {
          const m: Master = { id: MASTER_ID, anest: d.anest, tengah: d.tengah, lain: d.lain };
          masterRef.current = m; setMaster(m);
        }
      } catch { /* tanpa koneksi: pakai daftar bawaan */ }
    })();
    return () => { batal = true; };
  }, []);

  /* muat rekening pasien terpilih */
  useEffect(() => {
    if (!op) { setRek(null); setDirty(false); setGalat(""); return; }
    const tok = ++token.current;
    const id = uuidDari(String(op.id));
    setMuat(true); setGalat("");
    (async () => {
      let data: any = null; let err = "";
      try {
        const r: any = await SUPA_CLIENT.from(TABEL).select("data").eq("id", id);
        if (r?.error) err = String(r.error.message || r.error.code || "gagal memuat");
        data = r?.data?.[0]?.data || null;
      } catch (x: any) { err = String(x?.message || "tidak ada koneksi"); }
      if (tok !== token.current) return;
      const dasar = rekBaru(op, masterRef.current);
      setRek(data ? {
        ...dasar, ...data,
        ov: data.ov && typeof data.ov === "object" ? data.ov : {},
        anest: Array.isArray(data.anest) ? data.anest : dasar.anest, tengah: Array.isArray(data.tengah) ? data.tengah : dasar.tengah,
        lain: Array.isArray(data.lain) ? data.lain : dasar.lain, benang: Array.isArray(data.benang) ? data.benang : [], implant: Array.isArray(data.implant) ? data.implant : [],
      } : dasar);
      setDirty(false); setMuat(false);
      if (err) setGalat("Rekening lama belum bisa dimuat (" + err + "). Isian baru tetap bisa dibuat, tetapi simpan baru berhasil setelah tabel kb_rekening_ok tersedia.");
    })();
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [pilihId]);

  const daftar = useMemo(() => {
    const t = q.trim().toLowerCase();
    return daftarOps
      .filter((o: any) => o && o.status !== "batal" && (semua || String(o.date || "") === tanggal) && (!t || String(o.patient || "").toLowerCase().includes(t) || String(o.rm || "").toLowerCase().includes(t)))
      .sort((a: any, b: any) => String(b.date || "").localeCompare(String(a.date || "")) || String(a.time || "").localeCompare(String(b.time || "")))
      .slice(0, 80);
  }, [daftarOps, tanggal, semua, q]);

  const pilihPasien = (id: string) => {
    if (id === pilihId) return;
    if (dirty && !window.confirm("Ada isian yang belum disimpan. Pindah pasien dan buang perubahan?")) return;
    setPilihId(id);
  };
  const ubahRek = (patch: Partial<Rek>) => { setRek((r) => (r ? { ...r, ...patch } : r)); setDirty(true); };
  const setOv = (k: string, v: string) => { setRek((r) => (r ? { ...r, ov: { ...r.ov, [k]: v } } : r)); setDirty(true); };

  const simpanRek = async () => {
    if (!rek) return;
    setSimpanBusy(true);
    const baris: Rek = { ...rek, updated_at: new Date().toISOString() };
    const res = await upsertOneToSupa(TABEL, baris);
    setSimpanBusy(false);
    if (res?.ok) { setRek(baris); setDirty(false); showToast("Rekening tersimpan"); }
    else showToast("Gagal menyimpan: " + (res?.error || "error") + ". Isian masih di layar ini.", C.d);
  };

  const simpanMaster = async (m: Master, pesanOk: string) => {
    masterRef.current = m; setMaster(m);
    const res = await upsertOneToSupa(TABEL, { ...m, updated_at: new Date().toISOString() });
    if (res?.ok) showToast(pesanOk);
    else showToast("Gagal menyimpan daftar: " + (res?.error || "error"), C.d);
  };
  const simpanKeDaftar = (sec: Sek, r: Baris) => {
    if (!admin) return;
    const m = masterRef.current;
    if (m[sec].some((x) => norm(x.n) === norm(r.n))) return;
    const baru: Master = { ...m, [sec]: [...m[sec], { ...r, v: "", p: r.t === "pilih" ? "" : r.p, id: idBaru() }] };
    void simpanMaster(baru, "Ditambahkan ke daftar");
  };
  const simpanPilihan = (sec: Sek, r: Baris) => {
    if (!admin) return;
    const m = masterRef.current;
    const baru: Master = { ...m, [sec]: m[sec].map((x) => (norm(x.n) === norm(r.n) ? { ...x, o: Array.from(new Set([...(x.o || []), ...(r.o || [])])) } : x)) };
    void simpanMaster(baru, "Pilihan ditambahkan");
  };

  /* ───── Print, Word, Excel ───── */
  /* admin: susunan obat, cairan/alat, dan lain-lain di formulir ini jadi master untuk pasien baru */
  const simpanFormulirJadiMaster = () => {
    if (!admin || !rek) return;
    if (!window.confirm("Jadikan susunan obat, cairan/alat, dan lain-lain pada formulir ini sebagai master? Pasien baru akan memakai susunan ini. Rekening pasien yang sudah tersimpan tidak berubah.")) return;
    const isi = (a: Baris[]): Baris[] => salin(a.filter((r) => r.n.trim() !== ""));
    void simpanMaster({ ...masterRef.current, anest: isi(rek.anest), tengah: isi(rek.tengah), lain: isi(rek.lain) }, "Master formulir disimpan");
  };

  const cetak = () => {
    if (!rek || !op) return;
    const k = cariKertas(kertas);
    const inner = htmlLembar(nilaiField(rek, op), rek);
    const LEBAR = 760;
    const ukur = document.createElement("div");
    ukur.style.cssText = "position:fixed;left:-9999px;top:0;width:" + LEBAR + "px";
    ukur.innerHTML = inner;
    document.body.appendChild(ukur);
    const tinggi = Math.max(1, ukur.offsetHeight);
    ukur.remove();
    const s = skalaMuat(k.w, k.h, LEBAR, tinggi);
    const ifr = document.createElement("iframe");
    ifr.style.cssText = "position:fixed;right:0;bottom:0;width:0;height:0;border:0";
    document.body.appendChild(ifr);
    const w = ifr.contentWindow;
    if (!w) { ifr.remove(); return; }
    w.document.open();
    w.document.write("<!doctype html><html><head><meta charset='utf-8'><title>Rekening</title><style>@page{size:" + k.w + "mm " + k.h + "mm;margin:6mm}html,body{margin:0;padding:0}.pb{width:" + (Math.ceil(LEBAR * s) + 3) + "px;height:" + (Math.ceil(tinggi * s) + 3) + "px;overflow:hidden}.lb{width:" + LEBAR + "px;transform:scale(" + s + ");transform-origin:0 0}</style></head><body><div class='pb'><div class='lb'>" + inner + "</div></div></body></html>");
    w.document.close();
    setTimeout(() => { w.focus(); w.print(); setTimeout(() => ifr.remove(), 3000); }, 400);
  };
  const namaFile = (ext: string): string => "Rekening_" + (op ? String(op.patient || "pasien").replace(/[^A-Za-z0-9]+/g, "_") : "pasien") + "_" + (rek ? rek.tanggal : "") + "." + ext;
  const unduhWord = () => {
    if (!rek || !op) return;
    const el = document.createElement("div");
    el.innerHTML = "<div style=\"width:760px\">" + htmlLembar(nilaiField(rek, op), rek) + "</div>";
    downloadElementAsWord(el, namaFile("doc"));
  };
  const unduhExcel = () => {
    if (!rek || !op) return;
    const f = nilaiField(rek, op);
    const rows: string[][] = [
      ["LAPORAN TINDAKAN PEMBEDAHAN KAMAR BEDAH RS. PANTI RINI"], [],
      ["No. RM", f.rm || "", "No. RN", f.rn || "", "Tanggal", fmtTgl(f.tanggal || "")],
      ["Nama", f.nama || "", "Umur", f.umur || "", "Ruang / Klas", (f.ruang || "") + (f.kelas ? " / " + f.kelas : "")],
      ["Alamat", f.alamat || ""], ["Operator", f.operator || ""], ["Anaesthesi", f.anestesi || ""], ["Asisten", f.asisten || ""], ["Konsulent", f.konsulent || ""],
      ["Lama anaesthesi", (f.jamA || "") + " s/d " + (f.jamB || "") + " Jam"], ["PA", rek.pa === "ya" ? "Ya" : rek.pa === "tidak" ? "Tidak" : ""],
      ["Diagnosa pre op", f.diagnosa || ""], ["Diagnosa post op", f.postOp || ""], ["Pembedahan", f.pembedahan || ""], ["Kamar", rek.kamar], [],
      ["Bagian", "Nama", "Jumlah", "Satuan"],
    ];
    const tambah = (bag: string, a: Baris[]) => a.forEach((r) => rows.push([bag, labelBaris(r), nilaiBaris(r), r.u]));
    tambah("Obat anaesthesi", rek.anest); tambah("Cairan dan alat", rek.tengah); tambah("Lain-lain", rek.lain);
    rek.implant.forEach((i) => rows.push(["Implant", i.nama, i.jumlah, ""]));
    tambah("Benang", rek.benang);
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{ wch: 22 }, { wch: 44 }, { wch: 18 }, { wch: 18 }, { wch: 14 }, { wch: 18 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Rekening");
    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href = url; a.download = namaFile("xlsx");
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  };

  const f = rek && op ? nilaiField(rek, op) : null;
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
      <div style={kartu}>
        <div style={{ fontWeight: 800, fontSize: 17, color: KEL, marginBottom: 10 }}>Rekening Laporan Tindakan Pembedahan</div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
          <input type="date" value={tanggal} disabled={semua} onChange={(ev) => setTanggal(ev.target.value)} style={inp} />
          <label style={{ display: "flex", gap: 6, alignItems: "center", fontSize: 13, color: KEL }}>
            <input type="checkbox" checked={semua} onChange={(ev) => setSemua(ev.target.checked)} /> Semua tanggal
          </label>
          <input value={q} placeholder="Cari nama atau No. RM" onChange={(ev) => setQ(ev.target.value)} style={{ ...inp, flex: "1 1 180px" }} />
        </div>
        <div style={{ marginTop: 10, maxHeight: 260, overflowY: "auto", display: "flex", flexDirection: "column", gap: 6 }}>
          {daftar.length === 0 && <div style={{ color: "#6b7280", fontSize: 13 }}>Tidak ada pasien pada pilihan ini. Pasien harus sudah terdaftar di tab Daftar.</div>}
          {daftar.map((o: any) => (
            <div key={String(o.id)} onClick={() => pilihPasien(String(o.id))}
              style={{ padding: "8px 10px", borderRadius: 10, cursor: "pointer", border: "1px solid " + (String(o.id) === pilihId ? "#4f46e5" : "#e5e7eb"), background: String(o.id) === pilihId ? "#eef2ff" : "#fff" }}>
              <div style={{ fontWeight: 700, fontSize: 14, color: KEL }}>{privacyMode ? maskName(String(o.patient || ""), true) : String(o.patient || "")}</div>
              <div style={{ fontSize: 12, color: "#6b7280" }}>{fmtTgl(String(o.date || ""))} {String(o.time || "")} · RM {String(o.rm || "-")} · {String(o.procedure || "")}</div>
            </div>
          ))}
        </div>
      </div>

      {op && muat && <div style={{ ...kartu, color: "#6b7280" }}>Memuat rekening...</div>}
      {op && !muat && rek && f && (
        <div style={kartu}>
          {galat && <div style={{ marginBottom: 10, padding: "8px 10px", borderRadius: 10, background: "#fffbeb", border: "1px solid #fcd34d", fontSize: 13, color: "#92400e" }}>{galat}</div>}

          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center", marginBottom: 12 }}>
            <button type="button" disabled={simpanBusy || !dirty} style={{ ...tbl("#4f46e5", true), opacity: simpanBusy || !dirty ? 0.5 : 1 }} onClick={simpanRek}>{simpanBusy ? "Menyimpan..." : dirty ? "Simpan" : "Tersimpan"}</button>
            <select value={kertas} onChange={(ev) => setKertas(ev.target.value as KunciKertas)} style={inp} aria-label="Ukuran kertas">
              {LIST_KERTAS.map((k) => <option key={k.k} value={k.k}>{k.l}</option>)}
            </select>
            <button type="button" style={tbl("#4f46e5")} onClick={cetak}>Print</button>
            <button type="button" style={tbl("#2563eb")} onClick={unduhWord}>Word</button>
            <button type="button" style={tbl("#059669")} onClick={unduhExcel}>Excel</button>
            {admin && <button type="button" style={tbl("#b45309")} title="Jadikan susunan formulir ini master untuk pasien baru" onClick={simpanFormulirJadiMaster}>Simpan sebagai master formulir</button>}
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(200px,1fr))", gap: 10 }}>
            {FIELD_ATAS.map(([k, lb]) => (
              <div key={k} style={k === "alamat" || k === "diagnosa" || k === "postOp" || k === "pembedahan" ? { gridColumn: "1 / -1" } : undefined}>
                <div style={label}>{lb}</div>
                <input value={f[k] || ""} onChange={(ev) => setOv(k, ev.target.value)} style={{ ...inp, width: "100%" }} />
              </div>
            ))}
            <div>
              <div style={label}>PA</div>
              <select value={rek.pa} onChange={(ev) => ubahRek({ pa: ev.target.value })} style={{ ...inp, width: "100%" }}>
                <option value="">-</option><option value="ya">Ya</option><option value="tidak">Tidak</option>
              </select>
            </div>
            <div>
              <div style={label}>Kamar</div>
              <select value={rek.kamar} onChange={(ev) => ubahRek({ kamar: ev.target.value })} style={{ ...inp, width: "100%" }}>
                <option value="">-</option><option value="Kecil">Kecil</option><option value="Sedang">Sedang</option><option value="Besar">Besar</option><option value="Khusus">Khusus</option>
              </select>
            </div>
          </div>

          <Seksi judul="Obat-obat Anaesthesi" rows={rek.anest} ubah={(r) => ubahRek({ anest: r })} master={master.anest} simpan={admin ? (r) => simpanKeDaftar("anest", r) : undefined} />
          <Seksi judul="Cairan, set, dan alat" rows={rek.tengah} ubah={(r) => ubahRek({ tengah: r })} master={master.tengah} simpan={admin ? (r) => simpanKeDaftar("tengah", r) : undefined} simpanPilihan={admin ? (r) => simpanPilihan("tengah", r) : undefined} />
          <Seksi judul="Lain-lain" rows={rek.lain} ubah={(r) => ubahRek({ lain: r })} master={master.lain} simpan={admin ? (r) => simpanKeDaftar("lain", r) : undefined} />

          <div style={{ marginTop: 14 }}>
            <div style={{ fontWeight: 800, color: KEL, fontSize: 14, marginBottom: 8 }}>Implant</div>
            {rek.implant.map((it, i) => (
              <div key={it.id} style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 6, flexWrap: "wrap" }}>
                <input value={it.nama} placeholder="Nama implant" onChange={(ev) => ubahRek({ implant: rek.implant.map((x, k) => (k === i ? { ...x, nama: ev.target.value } : x)) })} style={{ ...inp, flex: "3 1 180px" }} />
                <input value={it.jumlah} placeholder="Jumlah" onChange={(ev) => ubahRek({ implant: rek.implant.map((x, k) => (k === i ? { ...x, jumlah: ev.target.value } : x)) })} style={{ ...inp, flex: "0 1 90px" }} />
                <button type="button" aria-label="Hapus implant" style={tbl("#dc2626")} onClick={() => ubahRek({ implant: rek.implant.filter((_, k) => k !== i) })}>✕</button>
              </div>
            ))}
            <TambahImplant onTambah={(nama) => ubahRek({ implant: [...rek.implant, { id: idBaru(), nama, jumlah: "1" }] })} />
          </div>

          <Seksi judul="Benang (isi manual)" rows={rek.benang} ubah={(r) => ubahRek({ benang: r })} />
        </div>
      )}
    </div>
  );
}
