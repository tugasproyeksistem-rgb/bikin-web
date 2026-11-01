/* Tab Laporan Harian OK, pilihan "Laporan Kepala Jaga" (Versi 2).
   Lembar jadi satu hari penuh dari data operasi dan jadwal siaga. Tanpa formulir.
   Simpan = salinan tetap ke kb_laporan_kepala_jaga (id kj_YYYY-MM-DD).
   Bagian bawah diambil dari Laporan Harian OK tanggal yang sama. */
import { useState, useEffect, useMemo, useRef } from "react";
import type { CSSProperties, ReactNode } from "react";
import html2canvas from "html2canvas";
import { downloadElementAsWord, elementToWordBlob } from "../utils/word-export";
import { normalizePhone62 } from "../utils";
import { C, iS } from "../constants";
import { maskName } from "../utils";
import { SUPA_CLIENT, AmbilLogMedis } from "../services/supabase";
import { dropboxUploadBytes } from "../services/dropbox";
import { Card, SH, LF, Btn } from "./primitives";
import type { LaporanHarian } from "../utils/laporan-harian";
import {
  SHIFT_URUT, SHIFT_LABEL, ALAT_BUTIR, KOSONG_OPERASI,
  isoLokal, tanggalPanjang, fmtTgl, susunRencana, susunKandidat, buatPdfDariJpeg,
} from "../utils/laporan-harian";
import { buatWorkbookKepalaJaga } from "../utils/laporan-kepala-jaga-excel";
import { ViewLaporanHarian } from "./view-laporan-harian";

const TABEL = "kb_laporan_kepala_jaga";

const pecahNama = (s: any): string[] =>
  String(s || "").split(/[,;]/).map((x) => x.trim()).filter(Boolean);

const daftarDari = (list: any, tunggal: any): string[] =>
  Array.isArray(list) && list.filter(Boolean).length
    ? list.map((x: any) => String(x).trim()).filter(Boolean)
    : pecahNama(tunggal);

const kondTeks = (v: string) => (v === "baik" ? "Baik" : v === "tidak" ? "Tidak baik" : "-");

function buatTeks(d: any, mask: (s: string) => string): string {
  const L: LaporanHarian | null = d.v1;
  const x = (v: any) => (L ? String(v || "-") : "-");
  const out: string[] = [];
  out.push("LAPORAN KEPALA JAGA KAMAR BEDAH", "RS Panti Rini", "");
  out.push(`Tanggal: ${tanggalPanjang(d.tanggal)}`, `Shift: ${d.shiftTeks}`);
  out.push(`Serah terima: ${L ? `${SHIFT_LABEL[L.dari]} ke ${SHIFT_LABEL[L.ke]}` : "-"}`);
  out.push(`Jumlah operasi: ${d.jumlah}`, `Operasi darurat / cyto: ${d.adaCyto ? "Ya" : "Tidak"}`, `Total pasien: ${d.jumlah}`, "");
  out.push(`Siaga anestesi: ${(d.anest || []).join(", ") || "-"}`);
  out.push(`Siaga anestesi cyto: ${(d.cyto || []).join(", ") || "-"}`);
  out.push(`Perawat siaga: ${(d.nurses || []).join(", ") || "-"}`);
  out.push(`Pembawa HP: ${d.pembawaHP || "-"}`, "");
  const op = (judul: string, rows: any[]) => {
    out.push(judul);
    if (!rows.length) out.push(KOSONG_OPERASI);
    rows.forEach((r: any, i: number) => {
      out.push(`${i + 1}. ${r.jam} | ${mask(r.pasien)} | ${r.jenis} | Operator: ${r.operator || "-"} | Anestesi: ${r.anestesi || "-"}`);
      if ((r.perawat || []).length) out.push(`   Perawat: ${r.perawat.join("; ")}`);
    });
    out.push("");
  };
  op("OPERASI PAGI", d.pagi || []);
  op("OPERASI SIANG", d.siang || []);
  out.push("OPERASI BESOK");
  if (!(d.besok || []).length) out.push(KOSONG_OPERASI);
  (d.besok || []).forEach((r: any, i: number) =>
    out.push(`${i + 1}. ${fmtTgl(r.tanggal)} ${r.jam} | ${r.dokter || "-"} | ${mask(r.pasien)} | ${r.umur || "-"} | ${r.jenis} | ${r.keterangan || "-"}`));
  out.push("", "KONDISI PERALATAN & KAMAR");
  ALAT_BUTIR.forEach((b) => out.push(`${b.l}: ${L ? kondTeks((L.alat || {})[b.k]) : "-"}`));
  out.push(`Lain-lain: ${x(L?.alatLain)}`, `Catatan: ${x(L?.alatCatatan)}`, "", "STOK & BAHAN");
  out.push(`Instrumen: ${x(L?.instrumen)}`, `Implan / Bahan Habis Pakai: ${x(L?.implan)}`, `Alkes yang dipakai: ${x(L?.alkesDipakai)}`);
  out.push(`Alkes persediaan: ${x(L?.alkesPersediaan)}`, `PA: ${x(L?.pa)}`, `Lain-lain: ${x(L?.stokLain)}`, `Catatan: ${x(L?.stokCatatan)}`, "");
  out.push("KEJADIAN / INSIDEN / KENDALA", `Ada kejadian? ${L ? (L.kejadian === "ya" ? "Ya" : "Tidak") : "-"}`);
  out.push(`Uraian: ${L && L.kejadian === "ya" ? L.uraian || "-" : "-"}`, "");
  out.push(`Petugas Shift Lama: ${x(L?.lama)}`, `Petugas Shift Baru: ${x(L?.baru)}`, `Mengetahui / Kepala Kamar Bedah: ${x(L?.kepala)}`);
  return out.join("\n");
}

function Lembar2({ d, mask, cytoIds }: { d: any; mask: (s: string) => string; cytoIds?: Set<string> }) {
  const L: LaporanHarian | null = d.v1;
  const th: CSSProperties = { border: "1px solid #111", padding: "4px 5px", fontSize: 11, fontWeight: 700, textAlign: "center", background: "#EEF2F7" };
  const td: CSSProperties = { border: "1px solid #111", padding: "4px 5px", fontSize: 11, verticalAlign: "top", height: 20 };
  const judul = (t: string) => <div style={{ fontWeight: 800, fontSize: 15, margin: "14px 0 5px" }}>{t}</div>;
  const tabel = (head: string[], lebar: string[], rows: ReactNode[][], minRows: number) => {
    const isi: ReactNode[][] = [...rows];
    while (isi.length < minRows) isi.push(head.map(() => ""));
    return (
      <table style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead><tr>{head.map((h, i) => <th key={i} style={{ ...th, width: lebar[i] }}>{h}</th>)}</tr></thead>
        <tbody>
          {rows.length
            ? isi.map((r, i) => <tr key={i}>{r.map((c, j) => <td key={j} style={td}>{c}</td>)}</tr>)
            : <tr><td colSpan={head.length} style={{ ...td, textAlign: "center", fontStyle: "italic" }}>{KOSONG_OPERASI}</td></tr>}
        </tbody>
      </table>
    );
  };
  const baris = (a: string, b: string, w = "34%") => <tr><td style={{ ...td, width: w, fontWeight: 700 }}>{a}</td><td style={td}>{b}</td></tr>;
  const perawatSel = (list: string[]): ReactNode => (
    <div>{(list || []).map((t: string, k: number) => <div key={k}>{t}</div>)}</div>
  );
  const adaCyto = (opId: string) => !!cytoIds && cytoIds.has(String(opId));
  const merah: CSSProperties = { color: "#DC2626", fontWeight: 800, textTransform: "uppercase" };
  const pasienCyto = (opId: string, nama: string): ReactNode => adaCyto(opId)
    ? <div><div>{nama}</div><div style={merah}>OPERASI CYTO</div></div>
    : nama;
  const ketCyto = (opId: string, teks: string): ReactNode => adaCyto(opId)
    ? <div>{teks && <div>{teks}</div>}<div style={merah}>OPERASI CYTO</div></div>
    : teks;
  const barisOp = (rows: any[]): ReactNode[][] =>
    rows.map((r: any, i: number) => [String(i + 1), r.jam, pasienCyto(r.opId, mask(r.pasien)), r.jenis, r.operator, r.anestesi, perawatSel(r.perawat)]);
  const kepalaTabel = ["No", "Jam", "Nama Pasien", "Jenis Operasi", "Dokter Operator", "Dokter Anestesi", "Perawat"];
  const kepalaLebar = ["5%", "8%", "17%", "20%", "14%", "14%", "22%"];
  return (
    <div className="kb-lembar" style={{ width: 794, boxSizing: "border-box", padding: 28, background: "#fff", color: "#111", fontFamily: "'Times New Roman',Times,serif" }}>
      <style>{".kb-lembar table,.kb-lembar th,.kb-lembar td{font-size:15px !important;font-family:Times New Roman,Times,serif !important}"}</style>
      <div style={{ textAlign: "center", fontWeight: 800, fontSize: 17 }}>LAPORAN KEPALA JAGA KAMAR BEDAH</div>
      <div style={{ textAlign: "center", fontSize: 15, marginBottom: 10 }}>RS Panti Rini</div>
      <div style={{ display: "flex", gap: 10, alignItems: "flex-start" }}>
        <table style={{ flex: 1, minWidth: 0, borderCollapse: "collapse" }}><tbody>
          {baris("Tanggal", tanggalPanjang(d.tanggal), "38%")}
          {baris("Shift", d.shiftTeks, "38%")}
          {baris("Serah terima", L ? `${SHIFT_LABEL[L.dari]} ke ${SHIFT_LABEL[L.ke]}` : "-", "38%")}
          {baris("Jumlah operasi", String(d.jumlah), "38%")}
          {baris("Operasi darurat / cyto", d.adaCyto ? "Ya" : "Tidak", "38%")}
          {baris("Total pasien", String(d.jumlah), "38%")}
        </tbody></table>
        <table style={{ flex: 1, minWidth: 0, borderCollapse: "collapse" }}><tbody>
          {baris("Siaga anestesi", (d.anest || []).join(", ") || "-", "38%")}
          {baris("Siaga anestesi cyto", (d.cyto || []).join(", ") || "-", "38%")}
          {baris("Perawat siaga", (d.nurses || []).join(", ") || "-", "38%")}
          {baris("Pembawa HP", d.pembawaHP || "-", "38%")}
        </tbody></table>
      </div>
      {judul("OPERASI PAGI")}
      {tabel(kepalaTabel, kepalaLebar, barisOp(d.pagi || []), 3)}
      {judul("OPERASI SIANG")}
      {tabel(kepalaTabel, kepalaLebar, barisOp(d.siang || []), 3)}
      {judul("OPERASI BESOK")}
      {tabel(["Tanggal", "Jam", "Dokter", "Nama", "Umur", "Jenis Operasi", "Keterangan"],
        ["12%", "8%", "18%", "20%", "10%", "22%", "10%"],
        (d.besok || []).map((r: any) => [fmtTgl(r.tanggal), r.jam, r.dokter, mask(r.pasien), r.umur, r.jenis, ketCyto(r.opId, r.keterangan)]), 3)}
      {judul("KONDISI PERALATAN & KAMAR")}
      <table style={{ width: "100%", borderCollapse: "collapse" }}><tbody>
        {ALAT_BUTIR.map((b) => <tr key={b.k}><td style={{ ...td, width: "34%", fontWeight: 700 }}>{b.l}</td><td style={td}>{L ? kondTeks((L.alat || {})[b.k]) : "-"}</td></tr>)}
        {baris("Lain-lain", L ? L.alatLain : "-")}
        {baris("Catatan", L ? L.alatCatatan : "-")}
      </tbody></table>
      {judul("STOK & BAHAN")}
      <table style={{ width: "100%", borderCollapse: "collapse" }}><tbody>
        {baris("Instrumen", L ? L.instrumen : "-")}
        {baris("Implan / Bahan Habis Pakai", L ? L.implan : "-")}
        {baris("Alkes yang dipakai", L ? L.alkesDipakai : "-")}
        {baris("Alkes persediaan", L ? L.alkesPersediaan : "-")}
        {baris("PA", L ? L.pa : "-")}
        {baris("Lain-lain", L ? L.stokLain : "-")}
        {baris("Catatan", L ? L.stokCatatan : "-")}
      </tbody></table>
      {judul("KEJADIAN / INSIDEN / KENDALA")}
      <table style={{ width: "100%", borderCollapse: "collapse" }}><tbody>
        {baris("Ada kejadian?", L ? (L.kejadian === "ya" ? "Ya" : "Tidak") : "-")}
        {baris("Jika ada, jelaskan", L && L.kejadian === "ya" ? L.uraian : "-")}
      </tbody></table>
      <div style={{ display: "flex", gap: 10, marginTop: 22, fontSize: 15, textAlign: "center" }}>
        {[["Petugas Shift Lama", L?.lama], ["Petugas Shift Baru", L?.baru], ["Mengetahui / Kepala Kamar Bedah", L?.kepala]].map(([a, b]) => (
          <div key={String(a)} style={{ flex: 1 }}>
            <div style={{ fontWeight: 700 }}>{a}</div>
            <div style={{ height: 52 }} />
            <div style={{ borderTop: "1px solid #111", paddingTop: 3 }}>Nama Jelas: <b>{b || "........................"}</b></div>
          </div>
        ))}
      </div>
    </div>
  );
}

export function ViewLaporanKepalaJaga({ ops, roster, staff, showToast, role, privacyMode, dbxCfg }: any) {
  const [tanggal, setTanggal] = useState<string>(() => isoLokal(new Date()));
  const [malamAktif, setMalamAktif] = useState(false);
  const [harian, setHarian] = useState<LaporanHarian[]>([]);
  const [daftar, setDaftar] = useState<any[]>([]);
  const [lihat, setLihat] = useState<any | null>(null);
  const [busy, setBusy] = useState(false);
  const [segar, setSegar] = useState(0);
  const [manual, setManual] = useState<Record<string, string>>({});
  const [waBuka, setWaBuka] = useState(false);
  const [waCari, setWaCari] = useState("");
  const [waTerkirim, setWaTerkirim] = useState<string[]>([]);
  const [waFormat, setWaFormat] = useState<"teks" | "jpg" | "pdf" | "word" | "excel">("teks");
  const lembarRef = useRef<HTMLDivElement>(null);
  const mask = (s: string) => maskName(s, privacyMode);
  const cytoIds = useMemo(() => new Set<string>((ops || []).filter((o: any) => o && o.opType === "cyto").map((o: any) => String(o.id))), [ops]);

  useEffect(() => {
    let batal = false;
    (async () => {
      try {
        const c = await SUPA_CLIENT.from("kamar_bedah_config").select("value").eq("key", "laporan_harian_cfg");
        const v = c.data?.[0]?.value;
        if (v && !batal) { const p = typeof v === "string" ? JSON.parse(v) : v; setMalamAktif(!!p?.malamAktif); }
        const r = await SUPA_CLIENT.from("kb_laporan_harian").select("data").order("updated_at", { ascending: false });
        if (!batal && r.data) setHarian(r.data.map((x: any) => x.data).filter(Boolean));
        const k = await SUPA_CLIENT.from(TABEL).select("data").order("updated_at", { ascending: false });
        if (!batal && k.data) setDaftar(k.data.map((x: any) => x.data).filter(Boolean));
      } catch { /* offline: bagian bawah diisi "-" */ }
    })();
    return () => { batal = true; };
  }, [tanggal, segar]);

  const peta = useMemo(() => {
    const m: Record<string, any> = {};
    (ops || []).forEach((o: any) => { if (o) m[String(o.id)] = o; });
    return m;
  }, [ops]);
  const pagi = useMemo(() => susunRencana(ops, tanggal, "pagi", malamAktif), [ops, tanggal, malamAktif]);
  const siang = useMemo(() => susunRencana(ops, tanggal, "siang", malamAktif), [ops, tanggal, malamAktif]);
  const besok = useMemo(() => susunKandidat(ops, tanggal), [ops, tanggal]);
  const opsHari = useMemo(() => (ops || []).filter((o: any) => o && o.date === tanggal && o.status !== "batal"), [ops, tanggal]);
  const rs: any = (roster || []).find((r: any) => r && r.date === tanggal);
  const v1: LaporanHarian | null = useMemo(() => {
    const list = harian.filter((l) => l.tanggal === tanggal);
    if (!list.length) return null;
    return [...list].sort((a, b) => SHIFT_URUT.indexOf(b.shift) - SHIFT_URUT.indexOf(a.shift))[0];
  }, [harian, tanggal]);

  const perawat = (id: string): string[] => {
    const o = peta[id];
    if (!o) return [];
    const pasang: [string, any][] = [["Asisten", o.assistantNurse], ["Instrumen", o.circulatingNurse], ["Anestesi", o.anesthesiaNurse], ["Onloop", o.onloopNurse], ["RR/Katim", o.rrKatim]];
    return pasang.filter(([, v]) => String(v || "").trim()).map(([k, v]) => `${k}: ${String(v).trim()}`);
  };

    const pecah = (t: string): string[] => t.split(/[,\n]/).map((x) => x.trim()).filter(Boolean);
  const mOk = manual.tgl === tanggal;
  const autoSiaga = {
    anest: rs ? daftarDari(rs.anestJagaList, rs.anestJaga) : [],
    cyto: rs ? daftarDari(rs.anestCytoList, rs.anestCyto) : [],
    nurses: rs ? (rs.nurses || []).map((x: any) => String(x || "").trim()).filter(Boolean) : [],
    hp: rs ? String(rs.pembawaHP || "").trim() : "",
  };
  const nilaiSiaga = (k: "anest" | "cyto" | "nurses" | "hp"): string => (mOk && manual[k] !== undefined ? manual[k] : k === "hp" ? autoSiaga.hp : (autoSiaga[k] as string[]).join(", "));
  const ubahSiaga = (k: string, v: string) => setManual((p) => ({ ...(p.tgl === tanggal ? p : { tgl: tanggal }), [k]: v }));
  const live: any = {
    id: "kj_" + tanggal,
    tanggal,
    shiftTeks: malamAktif ? "Pagi, Siang, Malam" : "Pagi, Siang",
    jumlah: opsHari.length,
    adaCyto: opsHari.some((o: any) => o.opType === "cyto"),
    anest: mOk && manual.anest !== undefined ? pecah(manual.anest) : autoSiaga.anest,
    cyto: mOk && manual.cyto !== undefined ? pecah(manual.cyto) : autoSiaga.cyto,
    nurses: mOk && manual.nurses !== undefined ? pecah(manual.nurses) : autoSiaga.nurses,
    pembawaHP: mOk && manual.hp !== undefined ? manual.hp.trim() : autoSiaga.hp,
    pagi: pagi.map((r: any) => ({ ...r, perawat: perawat(r.opId) })),
    siang: siang.map((r: any) => ({ ...r, perawat: perawat(r.opId) })),
    besok: besok.map((r: any) => ({ ...r })),
    v1,
  };
  const tampil: any = lihat || live;

  const ubahTanggal = (v: string) => { if (v) { setLihat(null); setTanggal(v); } };

  const simpan = async () => {
    if (lihat) return;
    if (daftar.some((x) => x.id === live.id) && !window.confirm(`Laporan ${fmtTgl(live.tanggal)} sudah tersimpan. Timpa dengan data terbaru?`)) return;
    setBusy(true);
    const nowIso = new Date().toISOString();
    const snap = { ...live, updated_at: nowIso };
    let ok = false;
    try {
      const { error } = await SUPA_CLIENT.from(TABEL).upsert({ id: snap.id, data: snap, updated_at: nowIso }, { onConflict: "id" });
      if (error) showToast("⚠ Gagal menyimpan ke server: " + (error.message || "error"), C.d);
      else { ok = true; showToast("✓ Laporan kepala jaga tersimpan", C.s); }
    } catch (e: any) { showToast("⚠ Gagal menyimpan: " + (e?.message || "error"), C.d); }
    if (ok) {
      setDaftar((x) => [snap, ...x.filter((r) => r.id !== snap.id)]);
      try {
        const folder = (dbxCfg?.path || "/kamarbedah/").replace(/[^/]+$/, "");
        const buf = await buatWorkbookKepalaJaga([snap], { cytoIds }).xlsx.writeBuffer();
        const d = await dropboxUploadBytes(`${folder}LaporanKepalaJaga_${snap.tanggal}.xlsx`, buf as unknown as ArrayBuffer);
        if (!d?.ok) showToast("⚠ Excel ke Dropbox gagal: " + (d?.msg || "error"), C.w);
      } catch { showToast("⚠ Excel ke Dropbox gagal", C.w); }
      try {
        AmbilLogMedis("LAPORAN_HARIAN", "Kepala Jaga", "SIMPAN_LAPORAN_KEPALA_JAGA",
          `Laporan kepala jaga ${snap.tanggal} disimpan`, { id: snap.id, jumlahOperasi: snap.jumlah });
      } catch { /* abaikan */ }
    }
    setBusy(false);
  };

  const hapus = async (l: any) => {
    if (!window.confirm(`Hapus laporan kepala jaga ${fmtTgl(l.tanggal)}?`)) return;
    try {
      const { error } = await SUPA_CLIENT.from(TABEL).delete().eq("id", l.id);
      if (error) { showToast("⚠ Gagal menghapus: " + (error.message || "error"), C.d); return; }
    } catch (e: any) { showToast("⚠ Gagal menghapus: " + (e?.message || "error"), C.d); return; }
    setDaftar((x) => x.filter((r) => r.id !== l.id));
    if (lihat && lihat.id === l.id) setLihat(null);
    showToast("✓ Laporan dihapus", C.s);
  };

  const cetak = () => {
    const el = lembarRef.current;
    if (!el) { showToast("Lembar belum siap dicetak", C.w); return; }
    const ifr = document.createElement("iframe");
    ifr.style.cssText = "position:fixed;right:0;bottom:0;width:0;height:0;border:0";
    document.body.appendChild(ifr);
    const w = ifr.contentWindow;
    if (!w) { ifr.remove(); return; }
    w.document.open();
    w.document.write("<!doctype html><html><head><meta charset='utf-8'><title>Laporan Kepala Jaga</title><style>@page{size:A4;margin:6mm}body{margin:0}.kb-lembar{zoom:0.94}</style></head><body>" + el.innerHTML + "</body></html>");
    w.document.close();
    setTimeout(() => { w.focus(); w.print(); setTimeout(() => ifr.remove(), 3000); }, 400);
  };
  const ambilCanvas = async (skala = 2) => {
    const el = lembarRef.current;
    if (!el) throw new Error("Lembar belum siap");
    return html2canvas(el, { backgroundColor: "#ffffff", scale: skala });
  };
  const namaFile = (ext: string) => `LaporanKepalaJaga_${tampil.tanggal}.${ext}`;
  const unduhJpg = async () => {
    setBusy(true);
    try {
      const cv = await ambilCanvas(3);
      const a = document.createElement("a");
      a.href = cv.toDataURL("image/jpeg", 0.95); a.download = namaFile("jpg"); a.click();
      showToast("✓ JPG diunduh", C.s);
    } catch (e: any) { showToast("Gagal membuat JPG: " + (e?.message || "error"), C.d); }
    setBusy(false);
  };
  const unduhPdf = async () => {
    setBusy(true);
    try {
      const cv = await ambilCanvas();
      const PW = 595.28; const PH = 841.89;
      const tinggiHal = Math.floor((cv.width * PH) / PW);
      const hal: { jpeg: Uint8Array; w: number; h: number }[] = [];
      for (let y = 0; y < cv.height; y += tinggiHal) {
        const h = Math.min(tinggiHal, cv.height - y);
        const part = document.createElement("canvas");
        part.width = cv.width; part.height = h;
        const ctx = part.getContext("2d");
        if (!ctx) throw new Error("Canvas tidak tersedia");
        ctx.fillStyle = "#ffffff"; ctx.fillRect(0, 0, part.width, h);
        ctx.drawImage(cv, 0, y, cv.width, h, 0, 0, cv.width, h);
        const bin = atob(part.toDataURL("image/jpeg", 0.92).split(",")[1]);
        const arr = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
        hal.push({ jpeg: arr, w: part.width, h });
      }
      const url = URL.createObjectURL(buatPdfDariJpeg(hal));
      const a = document.createElement("a");
      a.href = url; a.download = namaFile("pdf"); a.click();
      setTimeout(() => URL.revokeObjectURL(url), 5000);
      showToast("✓ PDF diunduh", C.s);
    } catch (e: any) { showToast("Gagal membuat PDF: " + (e?.message || "error"), C.d); }
    setBusy(false);
  };
  const salinTeks = async () => {
    try { await navigator.clipboard.writeText(buatTeks(tampil, mask)); showToast("✓ Teks laporan disalin", C.s); }
    catch { showToast("Gagal menyalin teks", C.w); }
  };
  const labelStaf = (s: any): string => (s.type === "surgeon" ? "Dokter bedah" : s.type === "anesthesiologist" ? "Dokter anestesi" : "Perawat");
  const daftarWa = (): any[] => {
    const q = waCari.trim().toLowerCase();
    return (staff || [])
      .filter((s: any) => s && s.name && String(s.phone || "").trim() && (!q || String(s.name).toLowerCase().includes(q)))
      .sort((a: any, b: any) => String(a.name).localeCompare(String(b.name)));
  };
  const kirimWa = (s: any) => {
    const ph = normalizePhone62(String(s.phone));
    window.open("https://wa.me/" + ph + "?text=" + encodeURIComponent(buatTeks(tampil, mask)), "_blank");
    setWaTerkirim((x) => (x.includes(String(s.id)) ? x : [...x, String(s.id)]));
  };
  const buatFile = async (fmt: string): Promise<{ blob: Blob; nama: string }> => {
    if (fmt === "word") {
      const el = lembarRef.current;
      if (!el) throw new Error("Lembar belum siap");
      return { blob: elementToWordBlob(el), nama: namaFile("doc") };
    }
    if (fmt === "excel") {
      const buf = await buatWorkbookKepalaJaga([tampil], { cytoIds }).xlsx.writeBuffer();
      return { blob: new Blob([buf as unknown as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), nama: namaFile("xlsx") };
    }
    const cv = await ambilCanvas(fmt === "jpg" ? 3 : 2);
    if (fmt === "jpg") {
      const blob = await new Promise<Blob | null>((res) => cv.toBlob((b) => res(b), "image/jpeg", 0.95));
      if (!blob) throw new Error("Gagal membuat JPG");
      return { blob, nama: namaFile("jpg") };
    }
    const PW = 595.28; const PH = 841.89;
    const tinggiHal = Math.floor((cv.width * PH) / PW);
    const hal: { jpeg: Uint8Array; w: number; h: number }[] = [];
    for (let y = 0; y < cv.height; y += tinggiHal) {
      const h = Math.min(tinggiHal, cv.height - y);
      const part = document.createElement("canvas");
      part.width = cv.width; part.height = h;
      const ctx = part.getContext("2d");
      if (!ctx) throw new Error("Canvas tidak tersedia");
      ctx.fillStyle = "#ffffff"; ctx.fillRect(0, 0, part.width, h);
      ctx.drawImage(cv, 0, y, cv.width, h, 0, 0, cv.width, h);
      const bin = atob(part.toDataURL("image/jpeg", 0.92).split(",")[1]);
      const arr = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
      hal.push({ jpeg: arr, w: part.width, h });
    }
    return { blob: buatPdfDariJpeg(hal), nama: namaFile("pdf") };
  };
  const unduhBlob = (blob: Blob, nama: string) => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = nama; a.click();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  };
  const kirimWaPilihan = async (s: any | null) => {
    if (waFormat === "teks") { if (s) kirimWa(s); else kirimWaBebas(); return; }
    const jendela = window.open("", "_blank");
    setBusy(true);
    try {
      const { blob, nama } = await buatFile(waFormat);
      unduhBlob(blob, nama);
      const pesan = "Laporan Kepala Jaga Kamar Bedah " + fmtTgl(tampil.tanggal) + " (file " + nama + " terlampir)";
      const link = "https://wa.me/" + (s ? normalizePhone62(String(s.phone)) : "") + "?text=" + encodeURIComponent(pesan);
      if (jendela) jendela.location.href = link; else window.open(link, "_blank");
      if (s) setWaTerkirim((x) => (x.includes(String(s.id)) ? x : [...x, String(s.id)]));
      showToast("✓ File diunduh. Lampirkan " + nama + " di chat WhatsApp", C.s);
    } catch (e: any) {
      if (jendela) jendela.close();
      showToast("Gagal menyiapkan file: " + (e?.message || "error"), C.d);
    }
    setBusy(false);
  };
  const bagikanFile = async () => {
    if (waFormat === "teks") return;
    setBusy(true);
    try {
      const { blob, nama } = await buatFile(waFormat);
      const file = new File([blob], nama, { type: blob.type || "application/octet-stream" });
      const nav: any = navigator;
      const pesan = "Laporan Kepala Jaga Kamar Bedah " + fmtTgl(tampil.tanggal);
      if (nav.canShare && nav.canShare({ files: [file] })) {
        await nav.share({ files: [file], text: pesan, title: pesan });
      } else {
        unduhBlob(blob, nama);
        showToast("Perangkat ini belum mendukung bagikan file langsung. File diunduh, lampirkan manual di WhatsApp.", C.w);
      }
    } catch (e: any) {
      if (e?.name !== "AbortError") showToast("Gagal membagikan file: " + (e?.message || "error"), C.d);
    }
    setBusy(false);
  };
  const kirimWaBebas = () => {
    window.open("https://wa.me/?text=" + encodeURIComponent(buatTeks(tampil, mask)), "_blank");
  };
  const unduhWord = () => {
    const el = lembarRef.current;
    if (!el) { showToast("Lembar belum siap", C.w); return; }
    downloadElementAsWord(el, namaFile("doc"));
    showToast("✓ Word diunduh", C.s);
  };
  const unduhExcel = () => {
    buatWorkbookKepalaJaga([tampil], { cytoIds }).xlsx.writeBuffer().then((buf) => {
      const url = URL.createObjectURL(new Blob([buf as unknown as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
      const a = document.createElement("a");
      a.href = url; a.download = namaFile("xlsx"); a.click();
      setTimeout(() => URL.revokeObjectURL(url), 5000);
      showToast("✓ Excel diunduh", C.s);
    }).catch((e) => showToast("Gagal membuat Excel: " + ((e as any)?.message || "error"), C.d));
  };

  const riwayat = [...daftar].sort((a, b) => String(b.tanggal).localeCompare(String(a.tanggal))).slice(0, 8);

  return (
    <div style={{ paddingBottom: 24 }}>
      <Card>
        <SH label="Laporan Kepala Jaga" />
        <LF label="Tanggal laporan">
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            <input style={iS} type="date" value={tampil.tanggal} onChange={(e) => ubahTanggal(e.target.value)} />
            <Btn sm outline onClick={() => setSegar((n) => n + 1)} disabled={busy}>↻ Muat ulang</Btn>
          </div>
        </LF>
        {lihat
          ? <div style={{ fontSize: 13, fontWeight: 700, color: "#166534", background: "#DCFCE7", borderRadius: 10, padding: "8px 12px", marginTop: 8 }}>
              ✓ Menampilkan salinan tersimpan ({fmtTgl(lihat.tanggal)}). Data hari ini tidak dipakai.
            </div>
          : <>
              {!rs && <div style={{ fontSize: 12, color: C.w, marginTop: 6 }}>⚠ Belum ada jadwal siaga untuk tanggal ini. Bagian siaga di lembar kosong.</div>}
              <div style={{ fontSize: 12, color: C.tL, marginTop: 6 }}>
                {v1
                  ? `Bagian bawah (peralatan, stok, kejadian, tanda tangan) diambil dari Laporan Harian OK ${SHIFT_LABEL[v1.shift]} tanggal ini.`
                  : "Belum ada Laporan Harian OK untuk tanggal ini. Bagian bawah diisi tanda “-”."}
              </div>
            </>}
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 10 }}>
          {!lihat && (
            <div style={{ flexBasis: "100%", border: "1px solid " + C.b, borderRadius: 10, padding: 10 }}>
              <div style={{ fontWeight: 700, marginBottom: 6 }}>✎ Siaga dan Pembawa HP (isi manual bila ada perubahan)</div>
              {([["anest", "Siaga anestesi"], ["cyto", "Siaga anestesi cyto"], ["nurses", "Perawat siaga"], ["hp", "Pembawa HP"]] as const).map(([k, l]) => (
                <div key={k} style={{ marginBottom: 6 }}>
                  <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 2 }}>{l}</div>
                  <input style={iS} value={nilaiSiaga(k)} onChange={(e) => ubahSiaga(k, e.target.value)} placeholder={k === "hp" ? "Nama pembawa HP" : "Pisahkan nama dengan koma"} />
                </div>
              ))}
              <Btn sm outline onClick={() => setManual({})}>↺ Kembalikan ke jadwal</Btn>
            </div>
          )}
          {!lihat && (() => {
            const sudah = daftar.some((x) => x.id === live.id);
            return (
              <div style={{ flexBasis: "100%" }}>
                {sudah
                  ? <div style={{ fontSize: 14, fontWeight: 700, color: "#166534", background: "#DCFCE7", borderRadius: 10, padding: "10px 12px", marginBottom: 8 }}>
                      ✓ Laporan tanggal ini sudah tersimpan. Jika ada perubahan operasi atau jadwal, klik SIMPAN lagi untuk memperbarui.
                    </div>
                  : <div style={{ fontSize: 20, fontWeight: 800, color: "#fff", background: "#DC2626", borderRadius: 10, padding: "14px 12px", marginBottom: 8, textAlign: "center", lineHeight: 1.4 }}>
                      ⚠ HARAP KLIK SIMPAN! Laporan belum tersimpan.
                    </div>}
                <button type="button" onClick={simpan} disabled={busy}
                  style={{ width: "100%", fontSize: 20, fontWeight: 800, padding: "16px 12px", borderRadius: 12, border: "none", cursor: busy ? "default" : "pointer", opacity: busy ? 0.6 : 1, color: "#fff", background: sudah ? "#16A34A" : "#DC2626" }}>
                  {busy ? "Menyimpan..." : sudah ? "💾 SIMPAN ULANG" : "💾 SIMPAN LAPORAN"}
                </button>
              </div>
            );
          })()}
          <Btn outline onClick={cetak} disabled={busy}>🖨 Print</Btn>
          <Btn outline onClick={unduhPdf} disabled={busy}>⬇ PDF</Btn>
          <Btn outline onClick={unduhJpg} disabled={busy}>⬇ JPG</Btn>
          <Btn outline onClick={salinTeks} disabled={busy}>📋 Salin teks</Btn>
          <Btn outline onClick={unduhExcel} disabled={busy}>⬇ Excel</Btn>
          <Btn outline onClick={unduhWord} disabled={busy}>⬇ Word</Btn>
          <button type="button" onClick={() => setWaBuka((v) => !v)} disabled={busy}
            style={{ padding: "10px 18px", borderRadius: 10, border: "none", background: "#16A34A", color: "#fff", fontSize: 15, fontWeight: 800, cursor: "pointer", boxShadow: "0 2px 8px rgba(22,163,74,.5)" }}>📲 KIRIM WA</button>
          {waBuka && (
            <div style={{ flexBasis: "100%", border: "2px solid #16A34A", borderRadius: 10, padding: 10, background: "#F0FDF4" }}>
              <div style={{ fontWeight: 800, marginBottom: 6 }}>Pilih penerima laporan ({fmtTgl(tampil.tanggal)})</div>
              <div style={{ fontWeight: 700, fontSize: 13, marginBottom: 4 }}>Kirim dalam bentuk:</div>
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 8 }}>
                {([["teks", "💬 Teks"], ["jpg", "🖼 JPG"], ["pdf", "📄 PDF"], ["word", "📝 Word"], ["excel", "📊 Excel"]] as const).map(([k2, l]) => (
                  <button key={k2} type="button" onClick={() => setWaFormat(k2)}
                    style={{ padding: "8px 12px", borderRadius: 10, border: "2px solid #16A34A", background: waFormat === k2 ? "#16A34A" : "#fff", color: waFormat === k2 ? "#fff" : "#166534", fontSize: 14, fontWeight: 800, cursor: "pointer" }}>{l}</button>
                ))}
              </div>
              {waFormat !== "teks" && (
                <div style={{ marginBottom: 8 }}>
                  <Btn onClick={bagikanFile} disabled={busy}>📤 Bagikan file {waFormat.toUpperCase()} langsung ke WhatsApp (HP)</Btn>
                  <div style={{ fontSize: 12, color: C.tL, marginTop: 4 }}>Di HP: pilih WhatsApp lalu penerima atau grup. Tombol Kirim per nama: file diunduh, chat WhatsApp terbuka, lalu Anda tinggal melampirkan file yang baru diunduh.</div>
                </div>
              )}
              <input style={iS} value={waCari} onChange={(e) => setWaCari(e.target.value)} placeholder="Cari nama..." />
              <div style={{ maxHeight: 260, overflowY: "auto", marginTop: 8 }}>
                {daftarWa().map((s: any) => (
                  <div key={s.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8, padding: "6px 0", borderBottom: "1px solid " + C.b }}>
                    <div style={{ fontSize: 14 }}><b>{s.name}</b><div style={{ fontSize: 12, color: C.tL }}>{labelStaf(s)}</div></div>
                    <button type="button" onClick={() => kirimWaPilihan(s)}
                      style={{ padding: "8px 14px", borderRadius: 10, border: "none", background: waTerkirim.includes(String(s.id)) ? "#15803D" : "#16A34A", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>
                      {waTerkirim.includes(String(s.id)) ? "✓ Kirim lagi" : "Kirim"}
                    </button>
                  </div>
                ))}
                {!daftarWa().length && <div style={{ fontSize: 13, color: C.tL, padding: "6px 0" }}>Tidak ada staf bernomor WA yang cocok.</div>}
              </div>
              <div style={{ marginTop: 8 }}><Btn sm outline onClick={() => kirimWaPilihan(null)}>Pilih kontak / grup di WhatsApp</Btn></div>
            </div>
          )}
          {lihat && <Btn outline onClick={() => setLihat(null)} disabled={busy}>✕ Tutup</Btn>}
        </div>
      </Card>

      <div style={{ overflowX: "auto", marginTop: 12, background: "#E5E7EB", padding: 8, borderRadius: 10 }}>
        <div ref={lembarRef} style={{ width: 794 }}><Lembar2 d={tampil} mask={mask} cytoIds={cytoIds} /></div>
      </div>

      <Card style={{ marginTop: 12 }}>
        <SH label="Riwayat tersimpan" />
        {!riwayat.length && <div style={{ fontSize: 13, color: C.tL }}>Belum ada laporan kepala jaga tersimpan.</div>}
        {riwayat.map((l) => (
          <div key={l.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8, padding: "6px 0", borderBottom: `1px solid ${C.b}` }}>
            <div style={{ fontSize: 13 }}><b>{fmtTgl(l.tanggal)}</b> • {l.jumlah} operasi</div>
            <div style={{ display: "flex", gap: 6 }}>
              <Btn sm outline onClick={() => { setLihat(l); window.scrollTo({ top: 0, behavior: "smooth" }); }}>Buka</Btn>
              <Btn sm outline onClick={() => { setLihat(l); window.scrollTo({ top: 0 }); setTimeout(cetak, 700); }}>🖨 Print</Btn>
              {role === "admin" && <Btn sm outline color={C.d} onClick={() => hapus(l)}>Hapus</Btn>}
            </div>
          </div>
        ))}
      </Card>
    </div>
  );
}

/* Pemilih versi di atas tab. Kedua tampilan tetap terpasang (hanya disembunyikan)
   supaya preview Laporan Harian OK yang belum disimpan tidak hilang saat berpindah. */
export function LaporanOKTab(props: any) {
  const [v, setV] = useState<"harian" | "kj">("harian");
  return (
    <div>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 10 }}>
        <Btn onClick={() => setV("harian")} outline={v !== "harian"}>📋 Laporan Harian OK</Btn>
        <Btn onClick={() => setV("kj")} outline={v !== "kj"}>📝 Laporan Kepala Jaga</Btn>
      </div>
      <div style={{ display: v === "harian" ? "block" : "none" }}><ViewLaporanHarian {...props} /></div>
      <div style={{ display: v === "kj" ? "block" : "none" }}><ViewLaporanKepalaJaga {...props} /></div>
    </div>
  );
}
