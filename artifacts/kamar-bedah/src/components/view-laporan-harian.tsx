/* src/components/view-laporan-harian.tsx
   Tab "Laporan Harian OK" — laporan serah terima antar shift. Data jadwal
   diambil dari daftar operasi (ops), isian lain manual. Preview bergaya
   formulir kertas, bisa diunduh PDF/JPG. Disimpan ke Supabase
   (kb_laporan_harian), Excel bulanan ke Dropbox, dan log audit ke Turso
   (tanpa nama pasien). Aturan shift ada di utils/laporan-harian.ts. */

import { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import html2canvas from "html2canvas";
import { C, iS } from "../constants";
import { maskName } from "../utils";
import { SUPA_CLIENT, AmbilLogMedis } from "../services/supabase";
import { dropboxUploadExcel } from "../services/dropbox";
import { Card, SH, LF, Btn, Toggle } from "./primitives";
import type { ShiftKey, LaporanHarian, JadwalRow, RencanaRow, KandidatRow } from "../utils/laporan-harian";
import {
  SHIFT_URUT, SHIFT_LABEL, STOK_OPSI, STATUS_OPSI, ALAT_BUTIR,
  pad2, isoLokal, tanggalPanjang, fmtTgl, namaHari, shiftDariJam, shiftBerikut, tanggalTujuan,
  susunJadwal, susunRencana, susunKandidat, rencanaHariPenuh, susunRencanaHariPenuh, buatTeksLaporan, buatWorkbookLaporan, buatPdfDariJpeg, filterLaporan, urutLaporanTurun,
} from "../utils/laporan-harian";

const DRAFT_KEY = "kb_laporan_harian_draft_v1";

interface FormState {
  tanggal: string; shift: ShiftKey; dari: ShiftKey | ""; ke: ShiftKey | "";
  jumlahManual: string; totalManual: string; cyto: "" | "ya" | "tidak";
  ov: Record<string, any>;
  alat: Record<string, string>; alatLain: string; alatCatatan: string;
  instrumen: string; implan: string; alkesDipakai: string; alkesPersediaan: string; pa: string; stokLain: string; stokCatatan: string;
  kejadian: "" | "ya" | "tidak"; uraian: string;
  lama: string; baru: string; kepala: string;
}
function formAwal(tanggal: string, shift: ShiftKey): FormState {
  return { tanggal, shift, dari: "", ke: "", jumlahManual: "", totalManual: "", cyto: "", ov: {}, alat: {}, alatLain: "", alatCatatan: "",
    instrumen: "", implan: "", alkesDipakai: "", alkesPersediaan: "", pa: "", stokLain: "", stokCatatan: "", kejadian: "", uraian: "", lama: "", baru: "", kepala: "" };
}

function Pilih({ value, onChange, opsi }: { value: string; onChange: (v: string) => void; opsi: { v: string; l: string; c?: string; off?: boolean }[] }) {
  return (
    <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
      {opsi.map((o) => {
        const on = value === o.v; const warna = o.c || C.p;
        return (
          <button key={o.v} type="button" disabled={o.off} onClick={() => onChange(o.v)}
            style={{ flex: 1, minWidth: 74, padding: "8px 10px", borderRadius: 10, border: `1.5px solid ${on ? warna : C.b}`, background: on ? warna : "#FAFBFD",
              color: on ? "#fff" : (o.off ? C.tL : C.t), fontSize: 13, fontWeight: 700, cursor: o.off ? "not-allowed" : "pointer", opacity: o.off ? 0.5 : 1, fontFamily: "inherit" }}>
            {o.l}
          </button>
        );
      })}
    </div>
  );
}

/* ── Lembar preview bergaya formulir kertas (juga dipakai untuk PDF/JPG) ── */
function Lembar({ l, mask }: { l: LaporanHarian; mask: (s: string) => string }) {
  const th: React.CSSProperties = { border: "1px solid #111", padding: "4px 5px", fontSize: 11, fontWeight: 700, textAlign: "center", background: "#EEF2F7" };
  const td: React.CSSProperties = { border: "1px solid #111", padding: "4px 5px", fontSize: 11, verticalAlign: "top", height: 20 };
  const judul = (t: string) => <div style={{ fontWeight: 800, fontSize: 13, margin: "14px 0 5px" }}>{t}</div>;
  const tabel = (head: string[], rows: string[][], minRows: number) => {
    const isi = [...rows];
    while (isi.length < minRows) isi.push(head.map(() => ""));
    return (
      <table style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead><tr>{head.map((h, i) => <th key={i} style={th}>{h}</th>)}</tr></thead>
        <tbody>{isi.map((r, i) => <tr key={i}>{r.map((c, j) => <td key={j} style={td}>{c}</td>)}</tr>)}</tbody>
      </table>
    );
  };
  const kond = (v: string) => (v === "baik" ? "Baik" : v === "tidak" ? "Tidak baik" : "-");
  const baris = (a: string, b: string) => <tr><td style={{ ...td, width: "34%", fontWeight: 700 }}>{a}</td><td style={td}>{b}</td></tr>;
  return (
    <div style={{ width: 794, boxSizing: "border-box", padding: 28, background: "#fff", color: "#111", fontFamily: "'Times New Roman',Times,serif" }}>
      <div style={{ textAlign: "center", fontWeight: 800, fontSize: 17 }}>LAPORAN HARIAN KAMAR BEDAH</div>
      <div style={{ textAlign: "center", fontSize: 12, marginBottom: 10 }}>RS Panti Rini</div>
      <table style={{ width: "100%", borderCollapse: "collapse" }}><tbody>
        {baris("Tanggal", tanggalPanjang(l.tanggal))}
        {baris("Shift", SHIFT_LABEL[l.shift])}
        {baris("Serah terima", `${SHIFT_LABEL[l.dari]} ke ${SHIFT_LABEL[l.ke]}`)}
        {baris("Jumlah operasi", l.jumlahOperasi)}
        {baris("Operasi darurat / cyto", l.adaCyto === "ya" ? "Ya" : "Tidak")}
        {baris("Total pasien", l.totalPasien)}
      </tbody></table>
      {judul("JADWAL OPERASI")}
      {tabel(["No", "Jam", "Nama Pasien", "Jenis Operasi", "Dokter Operator", "Dokter Anestesi", "Status", "Catatan"],
        l.jadwal.map((r, i) => [String(i + 1), r.jam, mask(r.pasien), r.jenis, r.operator, r.anestesi, r.status, r.catatan]), 10)}
      {judul(`RENCANA OPERASI — ${rencanaHariPenuh(l.dari, l.ke) ? "Semua operasi besok" : SHIFT_LABEL[l.ke]}, ${tanggalPanjang(l.tanggalRencana)}`)}
      {tabel(["Tanggal", "Jam", "Dokter Operator", "Dokter Anestesi", "Nama", "Umur", "Kamar", "Jenis Operasi", "Alkes", "Keterangan"],
        l.rencana.map((r) => [fmtTgl(r.tanggal), r.jam, r.operator, r.anestesi, mask(r.pasien), r.umur, r.kamar, r.jenis, r.alkes, r.keterangan]), 4)}
      {judul("KANDIDAT (OPERASI BESOK)")}
      {tabel(["Tanggal", "Jam", "Dokter", "Nama", "Umur", "Jenis Operasi", "Keterangan"],
        l.kandidat.map((r) => [fmtTgl(r.tanggal), r.jam, r.dokter, mask(r.pasien), r.umur, r.jenis, r.keterangan]), 3)}
      {judul("KONDISI PERALATAN & KAMAR")}
      <table style={{ width: "100%", borderCollapse: "collapse" }}><tbody>
        {ALAT_BUTIR.map((b) => <tr key={b.k}>{[<td key="a" style={{ ...td, width: "34%", fontWeight: 700 }}>{b.l}</td>, <td key="b" style={td}>{kond(l.alat[b.k])}</td>]}</tr>)}
        {baris("Lain-lain", l.alatLain)}
        {baris("Catatan", l.alatCatatan)}
      </tbody></table>
      {judul("STOK & BAHAN")}
      <table style={{ width: "100%", borderCollapse: "collapse" }}><tbody>
        {baris("Instrumen", l.instrumen)}
        {baris("Implan / Bahan Habis Pakai", l.implan)}
        {baris("Alkes yang dipakai", l.alkesDipakai)}
        {baris("Alkes persediaan", l.alkesPersediaan)}
        {baris("PA", l.pa)}
        {baris("Lain-lain", l.stokLain)}
        {baris("Catatan", l.stokCatatan)}
      </tbody></table>
      {judul("KEJADIAN / INSIDEN / KENDALA")}
      <table style={{ width: "100%", borderCollapse: "collapse" }}><tbody>
        {baris("Ada kejadian?", l.kejadian === "ya" ? "Ya" : "Tidak")}
        {baris("Jika ada, jelaskan", l.kejadian === "ya" ? l.uraian : "-")}
      </tbody></table>
      <div style={{ display: "flex", gap: 10, marginTop: 22, fontSize: 12, textAlign: "center" }}>
        {[["Petugas Shift Lama", l.lama], ["Petugas Shift Baru", l.baru], ["Mengetahui / Kepala Kamar Bedah", l.kepala]].map(([a, b]) => (
          <div key={a} style={{ flex: 1 }}>
            <div style={{ fontWeight: 700 }}>{a}</div>
            <div style={{ height: 52 }} />
            <div style={{ borderTop: "1px solid #111", paddingTop: 3 }}>Nama Jelas: <b>{b || "........................"}</b></div>
          </div>
        ))}
      </div>
    </div>
  );
}

export function ViewLaporanHarian({ ops, staff, showToast, role, privacyMode, upsertOneToSupa, deleteFromSupa, dbxCfg }: any) {
  const [malamAktif, setMalamAktif] = useState(false);
  const [now, setNow] = useState<Date>(() => new Date());
  const [laporanList, setLaporanList] = useState<LaporanHarian[]>([]);
  const [pratinjau, setPratinjau] = useState<LaporanHarian | null>(null);
  const [busy, setBusy] = useState(false);
  const lembarRef = useRef<HTMLDivElement>(null);
  const [dlMode, setDlMode] = useState<"tanggal" | "bulan" | "tahun">("bulan");
  const hariIni = isoLokal(now);
  const [dlTgl, setDlTgl] = useState(hariIni);
  const [dlBln, setDlBln] = useState(hariIni.slice(0, 7));
  const [dlThn, setDlThn] = useState(hariIni.slice(0, 4));

  const [form, setForm] = useState<FormState>(() => {
    const n = new Date(); const t = isoLokal(n);
    const awal = formAwal(t, shiftDariJam(`${pad2(n.getHours())}:${pad2(n.getMinutes())}`, false));
    try {
      const raw = localStorage.getItem(DRAFT_KEY);
      if (raw) { const d = JSON.parse(raw); if (d && d.tanggal === t) return { ...awal, ...d }; }
    } catch { /* draf rusak — abaikan */ }
    return awal;
  });
  const set = (patch: Partial<FormState>) => setForm((p) => ({ ...p, ...patch }));
  const setOv = (key: string, patch: Record<string, any>) => setForm((p) => ({ ...p, ov: { ...p.ov, [key]: { ...(p.ov[key] || {}), ...patch } } }));

  const mask = (s: string) => maskName(s, privacyMode);

  /* jam berjalan (24 jam) */
  useEffect(() => { const i = setInterval(() => setNow(new Date()), 30000); return () => clearInterval(i); }, []);
  /* draf otomatis di perangkat */
  useEffect(() => { try { localStorage.setItem(DRAFT_KEY, JSON.stringify(form)); } catch { /* storage penuh/diblokir */ } }, [form]);

  /* muat awal + realtime */
  useEffect(() => {
    let batal = false;
    (async () => {
      try {
        const r = await SUPA_CLIENT.from("kb_laporan_harian").select("data").order("updated_at", { ascending: false });
        if (!batal && r.data) setLaporanList(r.data.map((x: any) => x.data).filter(Boolean));
        const c = await SUPA_CLIENT.from("kamar_bedah_config").select("value").eq("key", "laporan_harian_cfg");
        const v = c.data?.[0]?.value;
        if (v && !batal) { const p = typeof v === "string" ? JSON.parse(v) : v; setMalamAktif(!!p?.malamAktif); }
      } catch { /* offline — pakai state lokal */ }
    })();
    const ch = SUPA_CLIENT.channel("kb-laporan-harian-rt")
      .on("postgres_changes", { event: "*", schema: "public", table: "kb_laporan_harian" }, (p: any) => {
        if (p.eventType === "DELETE") { const id = p.old?.id; if (id) setLaporanList((x) => x.filter((r) => r.id !== id)); }
        else if (p.new?.data) { const d = p.new.data as LaporanHarian; setLaporanList((x) => [d, ...x.filter((r) => r.id !== d.id)]); }
      }).subscribe();
    return () => { batal = true; SUPA_CLIENT.removeChannel(ch); };
  }, []);

  const simpanCfg = async (v: boolean) => {
    setMalamAktif(v);
    if (!v && form.shift === "malam") set({ shift: "siang" });
    const { error } = await SUPA_CLIENT.from("kamar_bedah_config").upsert({ key: "laporan_harian_cfg", value: JSON.stringify({ malamAktif: v }) }, { onConflict: "key" });
    showToast(error ? "⚠ Pengaturan shift malam gagal disimpan ke server" : (v ? "✓ Shift Malam diaktifkan" : "✓ Shift Malam dinonaktifkan"), error ? C.d : C.s);
  };

  /* ── turunan dari form & data operasi ── */
  const dariEff: ShiftKey = (form.dari || form.shift) as ShiftKey;
  const keEff: ShiftKey = (form.ke || shiftBerikut(form.shift, malamAktif)) as ShiftKey;
  const tglRencana = tanggalTujuan(form.tanggal, dariEff, keEff);

  const jadwalAuto = useMemo(() => susunJadwal(ops, form.tanggal, form.shift, malamAktif), [ops, form.tanggal, form.shift, malamAktif]);
  const jadwal: JadwalRow[] = useMemo(() => jadwalAuto.map((r) => { const o = form.ov["j:" + r.opId] || {}; return { ...r, status: o.status ?? r.status, catatan: o.catatan ?? r.catatan }; }), [jadwalAuto, form.ov]);
  const penuh = rencanaHariPenuh(dariEff, keEff);
  const rencanaAuto = useMemo(() => penuh ? susunRencanaHariPenuh(ops, tglRencana) : susunRencana(ops, tglRencana, keEff, malamAktif), [ops, tglRencana, keEff, malamAktif, penuh]);
  const rencana: RencanaRow[] = useMemo(() => rencanaAuto.map((r) => { const o = form.ov["r:" + r.opId] || {}; return { ...r, alkes: o.alkes ?? r.alkes, keterangan: o.keterangan ?? r.keterangan }; }), [rencanaAuto, form.ov]);
  const kandidat: KandidatRow[] = useMemo(() => penuh ? [] : susunKandidat(ops, form.tanggal), [ops, form.tanggal, penuh]);

  const jumlahOperasi = form.jumlahManual !== "" ? form.jumlahManual : String(jadwal.length);
  const totalPasien = form.totalManual !== "" ? form.totalManual : String(jadwal.filter((r) => r.status !== "Batal").length);
  const adaCyto: "ya" | "tidak" = (form.cyto || (jadwal.some((r) => r.cyto) ? "ya" : "tidak")) as "ya" | "tidak";

  const validasi = (): string => {
    if (!form.tanggal) return "Tanggal belum diisi";
    if (form.shift === "malam" && !malamAktif) return "Shift Malam sedang dinonaktifkan";
    const belum = ALAT_BUTIR.filter((b) => !form.alat[b.k]);
    if (belum.length) return "Kondisi peralatan belum dipilih: " + belum.map((b) => b.l).join(", ");
    if (!form.instrumen || !form.implan) return "Pilihan Instrumen dan Implan/BHP belum dipilih";
    if (!form.kejadian) return "Pilih Ada Kejadian: Ya atau Tidak";
    if (form.kejadian === "ya" && !form.uraian.trim()) return "Jelaskan kejadian yang terjadi";
    if (!form.lama.trim() || !form.baru.trim()) return "Nama Petugas Shift Lama dan Shift Baru wajib diisi";
    return "";
  };

  const bangunLaporan = (): LaporanHarian => {
    const id = `${form.tanggal}_${form.shift}`;
    const sudah = laporanList.find((x) => x.id === id);
    const iso = new Date().toISOString();
    return {
      id, tanggal: form.tanggal, hari: namaHari(form.tanggal), shift: form.shift, dari: dariEff, ke: keEff, tanggalRencana: tglRencana,
      jumlahOperasi, adaCyto, totalPasien, jadwal, rencana, kandidat,
      alat: form.alat, alatLain: form.alatLain, alatCatatan: form.alatCatatan,
      instrumen: form.instrumen, implan: form.implan, alkesDipakai: form.alkesDipakai, alkesPersediaan: form.alkesPersediaan, pa: form.pa,
      stokLain: form.stokLain, stokCatatan: form.stokCatatan, kejadian: form.kejadian as "ya" | "tidak", uraian: form.uraian,
      lama: form.lama.trim(), baru: form.baru.trim(), kepala: form.kepala.trim(),
      operator: role === "admin" ? "Admin" : "Perawat", createdAt: sudah?.createdAt || iso, updated_at: iso,
    };
  };

  const keAtasPreview = () => {
    const err = validasi();
    if (err) { showToast("⚠ " + err, C.w); return; }
    setPratinjau(bangunLaporan());
    window.scrollTo({ top: 0, behavior: "smooth" });
  };

  const simpan = async () => {
    if (!pratinjau) return;
    setBusy(true);
    const l = pratinjau;
    const semua = [l, ...laporanList.filter((x) => x.id !== l.id)];
    setLaporanList(semua);
    const res = await upsertOneToSupa("kb_laporan_harian", l);
    if (!res?.ok) showToast(`⚠ Gagal menyimpan ke server: ${res?.error || "error"} (tersimpan di perangkat ini)`, C.d);
    else showToast("✓ Laporan harian tersimpan", C.s);
    try {
      const folder = (dbxCfg?.path || "/kamarbedah/").replace(/[^/]+$/, "");
      const ym = l.tanggal.slice(0, 7);
      const wb = buatWorkbookLaporan(filterLaporan(semua, "bulan", ym));
      const d = await dropboxUploadExcel(`${folder}LaporanHarian_${ym}.xlsx`, wb);
      if (!d?.ok) showToast("⚠ Excel ke Dropbox gagal: " + (d?.msg || "error"), C.w);
    } catch { showToast("⚠ Excel ke Dropbox gagal", C.w); }
    AmbilLogMedis("LAPORAN_HARIAN", l.operator, "SIMPAN_LAPORAN_HARIAN",
      `Laporan harian ${l.tanggal} ${SHIFT_LABEL[l.shift]} disimpan`, { id: l.id, shift: l.shift, jumlahOperasi: l.jumlahOperasi });
    if (res?.ok) { try { localStorage.removeItem(DRAFT_KEY); } catch { /* abaikan */ } }
    setBusy(false);
  };

  const ambilCanvas = async () => {
    const el = lembarRef.current;
    if (!el) throw new Error("Lembar belum siap");
    return html2canvas(el, { backgroundColor: "#ffffff", scale: 2 });
  };
  const namaFile = (ext: string) => `LaporanHarianOK_${pratinjau?.tanggal}_${pratinjau?.shift}.${ext}`;
  const unduhJpg = async () => {
    setBusy(true);
    try {
      const cv = await ambilCanvas();
      const a = document.createElement("a");
      a.href = cv.toDataURL("image/jpeg", 0.92); a.download = namaFile("jpg"); a.click();
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
    if (!pratinjau) return;
    try { await navigator.clipboard.writeText(buatTeksLaporan(pratinjau, mask)); showToast("✓ Teks laporan disalin", C.s); }
    catch { showToast("Gagal menyalin teks", C.w); }
  };

  const unduhExcel = () => {
    const nilai = dlMode === "tanggal" ? dlTgl : dlMode === "bulan" ? dlBln : dlThn;
    const list = filterLaporan(laporanList, dlMode, nilai);
    if (!list.length) { showToast("Tidak ada laporan harian untuk " + nilai, C.w); return; }
    XLSX.writeFile(buatWorkbookLaporan(list), `LaporanHarianOK_${nilai}.xlsx`);
    showToast(`✓ Excel diunduh (${list.length} laporan)`, C.s);
  };

  const buka = (l: LaporanHarian) => {
    const ov: Record<string, any> = {};
    l.jadwal.forEach((r) => { ov["j:" + r.opId] = { status: r.status, catatan: r.catatan }; });
    l.rencana.forEach((r) => { ov["r:" + r.opId] = { alkes: r.alkes, keterangan: r.keterangan }; });
    setForm({
      ...formAwal(l.tanggal, l.shift), dari: l.dari, ke: l.ke, jumlahManual: l.jumlahOperasi, totalManual: l.totalPasien, cyto: l.adaCyto, ov,
      alat: l.alat || {}, alatLain: l.alatLain || "", alatCatatan: l.alatCatatan || "", instrumen: l.instrumen || "", implan: l.implan || "",
      alkesDipakai: l.alkesDipakai || "", alkesPersediaan: l.alkesPersediaan || "", pa: l.pa || "", stokLain: l.stokLain || "", stokCatatan: l.stokCatatan || "",
      kejadian: l.kejadian || "", uraian: l.uraian || "", lama: l.lama || "", baru: l.baru || "", kepala: l.kepala || "",
    });
    setPratinjau(null);
    window.scrollTo({ top: 0, behavior: "smooth" });
  };
  const hapus = async (l: LaporanHarian) => {
    if (!window.confirm(`Hapus laporan ${fmtTgl(l.tanggal)} ${SHIFT_LABEL[l.shift]}?`)) return;
    const res = await deleteFromSupa?.("kb_laporan_harian", l.id);
    if (res && res.ok === false) { showToast("⚠ Gagal menghapus: " + (res.error || "error"), C.d); return; }
    setLaporanList((x) => x.filter((r) => r.id !== l.id));
    showToast("✓ Laporan dihapus", C.s);
  };

  const ta = { ...iS, minHeight: 60, resize: "vertical", lineHeight: 1.6 } as React.CSSProperties;
  const jamSekarang = `${pad2(now.getHours())}:${pad2(now.getMinutes())}`;

  /* ══════════ PREVIEW ══════════ */
  if (pratinjau) {
    return (
      <div style={{ paddingBottom: 24 }}>
        <Card>
          <SH label="Preview Laporan Harian OK" />
          <div style={{ fontSize: 12, color: C.tL, marginBottom: 10 }}>Periksa dulu. Simpan, lalu unduh PDF/JPG atau salin teks untuk dibacakan.</div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <Btn onClick={simpan} disabled={busy}>💾 Simpan</Btn>
            <Btn outline onClick={unduhPdf} disabled={busy}>⬇ PDF</Btn>
            <Btn outline onClick={unduhJpg} disabled={busy}>⬇ JPG</Btn>
            <Btn outline onClick={salinTeks} disabled={busy}>📋 Salin teks</Btn>
            <Btn outline onClick={() => setPratinjau(null)} disabled={busy}>← Kembali edit</Btn>
          </div>
        </Card>
        <div style={{ overflowX: "auto", marginTop: 12, background: "#E5E7EB", padding: 8, borderRadius: 10 }}>
          <div ref={lembarRef} style={{ width: 794 }}><Lembar l={pratinjau} mask={mask} /></div>
        </div>
      </div>
    );
  }

  /* ══════════ FORM ══════════ */
  const riwayat = [...laporanList].sort(urutLaporanTurun).slice(0, 8);
  return (
    <div style={{ paddingBottom: 24 }}>
      <datalist id="kb-lh-staf">{(staff || []).filter((s: any) => s.type !== "surgeon" && s.type !== "anesthesiologist").map((s: any) => <option key={s.id} value={String(s.name || "").split(",")[0]} />)}</datalist>

      <Card>
        <SH label="Laporan Harian OK" />
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, flexWrap: "wrap", marginBottom: 10 }}>
          <div style={{ fontWeight: 800, fontSize: 15 }}>{tanggalPanjang(hariIni)}</div>
          <div style={{ fontWeight: 800, fontSize: 18, color: C.p }}>{jamSekarang} WIB</div>
        </div>
        <LF label="Tanggal laporan">
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            <input style={iS} type="date" value={form.tanggal} onChange={(e) => e.target.value && set({ tanggal: e.target.value })} />
            {form.tanggal !== hariIni && <Btn sm outline onClick={() => set({ tanggal: hariIni })}>Hari ini</Btn>}
          </div>
          <div style={{ fontSize: 12, color: C.tL, marginTop: 4 }}>{namaHari(form.tanggal)}</div>
        </LF>
        <LF label="Shift">
          <Pilih value={form.shift} onChange={(v) => set({ shift: v as ShiftKey })}
            opsi={SHIFT_URUT.map((s) => ({ v: s, l: s === "malam" && !malamAktif ? "Malam (nonaktif)" : SHIFT_LABEL[s].replace("Shift ", ""), off: s === "malam" && !malamAktif }))} />
        </LF>
        <div style={{ display: "flex", gap: 10 }}>
          <div style={{ flex: 1 }}><LF label="Dari shift">
            <select style={iS} value={dariEff} onChange={(e) => set({ dari: e.target.value as ShiftKey })}>
              {SHIFT_URUT.map((s) => <option key={s} value={s} disabled={s === "malam" && !malamAktif}>{SHIFT_LABEL[s]}</option>)}
            </select></LF></div>
          <div style={{ flex: 1 }}><LF label="Ke shift">
            <select style={iS} value={keEff} onChange={(e) => set({ ke: e.target.value as ShiftKey })}>
              {SHIFT_URUT.map((s) => <option key={s} value={s} disabled={s === "malam" && !malamAktif}>{SHIFT_LABEL[s]}</option>)}
            </select></LF></div>
        </div>
        {role === "admin" && <Toggle value={malamAktif} onChange={(v: any) => simpanCfg(typeof v === "boolean" ? v : !malamAktif)} label="Shift Malam aktif" sub="Matikan jika tidak ada dinas malam" />}
      </Card>

      <Card style={{ marginTop: 12 }}>
        <SH label="1. Ringkasan" />
        <div style={{ display: "flex", gap: 10 }}>
          <div style={{ flex: 1 }}><LF label="Jumlah operasi">
            <input style={iS} inputMode="numeric" value={jumlahOperasi} onChange={(e) => set({ jumlahManual: e.target.value.replace(/\D/g, "") })} />
            {form.jumlahManual !== "" && <button type="button" onClick={() => set({ jumlahManual: "" })} style={{ background: "none", border: "none", color: C.p, fontSize: 11, cursor: "pointer", padding: 0 }}>↺ otomatis</button>}
          </LF></div>
          <div style={{ flex: 1 }}><LF label="Total pasien (shift ini)">
            <input style={iS} inputMode="numeric" value={totalPasien} onChange={(e) => set({ totalManual: e.target.value.replace(/\D/g, "") })} />
            {form.totalManual !== "" && <button type="button" onClick={() => set({ totalManual: "" })} style={{ background: "none", border: "none", color: C.p, fontSize: 11, cursor: "pointer", padding: 0 }}>↺ otomatis</button>}
          </LF></div>
        </div>
        <LF label="Operasi darurat / cyto">
          <Pilih value={adaCyto} onChange={(v) => set({ cyto: v as "ya" | "tidak" })} opsi={[{ v: "ya", l: "Ya", c: C.d }, { v: "tidak", l: "Tidak" }]} />
        </LF>
      </Card>

      <Card style={{ marginTop: 12 }}>
        <SH label={`2. Jadwal operasi — ${SHIFT_LABEL[form.shift]}`} />
        <div style={{ fontSize: 12, color: C.tL, marginBottom: 8 }}>Di sistem ada {(ops || []).filter((o: any) => o && o.date === form.tanggal).length} operasi pada {fmtTgl(form.tanggal)}; {jadwal.length} masuk {SHIFT_LABEL[form.shift]}.</div>
        {!jadwal.length && <div style={{ fontSize: 13, color: C.tL }}>Tidak ada operasi terdaftar pada shift ini.</div>}
        {jadwal.map((r) => (
          <div key={r.opId} style={{ border: `1px solid ${C.b}`, borderRadius: 10, padding: 10, marginBottom: 8 }}>
            <div style={{ fontWeight: 700, fontSize: 13 }}>{r.jam || "--:--"} • {mask(r.pasien)}{r.cyto ? " • CYTO" : ""}</div>
            <div style={{ fontSize: 12, color: C.tL, margin: "2px 0 6px" }}>{r.jenis} | Operator: {r.operator || "-"} | Anestesi: {r.anestesi || "-"}</div>
            <div style={{ display: "flex", gap: 8 }}>
              <select style={{ ...iS, flex: 1 }} value={r.status} onChange={(e) => setOv("j:" + r.opId, { status: e.target.value })}>
                {STATUS_OPSI.map((s) => <option key={s} value={s}>{s}</option>)}
              </select>
              <input style={{ ...iS, flex: 2 }} placeholder="Catatan (opsional)" value={r.catatan} onChange={(e) => setOv("j:" + r.opId, { catatan: e.target.value })} />
            </div>
          </div>
        ))}
      </Card>

      <Card style={{ marginTop: 12 }}>
        <SH label={`3. Rencana operasi — ${penuh ? "Semua operasi besok" : SHIFT_LABEL[keEff]}, ${tanggalPanjang(tglRencana)}`} />
        {!rencana.length && <div style={{ fontSize: 13, color: C.tL }}>Belum ada operasi terdaftar pada shift tujuan.</div>}
        {rencana.map((r) => (
          <div key={r.opId} style={{ border: `1px solid ${C.b}`, borderRadius: 10, padding: 10, marginBottom: 8 }}>
            <div style={{ fontWeight: 700, fontSize: 13 }}>{r.jam || "--:--"} • {mask(r.pasien)}{r.umur ? ` (${r.umur})` : ""} • Kamar {r.kamar || "-"}</div>
            <div style={{ fontSize: 12, color: C.tL, margin: "2px 0 6px" }}>{r.jenis} | Operator: {r.operator || "-"} | Anestesi: {r.anestesi || "-"}</div>
            <div style={{ display: "flex", gap: 8 }}>
              <input style={{ ...iS, flex: 1 }} placeholder="Alkes yang dipakai" value={r.alkes} onChange={(e) => setOv("r:" + r.opId, { alkes: e.target.value })} />
              <input style={{ ...iS, flex: 1 }} placeholder="Keterangan" value={r.keterangan} onChange={(e) => setOv("r:" + r.opId, { keterangan: e.target.value })} />
            </div>
          </div>
        ))}
        <div style={{ fontSize: 12, color: C.tL, marginTop: 4 }}>{penuh ? "Kandidat dikosongkan: semua operasi besok sudah ada di Rencana." : `Kandidat (operasi besok, otomatis): ${kandidat.length} operasi — tampil di preview.`}</div>
      </Card>

      <Card style={{ marginTop: 12 }}>
        <SH label="4. Kondisi peralatan & kamar" />
        {ALAT_BUTIR.map((b) => (
          <LF key={b.k} label={b.l}>
            <Pilih value={form.alat[b.k] || ""} onChange={(v) => set({ alat: { ...form.alat, [b.k]: v } })} opsi={[{ v: "baik", l: "Baik", c: C.s }, { v: "tidak", l: "Tidak baik", c: C.d }]} />
          </LF>
        ))}
        <LF label="Lain-lain"><input style={iS} value={form.alatLain} onChange={(e) => set({ alatLain: e.target.value })} /></LF>
        <LF label="Catatan"><textarea style={ta} value={form.alatCatatan} onChange={(e) => set({ alatCatatan: e.target.value })} /></LF>
      </Card>

      <Card style={{ marginTop: 12 }}>
        <SH label="5. Stok & bahan" />
        <LF label="Instrumen"><Pilih value={form.instrumen} onChange={(v) => set({ instrumen: v })} opsi={STOK_OPSI.map((s) => ({ v: s, l: s }))} /></LF>
        <LF label="Implan / Bahan habis pakai"><Pilih value={form.implan} onChange={(v) => set({ implan: v })} opsi={STOK_OPSI.map((s) => ({ v: s, l: s }))} /></LF>
        <LF label="Alkes yang dipakai"><input style={iS} value={form.alkesDipakai} onChange={(e) => set({ alkesDipakai: e.target.value })} /></LF>
        <LF label="Alkes persediaan"><input style={iS} value={form.alkesPersediaan} onChange={(e) => set({ alkesPersediaan: e.target.value })} /></LF>
        <LF label="PA"><input style={iS} value={form.pa} onChange={(e) => set({ pa: e.target.value })} /></LF>
        <LF label="Lain-lain"><input style={iS} value={form.stokLain} onChange={(e) => set({ stokLain: e.target.value })} /></LF>
        <LF label="Catatan"><textarea style={ta} value={form.stokCatatan} onChange={(e) => set({ stokCatatan: e.target.value })} /></LF>
      </Card>

      <Card style={{ marginTop: 12 }}>
        <SH label="6. Kejadian / insiden / kendala" />
        <LF label="Ada kejadian?"><Pilih value={form.kejadian} onChange={(v) => set({ kejadian: v as "ya" | "tidak" })} opsi={[{ v: "ya", l: "Ya", c: C.d }, { v: "tidak", l: "Tidak" }]} /></LF>
        {form.kejadian === "ya" && <LF label="Jika ada, jelaskan"><textarea style={ta} value={form.uraian} onChange={(e) => set({ uraian: e.target.value })} /></LF>}
      </Card>

      <Card style={{ marginTop: 12 }}>
        <SH label="7. Serah terima" />
        <LF label="Petugas shift lama (nama jelas)"><input style={iS} list="kb-lh-staf" value={form.lama} onChange={(e) => set({ lama: e.target.value })} /></LF>
        <LF label="Petugas shift baru (nama jelas)"><input style={iS} list="kb-lh-staf" value={form.baru} onChange={(e) => set({ baru: e.target.value })} /></LF>
        <LF label="Mengetahui / Kepala Kamar Bedah (nama jelas)"><input style={iS} list="kb-lh-staf" value={form.kepala} onChange={(e) => set({ kepala: e.target.value })} /></LF>
        <Btn full onClick={keAtasPreview}>👁 Lihat preview laporan</Btn>
      </Card>

      <Card style={{ marginTop: 12 }}>
        <SH label="Riwayat & unduh Excel" />
        {!riwayat.length && <div style={{ fontSize: 13, color: C.tL }}>Belum ada laporan tersimpan.</div>}
        {riwayat.map((l) => (
          <div key={l.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8, padding: "6px 0", borderBottom: `1px solid ${C.b}` }}>
            <div style={{ fontSize: 13 }}><b>{fmtTgl(l.tanggal)}</b> • {SHIFT_LABEL[l.shift]} • {l.jumlahOperasi} operasi</div>
            <div style={{ display: "flex", gap: 6 }}>
              <Btn sm outline onClick={() => buka(l)}>Buka</Btn>
              {role === "admin" && <Btn sm outline color={C.d} onClick={() => hapus(l)}>Hapus</Btn>}
            </div>
          </div>
        ))}
        <div style={{ marginTop: 12 }}>
          <Pilih value={dlMode} onChange={(v) => setDlMode(v as "tanggal" | "bulan" | "tahun")} opsi={[{ v: "tanggal", l: "Per tanggal" }, { v: "bulan", l: "Per bulan" }, { v: "tahun", l: "Per tahun" }]} />
          <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
            {dlMode === "tanggal" && <input style={iS} type="date" value={dlTgl} onChange={(e) => setDlTgl(e.target.value)} />}
            {dlMode === "bulan" && <input style={iS} type="month" value={dlBln} onChange={(e) => setDlBln(e.target.value)} />}
            {dlMode === "tahun" && <input style={iS} inputMode="numeric" maxLength={4} value={dlThn} onChange={(e) => setDlThn(e.target.value.replace(/\D/g, ""))} />}
            <Btn onClick={unduhExcel}>⬇ Excel</Btn>
          </div>
        </div>
      </Card>
    </div>
  );
}
