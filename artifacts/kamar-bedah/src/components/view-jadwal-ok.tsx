/* src/components/view-jadwal-ok.tsx
   Tab "Jadwal Operasi dan Peserta Kamar Bedah". Perawat jaga diambil dari
   Generator (kb_roster_gen), baris tabel dari daftar operasi hari itu, semua
   bisa diedit. Disimpan ke Supabase (kb_jadwal_ok, realtime) dan Excel harian
   ke Dropbox. Unduh per hari / bulan / tahun. Font Times New Roman 14. */

import { useState, useEffect, useMemo, useRef } from "react";
import { C, iS } from "../constants";
import { maskName } from "../utils";
import { SUPA_CLIENT, AmbilLogMedis } from "../services/supabase";
import { dropboxUploadBytes } from "../services/dropbox";
import { Card, SH, Btn } from "./primitives";
import { isoLokal, fmtTgl, namaHari, tanggalPanjang } from "../utils/laporan-harian";
import type { JadwalOK, BarisOK, PerawatJaga } from "../utils/jadwal-ok";
import { barisKosong, susunBarisDariOps, perawatDariGenerator, buatWorkbookJadwalOK, workbookKeBlob, workbookKeBase64, filterJadwalOK } from "../utils/jadwal-ok";
import { perawatDariUpload } from "../utils/jadwal-perawat";

const FONT = "'Times New Roman',Times,serif";
const KOSONG: PerawatJaga = { pagi: "", pg: "", sore: "", sg: "", cuti: "" };
const KOLOM: { k: keyof BarisOK; l: string; w: number }[] = [
  { k: "jam", l: "JAM", w: 80 }, { k: "bgsl", l: "BGSL", w: 110 }, { k: "ok", l: "OK", w: 130 },
  { k: "dokter", l: "DOKTER/ANESTESI", w: 230 }, { k: "pasien", l: "NAMA PASIEN", w: 200 }, { k: "umur", l: "UMUR", w: 80 },
  { k: "jenis", l: "JENIS OPERASI", w: 220 }, { k: "asisten1", l: "ASISTEN I", w: 160 }, { k: "asisten2", l: "ASISTEN II", w: 160 },
  { k: "onloop", l: "ONLOOP", w: 160 }, { k: "rr", l: "RR", w: 130 },
];
const PERAWAT_FIELD: { k: keyof PerawatJaga; l: string }[] = [
  { k: "pagi", l: "PERAWAT JAGA PAGI" }, { k: "pg", l: "PG" }, { k: "sore", l: "SORE" }, { k: "sg", l: "SG" }, { k: "cuti", l: "CUTI (T)" },
];

interface FormJ { perawat: PerawatJaga; baris: BarisOK[]; }

export function ViewJadwalOK({ ops, showToast, role, privacyMode, upsertOneToSupa, deleteFromSupa, dbxCfg }: any) {
  const hariIni = isoLokal(new Date());
  const [tanggal, setTanggal] = useState<string>(hariIni);
  const [list, setList] = useState<JadwalOK[]>([]);
  const [loaded, setLoaded] = useState(false);
  const [gen, setGen] = useState<any>(null);
  const [genStatus, setGenStatus] = useState<"muat" | "ada" | "kosong">("muat");
  const [form, setForm] = useState<FormJ>({ perawat: KOSONG, baris: [] });
  const [dirty, setDirty] = useState(false);
  const [lihat, setLihat] = useState<JadwalOK | null>(null); /* lihatHasil */
  const lembarRef = useRef<HTMLDivElement>(null);
  const [busy, setBusy] = useState(false);
  const [dlMode, setDlMode] = useState<"tanggal" | "bulan" | "tahun">("bulan");
  const [dlTgl, setDlTgl] = useState(hariIni);
  const [dlBln, setDlBln] = useState(hariIni.slice(0, 7));
  const [dlThn, setDlThn] = useState(hariIni.slice(0, 4));

  /* muat awal + realtime */
  useEffect(() => {
    let batal = false;
    (async () => {
      try {
        const r = await SUPA_CLIENT.from("kb_jadwal_ok").select("data").order("updated_at", { ascending: false });
        if (!batal && r.data) setList(r.data.map((x: any) => x.data).filter(Boolean));
      } catch { /* offline — mulai dari kosong */ }
      if (!batal) setLoaded(true);
    })();
    const ch = SUPA_CLIENT.channel("kb-jadwal-ok-rt")
      .on("postgres_changes", { event: "*", schema: "public", table: "kb_jadwal_ok" }, (p: any) => {
        if (p.eventType === "DELETE") { const id = p.old?.id; if (id) setList((x) => x.filter((r) => r.id !== id)); }
        else if (p.new?.data) { const d = p.new.data as JadwalOK; setList((x) => [d, ...x.filter((r) => r.id !== d.id)]); }
      }).subscribe();
    return () => { batal = true; SUPA_CLIENT.removeChannel(ch); };
  }, []);

  /* jadwal Generator bulan yang dipilih */
  const bulanKey = tanggal.slice(0, 7);
  useEffect(() => {
    let batal = false;
    setGenStatus("muat");
    (async () => {
      try {
        const [y, m] = bulanKey.split("-");
        const rp: any = await SUPA_CLIENT.from("kb_jadwal_perawat").select("data").eq("id", `jp_${y}_${m}`);
        const up = rp.data?.[0]?.data;
        if (up) { if (!batal) { setGen({ __upload: up }); setGenStatus("ada"); } return; }
        const r = await SUPA_CLIENT.from("kb_roster_gen").select("*").eq("id", `roster_${y}_${m}`);
        const row: any = r.data?.[0];
        const d = row ? (row.data ?? row) : null;
        if (!batal) { setGen(d); setGenStatus(d ? "ada" : "kosong"); }
      } catch { if (!batal) { setGen(null); setGenStatus("kosong"); } }
    })();
    return () => { batal = true; };
  }, [bulanKey]);

  const tersimpan = useMemo(() => list.find((x) => x.id === tanggal) || null, [list, tanggal]);
  const perawatAuto = useMemo(() => (gen?.__upload ? perawatDariUpload(gen.__upload, tanggal) : perawatDariGenerator(gen, tanggal)) || KOSONG, [gen, tanggal]);
  const barisAuto = useMemo(() => susunBarisDariOps(ops, tanggal), [ops, tanggal]);

  /* isi form: dari salinan tersimpan, kalau belum ada dari Generator + jadwal operasi */
  useEffect(() => {
    if (!loaded || dirty) return;
    if (tersimpan) setForm({ perawat: { ...KOSONG, ...tersimpan.perawat }, baris: tersimpan.baris || [] });
    else setForm({ perawat: perawatAuto, baris: barisAuto });
  }, [loaded, dirty, tersimpan, perawatAuto, barisAuto]);

  const gantiTanggal = (t: string) => {
    if (!t || t === tanggal) return;
    if (dirty && !window.confirm("Ada perubahan yang belum disimpan. Pindah tanggal dan buang perubahan?")) return;
    setDirty(false);
    setTanggal(t);
  };
  const ubahPerawat = (k: keyof PerawatJaga, v: string) => { setForm((p) => ({ ...p, perawat: { ...p.perawat, [k]: v } })); setDirty(true); };
  const ubahBaris = (id: string, k: keyof BarisOK, v: string) => { setForm((p) => ({ ...p, baris: p.baris.map((b) => (b.id === id ? { ...b, [k]: v } : b)) })); setDirty(true); };
  const tambahBaris = () => { setForm((p) => ({ ...p, baris: [...p.baris, barisKosong()] })); setDirty(true); };
  const hapusBaris = (id: string) => { setForm((p) => ({ ...p, baris: p.baris.filter((b) => b.id !== id) })); setDirty(true); };
  const muatUlang = () => {
    if (!window.confirm("Ganti isi dengan data terbaru dari Generator dan jadwal operasi? Perubahan manual akan hilang.")) return;
    setForm({ perawat: perawatAuto, baris: barisAuto });
    setDirty(true);
  };

  const bangun = (): JadwalOK => ({
    id: tanggal, tanggal, hari: namaHari(tanggal), perawat: form.perawat, baris: form.baris,
    operator: role === "admin" ? "Admin" : "Perawat", updated_at: new Date().toISOString(),
  });

  const simpan = async () => {
    setBusy(true);
    const data = bangun();
    const res = await upsertOneToSupa("kb_jadwal_ok", data);
    if (!res?.ok) {
      showToast(`⚠ Gagal menyimpan ke server: ${res?.error || "error"}. Perubahan masih di layar ini.`, C.d);
      setBusy(false);
      return;
    }
    setList((x) => [data, ...x.filter((r) => r.id !== data.id)]);
    setDirty(false);
    showToast("✓ Jadwal operasi tersimpan", C.s);
    try {
      const folder = (dbxCfg?.path || "/kamarbedah/").replace(/[^/]+$/, "");
      const buf = await buatWorkbookJadwalOK([data]).xlsx.writeBuffer();
      const d = await dropboxUploadBytes(`${folder}JadwalOK_${data.tanggal}.xlsx`, buf as unknown as ArrayBuffer);
      if (!d.ok) showToast("⚠ Excel ke Dropbox gagal: " + d.msg, C.w);
    } catch { showToast("⚠ Excel ke Dropbox gagal", C.w); }
    AmbilLogMedis("JADWAL_OK", data.operator, "SIMPAN_JADWAL_OK", `Jadwal operasi ${data.tanggal} disimpan`, { id: data.id, jumlahBaris: data.baris.length });
    setBusy(false);
  };

  const hapus = async (l: JadwalOK) => {
    if (!window.confirm(`Hapus jadwal ${fmtTgl(l.tanggal)}?`)) return;
    const res = await deleteFromSupa?.("kb_jadwal_ok", l.id);
    if (res && res.ok === false) { showToast("⚠ Gagal menghapus: " + (res.error || "error"), C.d); return; }
    setList((x) => x.filter((r) => r.id !== l.id));
    showToast("✓ Jadwal dihapus", C.s);
  };

  const unduh = async () => {
    const nilai = dlMode === "tanggal" ? dlTgl : dlMode === "bulan" ? dlBln : dlThn;
    const hasil = dlMode === "tanggal" && nilai === tanggal ? [bangun()] : filterJadwalOK(list, dlMode, nilai);
    if (!hasil.length) { showToast("Tidak ada jadwal tersimpan untuk " + nilai, C.w); return; }
    setBusy(true);
    try {
      const blob = await workbookKeBlob(buatWorkbookJadwalOK(hasil));
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = `JadwalOK_${nilai}.xlsx`; a.click();
      setTimeout(() => URL.revokeObjectURL(url), 5000);
      showToast(`✓ Excel diunduh (${hasil.length} hari)`, C.s);
    } catch (e: any) { showToast("Gagal membuat Excel: " + (e?.message || "error"), C.d); }
    setBusy(false);
  };

  const gayaInput: React.CSSProperties = { ...iS, fontFamily: FONT, fontSize: 14, width: "100%", boxSizing: "border-box", padding: "6px 8px" };
  const th: React.CSSProperties = { border: "1px solid #111", padding: "6px 8px", fontSize: 14, fontWeight: 700, textAlign: "center", background: "#EEEEEE", whiteSpace: "nowrap" };
  const td: React.CSSProperties = { border: "1px solid #111", padding: 2, verticalAlign: "middle" };
  const riwayat = [...list].sort((a, b) => b.tanggal.localeCompare(a.tanggal)).slice(0, 8);
  const segmen = (v: "tanggal" | "bulan" | "tahun", l: string) => (
    <button key={v} type="button" onClick={() => setDlMode(v)}
      style={{ flex: 1, padding: "8px 10px", borderRadius: 10, border: `1.5px solid ${dlMode === v ? C.p : C.b}`, background: dlMode === v ? C.p : "#FAFBFD", color: dlMode === v ? "#fff" : C.t, fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: FONT }}>{l}</button>
  );

  /* ── hasil jadi (Buka dari Riwayat) ── */
  const cetak = () => {
    const el = lembarRef.current;
    if (!el) { showToast("Lembar belum siap dicetak", C.w); return; }
    const ifr = document.createElement("iframe");
    ifr.style.cssText = "position:fixed;right:0;bottom:0;width:0;height:0;border:0";
    document.body.appendChild(ifr);
    const w = ifr.contentWindow;
    if (!w) { ifr.remove(); return; }
    w.document.open();
    w.document.write("<!doctype html><html><head><meta charset='utf-8'><title>Jadwal Operasi Kamar Bedah</title><style>@page{size:A4 landscape;margin:6mm}body{margin:0}.kb-lembar{zoom:0.74}</style></head><body>" + el.innerHTML + "</body></html>");
    w.document.close();
    setTimeout(() => { w.focus(); w.print(); setTimeout(() => ifr.remove(), 3000); }, 400);
  };
  if (lihat) {
    const L = lihat;
    const hth: React.CSSProperties = { border: "1px solid #111", padding: "6px 8px", fontSize: 14, fontWeight: 700, textAlign: "center", background: "#EEEEEE", fontFamily: FONT };
    const htd: React.CSSProperties = { border: "1px solid #111", padding: "6px 8px", fontSize: 14, verticalAlign: "top", fontFamily: FONT, minHeight: 28, height: 30 };
    const jumlah = Math.max(10, L.baris.length);
    const p = { ...KOSONG, ...(L.perawat || {}) };
    const barisJaga: [string, string][] = [["PERAWAT JAGA PAGI", p.pagi], ["PG", p.pg], ["SORE", p.sore], ["SG", p.sg], ["CUTI (T)", p.cuti]];
    return (
      <div style={{ paddingBottom: 24, fontFamily: FONT, fontSize: 14 }}>
        <Card>
          <SH label="Hasil Jadwal Operasi" />
          <div style={{ fontWeight: 700, marginBottom: 10 }}>{tanggalPanjang(L.tanggal)}</div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <Btn onClick={cetak}>🖨 Print</Btn>
            <Btn outline onClick={() => { setLihat(null); gantiTanggal(L.tanggal); window.scrollTo({ top: 0, behavior: "smooth" }); }}>✎ Edit</Btn>
            <Btn outline onClick={() => setLihat(null)}>✕ Tutup</Btn>
          </div>
        </Card>
        <div style={{ overflowX: "auto", marginTop: 12, background: "#E5E7EB", padding: 8, borderRadius: 10 }}>
          <div ref={lembarRef} style={{ width: 1400 }}>
            <div className="kb-lembar" style={{ width: 1400, boxSizing: "border-box", padding: 28, background: "#fff", color: "#111", fontFamily: "'Times New Roman',Times,serif", fontSize: 14 }}>
              <div style={{ textAlign: "center", fontWeight: 800, fontSize: 18 }}>JADWAL OPERASI KAMAR BEDAH</div>
              <div style={{ textAlign: "center", fontWeight: 700, marginBottom: 10 }}>{tanggalPanjang(L.tanggal)}</div>
              {barisJaga.map(([a, b]) => <div key={a} style={{ marginBottom: 3 }}><b>{a} :</b> {b}</div>)}
              <table style={{ width: "100%", borderCollapse: "collapse", marginTop: 12 }}>
                <thead><tr><th style={hth}>TANGGAL</th>{KOLOM.map((k) => <th key={k.k} style={hth}>{k.l}</th>)}</tr></thead>
                <tbody>
                  {Array.from({ length: jumlah }).map((_, i) => {
                    const b: any = L.baris[i];
                    return (
                      <tr key={i}>
                        <td style={htd}>{b ? fmtTgl(L.tanggal) : ""}</td>
                        {KOLOM.map((k) => <td key={k.k} style={htd}>{b ? (k.k === "pasien" ? maskName(String(b.pasien || ""), privacyMode) : String(b[k.k] ?? "")) : ""}</td>)}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    );
  }
  return (
    <div style={{ paddingBottom: 24, fontFamily: FONT, fontSize: 14 }}>
      <Card>
        <SH label="Jadwal Operasi & Peserta Kamar Bedah" />
        <div style={{ textAlign: "center", fontWeight: 800, fontSize: 18, marginBottom: 10 }}>JADWAL OPERASI KAMAR BEDAH</div>
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", marginBottom: 6 }}>
          <input style={{ ...gayaInput, width: "auto", flex: 1, minWidth: 150 }} type="date" value={tanggal} onChange={(e) => gantiTanggal(e.target.value)} />
          {tanggal !== hariIni && <Btn sm outline onClick={() => gantiTanggal(hariIni)}>Hari ini</Btn>}
        </div>
        <div style={{ fontWeight: 700 }}>{tanggalPanjang(tanggal)}</div>
        <div style={{ fontSize: 12, color: C.tL, marginTop: 2 }}>
          {tersimpan ? "Sudah tersimpan" : "Belum disimpan"}{dirty ? " • ada perubahan belum disimpan" : ""}
        </div>
      </Card>

      <Card style={{ marginTop: 12 }}>
        <SH label="Perawat jaga" />
        {genStatus === "kosong" && <div style={{ fontSize: 12, color: C.w, marginBottom: 8 }}>Jadwal perawat bulan ini belum diupload (tab Staf) atau disimpan di Generator, isi nama perawat secara manual.</div>}
        {PERAWAT_FIELD.map((f) => (
          <div key={f.k} style={{ marginBottom: 8 }}>
            <div style={{ fontWeight: 700, marginBottom: 3 }}>{f.l} :</div>
            <input style={gayaInput} value={form.perawat[f.k]} onChange={(e) => ubahPerawat(f.k, e.target.value)} />
          </div>
        ))}
      </Card>

      <Card style={{ marginTop: 12 }}>
        <SH label="Jadwal operasi" />
        <div style={{ overflowX: "auto" }}>
          <table style={{ borderCollapse: "collapse", minWidth: 1700 }}>
            <thead>
              <tr>
                <th style={{ ...th, minWidth: 110 }}>TANGGAL</th>
                {KOLOM.map((k) => <th key={k.k} style={{ ...th, minWidth: k.w }}>{k.l}</th>)}
                <th style={{ ...th, minWidth: 44 }} />
              </tr>
            </thead>
            <tbody>
              {form.baris.map((b) => (
                <tr key={b.id}>
                  <td style={{ ...td, padding: "6px 8px", textAlign: "center" }}>{fmtTgl(tanggal)}</td>
                  {KOLOM.map((k) => (
                    <td key={k.k} style={td}>
                      {k.k === "pasien" && privacyMode
                        ? <div style={{ padding: "6px 8px" }}>{maskName(String(b.pasien || ""), privacyMode)}</div>
                        : <input style={{ ...gayaInput, border: "none", background: "transparent" }} value={b[k.k]} onChange={(e) => ubahBaris(b.id, k.k, e.target.value)} />}
                    </td>
                  ))}
                  <td style={{ ...td, textAlign: "center" }}>
                    <button type="button" aria-label="Hapus baris" onClick={() => hapusBaris(b.id)} style={{ border: "none", background: "transparent", color: C.d, fontSize: 16, cursor: "pointer" }}>✕</button>
                  </td>
                </tr>
              ))}
              {!form.baris.length && (
                <tr><td colSpan={KOLOM.length + 2} style={{ ...td, padding: 14, textAlign: "center", color: C.tL }}>Tidak ada operasi terdaftar pada tanggal ini. Tambahkan baris manual kalau perlu.</td></tr>
              )}
            </tbody>
          </table>
        </div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 12 }}>
          <Btn outline onClick={tambahBaris} disabled={busy}>+ Tambah baris</Btn>
          <Btn outline onClick={muatUlang} disabled={busy}>↺ Muat ulang dari jadwal</Btn>
          <Btn onClick={simpan} disabled={busy}>💾 Simpan</Btn>
        </div>
      </Card>

      <Card style={{ marginTop: 12 }}>
        <SH label="Riwayat & unduh Excel" />
        {!riwayat.length && <div style={{ color: C.tL }}>Belum ada jadwal tersimpan.</div>}
        {riwayat.map((l) => (
          <div key={l.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8, padding: "6px 0", borderBottom: `1px solid ${C.b}` }}>
            <div><b>{fmtTgl(l.tanggal)}</b> • {l.hari} • {l.baris.length} baris</div>
            <div style={{ display: "flex", gap: 6 }}>
              <Btn sm outline onClick={() => { setLihat(l); window.scrollTo({ top: 0, behavior: "smooth" }); }}>Buka</Btn>
              <Btn sm outline onClick={() => { setLihat(l); window.scrollTo({ top: 0 }); setTimeout(cetak, 700); }}>🖨 Print</Btn>
              {role === "admin" && <Btn sm outline color={C.d} onClick={() => hapus(l)}>Hapus</Btn>}
            </div>
          </div>
        ))}
        <div style={{ marginTop: 12 }}>
          <div style={{ display: "flex", gap: 6 }}>{segmen("tanggal", "Per hari")}{segmen("bulan", "Per bulan")}{segmen("tahun", "Per tahun")}</div>
          <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
            {dlMode === "tanggal" && <input style={gayaInput} type="date" value={dlTgl} onChange={(e) => setDlTgl(e.target.value)} />}
            {dlMode === "bulan" && <input style={gayaInput} type="month" value={dlBln} onChange={(e) => setDlBln(e.target.value)} />}
            {dlMode === "tahun" && <input style={gayaInput} inputMode="numeric" maxLength={4} value={dlThn} onChange={(e) => setDlThn(e.target.value.replace(/\D/g, ""))} />}
            <Btn onClick={unduh} disabled={busy}>⬇ Excel</Btn>
          </div>
        </div>
      </Card>
    </div>
  );
}
