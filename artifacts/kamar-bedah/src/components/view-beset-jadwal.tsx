/* src/components/view-beset-jadwal.tsx
   Tab "Beset Jadwal Perawat". Tabel nomor, nama perawat (dari tab Staf), kolom tanggal 1 bulan.
   Sel diisi tulisan (mis. P), sel berwarna sesuai perawat. Keterangan di bawah.
   Buka/Tutup beset dilindungi password admin. Disimpan ke kb_beset_jadwal (1 baris per bulan). */

import { useState, useEffect, useMemo, useRef, useCallback } from "react";
import { C, iS } from "../constants";
import { SUPA_CLIENT } from "../services/supabase";
import { Card, SH, Btn } from "./primitives";
import { isoLokal } from "../utils/laporan-harian";
import { escHtml } from "../utils/security";
import ExcelJS from "exceljs";

const TABEL = "kb_beset_jadwal";
const FONT = "'Times New Roman',Times,serif";
const LIBUR_WARNA = "#FCA5A5";
const PALET = ["#F87171", "#FACC15", "#4ADE80", "#60A5FA", "#C084FC", "#FB923C", "#2DD4BF", "#F472B6", "#A3E635", "#94A3B8", "#FB7185", "#38BDF8"];
const NAMA_BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

interface Beset {
  id: string; bulan: string; buka: boolean;
  sel: Record<string, string>; ket: Record<string, string>; warna: Record<string, string>;
  perawat: { id: string; name: string }[]; urutan: string[]; hilang: string[]; libur: number[]; updated_at?: string;
}
const kosong = (bulan: string): Beset => ({ id: "beset_" + bulan, bulan, buka: false, sel: {}, ket: {}, warna: {}, perawat: [], urutan: [], hilang: [], libur: [] });
const labelBulan = (b: string) => { const [y, m] = b.split("-").map(Number); return `${NAMA_BULAN[(m || 1) - 1]} ${y}`; };

export function ViewBesetJadwal({ staff, showToast, role, cekAdmin }: any) {
  const [bulan, setBulan] = useState<string>(isoLokal(new Date()).slice(0, 7));
  const [d, setD] = useState<Beset>(kosong(isoLokal(new Date()).slice(0, 7)));
  const [list, setList] = useState<Beset[]>([]);
  const [busy, setBusy] = useState(false);
  const [adaUbah, setAdaUbah] = useState(false);
  const [pinMode, setPinMode] = useState<null | "buka" | "tutup">(null);
  const [pin, setPin] = useState("");
  const dirty = useRef<Set<string>>(new Set());
  const struktur = useRef(false);

  const muat = useCallback(async (b: string) => {
    try {
      const r: any = await SUPA_CLIENT.from(TABEL).select("data").order("updated_at", { ascending: false });
      if (r.error) { showToast("⚠ Gagal memuat beset: " + (r.error.message || "error"), C.d); }
      const semua = ((r.data || []).map((x: any) => x.data).filter(Boolean)) as Beset[];
      setList(semua);
      const ada = semua.find((x) => x.bulan === b);
      setD(ada ? { ...kosong(b), ...ada } : kosong(b));
    } catch { setD(kosong(b)); }
    dirty.current.clear(); struktur.current = false; setAdaUbah(false);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  useEffect(() => { muat(bulan); }, [bulan, muat]);

  const sumber = useMemo(() => {
    const all = (staff || []).filter((s: any) => s && s.name);
    const nr = all.filter((s: any) => /nurse|perawat/i.test(String(s.type || "")));
    return nr.length ? nr : all.filter((s: any) => s.type !== "surgeon" && s.type !== "anesthesiologist");
  }, [staff]);

  const baris = useMemo(() => {
    const m = new Map<string, { id: string; name: string }>();
    sumber.forEach((s: any) => m.set(String(s.id), { id: String(s.id), name: String(s.name) }));
    (d.perawat || []).forEach((p) => { if (!m.has(p.id)) m.set(p.id, p); });
    const hil = new Set(d.hilang || []);
    const urut = d.urutan || [];
    const arr = Array.from(m.values()).filter((p) => !hil.has(p.id));
    arr.sort((a, b) => {
      const ia = urut.indexOf(a.id), ib = urut.indexOf(b.id);
      if (ia >= 0 && ib >= 0) return ia - ib;
      if (ia >= 0) return -1;
      if (ib >= 0) return 1;
      return a.name.localeCompare(b.name);
    });
    return arr;
  }, [sumber, d.perawat, d.urutan, d.hilang]);

  const warnaMap = useMemo(() => {
    const w: Record<string, string> = { ...(d.warna || {}) };
    const dipakai = new Set(Object.values(w));
    baris.forEach((b) => {
      if (!w[b.id]) {
        const c = PALET.find((x) => !dipakai.has(x)) || PALET[Object.keys(w).length % PALET.length];
        w[b.id] = c; dipakai.add(c);
      }
    });
    return w;
  }, [d.warna, baris]);

  const [y, m] = bulan.split("-").map(Number);
  const nHari = new Date(y, m, 0).getDate();
  const hari = Array.from({ length: nHari }, (_, i) => i + 1);

  const ubahSel = (id: string, h: number, v: string) => {
    if (!d.buka) return;
    const k = `${id}|${h}`;
    setD((p) => ({ ...p, sel: { ...p.sel, [k]: v } }));
    dirty.current.add("s|" + k); setAdaUbah(true);
  };
  const ubahKet = (id: string, v: string) => {
    if (!d.buka) return;
    setD((p) => ({ ...p, ket: { ...p.ket, [id]: v } }));
    dirty.current.add("k|" + id); setAdaUbah(true);
  };

  const pindahUrut = (id: string, posisi: number) => {
    if (role !== "admin") return;
    const ids = baris.map((b) => b.id).filter((x) => x !== id);
    ids.splice(posisi - 1, 0, id);
    setD((p) => ({ ...p, urutan: ids }));
    struktur.current = true; setAdaUbah(true);
  };
  const hapusPerawat = (b: { id: string; name: string }) => {
    if (role !== "admin") return;
    if (!window.confirm("Hapus " + b.name + " dari tabel beset " + labelBulan(bulan) + "? Isian tanggal dan keterangannya ikut dihapus.")) return;
    Object.keys(d.sel).forEach((k) => { if (k.startsWith(b.id + "|")) dirty.current.add("s|" + k); });
    dirty.current.add("k|" + b.id);
    setD((p) => {
      const sel = { ...p.sel }; Object.keys(sel).forEach((k) => { if (k.startsWith(b.id + "|")) delete sel[k]; });
      const ket = { ...p.ket }; delete ket[b.id];
      return { ...p, sel, ket, hilang: [...(p.hilang || []), b.id], urutan: (p.urutan || []).filter((x) => x !== b.id) };
    });
    struktur.current = true; setAdaUbah(true);
  };
  const toggleLibur = (h: number) => {
    if (role !== "admin") return;
    setD((p) => {
      const ada = (p.libur || []).includes(h);
      return { ...p, libur: ada ? (p.libur || []).filter((x) => x !== h) : [...(p.libur || []), h].sort((a, b) => a - b) };
    });
    struktur.current = true; setAdaUbah(true);
  };
  const kosongkanLibur = () => {
    if (role !== "admin") return;
    setD((p) => ({ ...p, libur: [] }));
    struktur.current = true; setAdaUbah(true);
  };
  const isLibur = (h: number) => (d.libur || []).includes(h);
  const tampilSemua = () => {
    if (role !== "admin") return;
    setD((p) => ({ ...p, hilang: [] }));
    struktur.current = true; setAdaUbah(true);
  };

  /* Simpan: ambil data terbaru dari server, timpa HANYA sel/keterangan yang diubah di layar ini. */
  const simpan = async (statusBuka?: boolean): Promise<boolean> => {
    setBusy(true);
    let ok = false;
    try {
      const r: any = await SUPA_CLIENT.from(TABEL).select("data").eq("id", d.id);
      const server: Beset = r.data?.[0]?.data ? { ...kosong(bulan), ...r.data[0].data } : kosong(bulan);
      const sel = { ...server.sel }; const ket = { ...server.ket };
      dirty.current.forEach((k) => {
        const kk = k.slice(2);
        if (k.startsWith("s|")) { const v = d.sel[kk]; if (v) sel[kk] = v; else delete sel[kk]; }
        else if (k.startsWith("k|")) { const v = d.ket[kk]; if (v) ket[kk] = v; else delete ket[kk]; }
      });
      const warna = { ...server.warna, ...Object.fromEntries(Object.entries(warnaMap).filter(([id]) => !server.warna?.[id])) };
      const nowIso = new Date().toISOString();
      const urutan = struktur.current ? baris.map((b) => b.id) : (server.urutan || []);
      const hilang = struktur.current ? (d.hilang || []) : (server.hilang || []);
      const libur = struktur.current ? (d.libur || []) : (server.libur || []);
      const data: Beset = { ...server, id: d.id, bulan, buka: statusBuka ?? server.buka, sel, ket, warna, perawat: baris, urutan, hilang, libur, updated_at: nowIso };
      const { error } = await SUPA_CLIENT.from(TABEL).upsert({ id: data.id, data, updated_at: nowIso }, { onConflict: "id" });
      if (error) showToast("⚠ Gagal menyimpan ke server: " + (error.message || "error"), C.d);
      else {
        ok = true;
        setD(data); dirty.current.clear(); struktur.current = false; setAdaUbah(false);
        setList((x) => [data, ...x.filter((r2) => r2.id !== data.id)]);
        showToast("✓ Beset jadwal tersimpan", C.s);
      }
    } catch (e: any) { showToast("⚠ Gagal menyimpan: " + (e?.message || "error"), C.d); }
    setBusy(false);
    return ok;
  };

  const konfirmasiPin = async () => {
    if (!pinMode) return;
    if (!cekAdmin) { showToast("⚠ Verifikasi password admin belum terpasang", C.d); return; }
    setBusy(true);
    let ok = false;
    try { ok = !!(await cekAdmin(pin)); } catch { ok = false; }
    if (!ok) { showToast("⚠ Password admin salah", C.d); setBusy(false); return; }
    const mode = pinMode;
    setPinMode(null); setPin(""); setBusy(false);
    await simpan(mode === "buka");
  };

  const hapus = async (l: Beset) => {
    if (!window.confirm(`Hapus beset ${labelBulan(l.bulan)}?`)) return;
    const { error } = await SUPA_CLIENT.from(TABEL).delete().eq("id", l.id);
    if (error) { showToast("⚠ Gagal menghapus: " + (error.message || "error"), C.d); return; }
    setList((x) => x.filter((r) => r.id !== l.id));
    if (l.id === d.id) setD(kosong(bulan));
    showToast("✓ Beset dihapus", C.s);
  };

  const unduhExcel = async () => {
    try {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet("Beset " + bulan);
      const NAMA = "Times New Roman";
      const tipis: Partial<ExcelJS.Borders> = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
      const sx = (v: string): string => (/^[=+\-@\t\r]/.test(v) ? "'" + v : v);
      const kol = 2 + nHari;
      ws.columns = [{ width: 5 }, { width: 26 }, ...hari.map(() => ({ width: 5 }))];
      const put = (r: number, c: number, v: any, o: { bold?: boolean; size?: number; abu?: boolean; warna?: string; kiri?: boolean; garis?: boolean } = {}) => {
        const cell = ws.getCell(r, c);
        cell.value = v;
        cell.font = { name: NAMA, size: o.size || 12, bold: !!o.bold };
        cell.alignment = { horizontal: o.kiri ? "left" : "center", vertical: "middle", wrapText: true };
        if (o.garis !== false) cell.border = tipis;
        if (o.abu) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEEEEEE" } };
        if (o.warna) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + o.warna.replace("#", "").toUpperCase() } };
      };
      ws.mergeCells(1, 1, 1, kol); put(1, 1, "BESET JADWAL PERAWAT KAMAR BEDAH", { bold: true, size: 14, garis: false });
      ws.mergeCells(2, 1, 2, kol); put(2, 1, labelBulan(bulan).toUpperCase(), { bold: true, size: 13, garis: false });
      put(4, 1, "No", { bold: true, abu: true }); put(4, 2, "Nama Perawat", { bold: true, abu: true });
      hari.forEach((x, j) => put(4, 3 + j, x, { bold: true, abu: true, warna: isLibur(x) ? LIBUR_WARNA : undefined }));
      baris.forEach((b, i) => {
        put(5 + i, 1, i + 1); put(5 + i, 2, sx(b.name), { kiri: true });
        hari.forEach((x, j) => { const v = d.sel[b.id + "|" + x] || ""; put(5 + i, 3 + j, sx(v), { bold: !!v, warna: v ? warnaMap[b.id] : isLibur(x) ? LIBUR_WARNA : undefined }); });
      });
      const k0 = 5 + baris.length + 1;
      put(k0, 1, "Keterangan", { bold: true, size: 13, kiri: true, garis: false });
      put(k0 + 1, 1, "No", { bold: true, abu: true }); put(k0 + 1, 2, "Nama", { bold: true, abu: true });
      ws.mergeCells(k0 + 1, 3, k0 + 1, kol); put(k0 + 1, 3, "Keterangan", { bold: true, abu: true });
      baris.forEach((b, i) => {
        const r = k0 + 2 + i;
        put(r, 1, i + 1); put(r, 2, sx(b.name), { kiri: true });
        ws.mergeCells(r, 3, r, kol); put(r, 3, sx(d.ket[b.id] || ""), { kiri: true });
      });
      const buf = await wb.xlsx.writeBuffer();
      const url = URL.createObjectURL(new Blob([buf as unknown as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
      const a = document.createElement("a");
      a.href = url; a.download = "BesetJadwal_" + bulan + ".xlsx"; a.click();
      setTimeout(() => URL.revokeObjectURL(url), 5000);
      showToast("✓ Excel diunduh", C.s);
    } catch (e: any) { showToast("Gagal membuat Excel: " + (e?.message || "error"), C.d); }
  };
  const gantiBulan = (b: string) => {
    if (!b) return;
    if (adaUbah && !window.confirm("Ada perubahan belum disimpan. Pindah bulan?")) return;
    setBulan(b);
  };

  const unduhWord = () => {
    const th = "border:1px solid #000;padding:3px;font-size:11pt;text-align:center;background:#EEEEEE";
    const td = "border:1px solid #000;padding:3px;font-size:11pt;";
    let h = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8"><title>Beset Jadwal</title><style>@page Section1{size:29.7cm 21cm;mso-page-orientation:landscape;margin:1.2cm}div.Section1{page:Section1}body{font-family:"Times New Roman",Times,serif}table{border-collapse:collapse}</style></head><body><div class="Section1">`;
    h += `<p style="text-align:center;font-weight:bold;font-size:15pt">BESET JADWAL PERAWAT KAMAR BEDAH<br/>${escHtml(labelBulan(bulan).toUpperCase())}</p>`;
    h += `<table style="width:100%"><tr><th style="${th}">No</th><th style="${th}">Nama Perawat</th>${hari.map((x) => `<th style="${th}${isLibur(x) ? ";background:#DC2626;color:#fff" : ""}">${x}</th>`).join("")}</tr>`;
    baris.forEach((b, i) => {
      h += `<tr><td style="${td}text-align:center">${i + 1}</td><td style="${td}">${escHtml(b.name)}</td>`;
      hari.forEach((x) => { const v = d.sel[`${b.id}|${x}`] || ""; h += `<td style="${td}text-align:center;${v ? `background:${warnaMap[b.id]};font-weight:bold` : isLibur(x) ? "background:#FCA5A5" : ""}">${escHtml(v)}</td>`; });
      h += "</tr>";
    });
    h += `</table><p style="font-weight:bold;font-size:13pt">Keterangan</p><table style="width:100%"><tr><th style="${th}">No</th><th style="${th}">Nama</th><th style="${th}">Keterangan</th></tr>`;
    baris.forEach((b, i) => { h += `<tr><td style="${td}text-align:center">${i + 1}</td><td style="${td}">${escHtml(b.name)}</td><td style="${td}">${escHtml(d.ket[b.id] || "")}</td></tr>`; });
    h += "</table></div></body></html>";
    const url = URL.createObjectURL(new Blob([h], { type: "application/msword" }));
    const a = document.createElement("a");
    a.href = url; a.download = `BesetJadwal_${bulan}.doc`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 5000);
    showToast("✓ Word diunduh", C.s);
  };

  const thS: React.CSSProperties = { border: "1px solid #111", padding: "6px 4px", fontSize: 15, fontWeight: 700, textAlign: "center", background: "#EEEEEE", fontFamily: FONT };
  const tdS: React.CSSProperties = { border: "1px solid #111", padding: "4px 6px", fontSize: 15, fontFamily: FONT };
  const riwayat = [...list].sort((a, b) => b.bulan.localeCompare(a.bulan)).slice(0, 12);

  return (
    <div style={{ paddingBottom: 24, fontFamily: FONT, fontSize: 15 }}>
      <Card>
        <SH label="Beset Jadwal Perawat" />
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          <input style={{ ...iS, width: "auto", flex: 1, minWidth: 150 }} type="month" value={bulan} onChange={(e) => gantiBulan(e.target.value)} />
          <Btn sm outline onClick={() => muat(bulan)} disabled={busy}>↻ Muat ulang</Btn>
        </div>
        <div style={{ fontWeight: 700, marginTop: 8 }}>{labelBulan(bulan)}</div>
        {d.buka
          ? <div style={{ fontSize: 15, fontWeight: 800, color: "#166534", background: "#DCFCE7", borderRadius: 10, padding: "10px 12px", marginTop: 8, textAlign: "center" }}>✓ BESET DIBUKA — perawat bisa mengisi</div>
          : <div style={{ fontSize: 15, fontWeight: 800, color: "#fff", background: "#DC2626", borderRadius: 10, padding: "10px 12px", marginTop: 8, textAlign: "center" }}>🔒 BESET DITUTUP — tabel tidak bisa diisi</div>}
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 10 }}>
          <Btn onClick={() => { setPin(""); setPinMode("buka"); }} disabled={busy || d.buka}>🔓 Beset Dibuka</Btn>
          <Btn outline color={C.d} onClick={() => { setPin(""); setPinMode("tutup"); }} disabled={busy || !d.buka}>🔒 Beset Ditutup</Btn>
        </div>
        {pinMode && (
          <div style={{ marginTop: 10, border: "1px solid " + C.b, borderRadius: 10, padding: 10 }}>
            <div style={{ fontWeight: 700, marginBottom: 6 }}>Masukkan password admin untuk {pinMode === "buka" ? "MEMBUKA" : "MENUTUP"} beset</div>
            <input style={iS} type="password" autoFocus value={pin} onChange={(e) => setPin(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") konfirmasiPin(); }} placeholder="Password admin" />
            <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
              <Btn sm onClick={konfirmasiPin} disabled={busy || !pin}>Konfirmasi</Btn>
              <Btn sm outline onClick={() => { setPinMode(null); setPin(""); }} disabled={busy}>Batal</Btn>
            </div>
          </div>
        )}
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 10 }}>
          <Btn onClick={() => simpan()} disabled={busy || !adaUbah}>💾 Simpan{adaUbah ? " (ada perubahan)" : ""}</Btn>
          <Btn outline onClick={unduhExcel} disabled={busy}>⬇ Excel</Btn>
          <Btn outline onClick={unduhWord} disabled={busy}>⬇ Word</Btn>
        </div>
      </Card>

      <Card style={{ marginTop: 12 }}>
        <SH label="Tabel beset" />
        {role === "admin" && (
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center", marginBottom: 8, fontSize: 13 }}>
            <span>Admin: ubah urutan lewat kotak No, hapus nama dengan ✕, lalu klik Simpan.</span>
            {(d.hilang || []).length > 0 && <Btn sm outline onClick={tampilSemua}>↺ Tampilkan perawat yang dihapus</Btn>}
          </div>
        )}
        {!baris.length && <div style={{ color: C.tL }}>Belum ada perawat di tab Staf.</div>}
        {role === "admin" && (
          <div style={{ border: "1px solid " + C.b, borderRadius: 10, padding: 8, marginBottom: 8 }}>
            <div style={{ fontWeight: 700, marginBottom: 6, fontSize: 14 }}>🗓 Pilih tanggal libur (blok satu hari penuh). Ketuk tanggal untuk memilih atau membatalkan, lalu klik Simpan.</div>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 4 }}>
              {hari.map((h) => (
                <button key={h} type="button" onClick={() => toggleLibur(h)}
                  style={{ width: 38, height: 34, borderRadius: 8, border: "1.5px solid " + (isLibur(h) ? "#DC2626" : C.b), background: isLibur(h) ? "#DC2626" : "#fff", color: isLibur(h) ? "#fff" : "#111", fontFamily: FONT, fontSize: 15, fontWeight: 700, cursor: "pointer" }}>{h}</button>
              ))}
            </div>
            {(d.libur || []).length > 0 && <div style={{ marginTop: 6 }}><Btn sm outline onClick={kosongkanLibur}>Hapus semua tanda libur</Btn></div>}
          </div>
        )}
        {(d.libur || []).length > 0 && <div style={{ fontSize: 13, marginBottom: 6 }}><span style={{ display: "inline-block", width: 12, height: 12, background: "#FCA5A5", border: "1px solid #111", marginRight: 6 }} />Kolom merah = tanggal libur (tetap bisa diisi untuk beset libur)</div>}
        <div style={{ overflowX: "auto" }}>
          <table style={{ borderCollapse: "collapse" }}>
            <thead>
              <tr>
                <th style={{ ...thS, minWidth: 36 }}>No</th>
                <th style={{ ...thS, minWidth: 170 }}>Nama Perawat</th>
                {hari.map((h) => <th key={h} style={{ ...thS, minWidth: 36, ...(isLibur(h) ? { background: "#DC2626", color: "#fff" } : {}) }}>{h}</th>)}
              </tr>
            </thead>
            <tbody>
              {baris.map((b, i) => (
                <tr key={b.id}>
                  <td style={{ ...tdS, textAlign: "center", padding: 2 }}>
                    {role === "admin"
                      ? <select value={i + 1} onChange={(e) => pindahUrut(b.id, Number(e.target.value))} style={{ fontFamily: FONT, fontSize: 15, fontWeight: 700, border: "none", background: "transparent" }}>{baris.map((_, n) => <option key={n} value={n + 1}>{n + 1}</option>)}</select>
                      : i + 1}
                  </td>
                  <td style={{ ...tdS, whiteSpace: "nowrap" }}>
                    <span style={{ display: "inline-block", width: 12, height: 12, background: warnaMap[b.id], border: "1px solid #111", marginRight: 6 }} />{b.name}{role === "admin" && <button type="button" aria-label="Hapus perawat" onClick={() => hapusPerawat(b)} style={{ border: "none", background: "transparent", color: C.d, fontSize: 16, cursor: "pointer", marginLeft: 6 }}>✕</button>}
                  </td>
                  {hari.map((h) => {
                    const v = d.sel[`${b.id}|${h}`] || "";
                    return (
                      <td key={h} style={{ ...tdS, padding: 0, background: v ? warnaMap[b.id] : isLibur(h) ? LIBUR_WARNA : "#fff" }}>
                        <input value={v} disabled={!d.buka} maxLength={6} onChange={(e) => ubahSel(b.id, h, e.target.value.toUpperCase())}
                          style={{ width: 36, border: "none", background: "transparent", textAlign: "center", fontFamily: FONT, fontSize: 15, fontWeight: 700, padding: "6px 0", color: "#111", boxSizing: "border-box" }} />
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>

      <Card style={{ marginTop: 12 }}>
        <SH label="Keterangan" />
        <div style={{ overflowX: "auto" }}>
          <table style={{ borderCollapse: "collapse", width: "100%", minWidth: 480 }}>
            <thead><tr><th style={{ ...thS, width: 40 }}>No</th><th style={{ ...thS, width: 200 }}>Nama</th><th style={thS}>Keterangan</th></tr></thead>
            <tbody>
              {baris.map((b, i) => (
                <tr key={b.id}>
                  <td style={{ ...tdS, textAlign: "center" }}>{i + 1}</td>
                  <td style={{ ...tdS, whiteSpace: "nowrap" }}>{b.name}</td>
                  <td style={{ ...tdS, padding: 2 }}>
                    <input value={d.ket[b.id] || ""} disabled={!d.buka} onChange={(e) => ubahKet(b.id, e.target.value)}
                      style={{ ...iS, border: "none", width: "100%", boxSizing: "border-box", fontFamily: FONT, fontSize: 15 }} />
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>

      <Card style={{ marginTop: 12 }}>
        <SH label="Riwayat bulanan" />
        {!riwayat.length && <div style={{ color: C.tL }}>Belum ada beset tersimpan.</div>}
        {riwayat.map((l) => (
          <div key={l.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8, padding: "6px 0", borderBottom: `1px solid ${C.b}` }}>
            <div><b>{labelBulan(l.bulan)}</b> • {l.buka ? "Dibuka" : "Ditutup"}</div>
            <div style={{ display: "flex", gap: 6 }}>
              <Btn sm outline onClick={() => { gantiBulan(l.bulan); window.scrollTo({ top: 0, behavior: "smooth" }); }}>Buka</Btn>
              {role === "admin" && <Btn sm outline color={C.d} onClick={() => hapus(l)}>Hapus</Btn>}
            </div>
          </div>
        ))}
      </Card>
    </div>
  );
}
