/* src/components/view-mutu-ok.tsx
   Tab "Pemantauan Kamar Bedah" (grup MUTU kamar bedah). Satu pasien = satu baris, otomatis dari
   daftar operasi (tab Jadwal). Isian perawat tersimpan otomatis ke Supabase (kb_mutu_ok, realtime).
   Tombol Simpan membuat Excel bulanan (template mutu.xlsx) ke Dropbox; Unduh mengunduh Excel yang sama. */
import { useState, useEffect, useMemo, useRef } from "react";
import { C } from "../constants";
import { maskName } from "../utils";
import { SUPA_CLIENT, AmbilLogMedis } from "../services/supabase";
import { dropboxUploadBytes } from "../services/dropbox";
import { isoLokal } from "../utils/laporan-harian";
import { workbookKeBlob } from "../utils/jadwal-ok";
import type { MutuRow, MutuItem, GrupForm, FieldDef } from "../utils/mutu-ok";
import { GRUP_FORM, NAMA_BULAN, adaIsi, adaIsiApapun, baruRow, buatWorkbookMutu, gabung, idDariOp, nilaiBaris } from "../utils/mutu-ok";

const fmtIso = (s: string): string => (s && s.length >= 10 ? `${s.slice(8, 10)}-${s.slice(5, 7)}-${s.slice(0, 4)}` : s);
const namaBulan = (b: string): string => { const n = parseInt(b.slice(5, 7), 10); const m = NAMA_BULAN[n - 1] || ""; return m.charAt(0) + m.slice(1).toLowerCase() + " " + b.slice(0, 4); };
const KEL = "#374151";

/* kepala tabel preview: [teks, colSpan, rowSpan] */
const H1: [string, number, number][] = [["NO", 1, 3], ["TGL", 1, 3], ["KATEGORI", 4, 1], ["NO RM", 1, 3], ["SKP", 3, 1], ["PEMBIUSAN", 4, 1], ["TIME OUT", 3, 1], ["NASIONAL", 3, 1], ["PEMBEDAHAN", 7, 1], ["PONEK", 4, 1], ["PPRA", 2, 1], ["UNIT", 2, 1]];
const H2: [string, number, number][] = [["K", 1, 2], ["S", 1, 2], ["B", 1, 2], ["KH", 1, 2], ["Visite", 2, 1], ["Anestesi", 1, 2], ["Konversi", 2, 1], ["Sat <92%", 2, 1], ["Time Out", 2, 1], ["Nama", 1, 2], ["Tunda >1 jam", 2, 1], ["Menit", 1, 2], ["Site Marking", 2, 1], ["Assesmen", 2, 1], ["Dokter Bedah", 1, 2], ["Salah Dx", 2, 1], ["SC <30 mnt", 2, 1], ["Daftar", 1, 2], ["Insisi", 1, 2], ["FAM tanpa AB", 2, 1], ["SSC", 2, 1]];
const JML_PASANG = 11;

export function ViewMutuOK({ ops, showToast, role, privacyMode, upsertOneToSupa, dbxCfg }: any) {
  const hariIni = isoLokal(new Date());
  const [tanggal, setTanggal] = useState<string>(hariIni);
  const [list, setList] = useState<MutuRow[]>([]);
  const [buka, setBuka] = useState<string | null>(null);
  const [preview, setPreview] = useState(false);
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState<"" | "menyimpan" | "tersimpan" | "gagal">("");
  const [semuaBulan, setSemuaBulan] = useState(false);
  const listRef = useRef<MutuRow[]>([]);
  const timers = useRef<Record<string, any>>({});
  const pending = useRef<Set<string>>(new Set());
  const bulan = tanggal.slice(0, 7);

  /* muat awal + realtime */
  useEffect(() => {
    let batal = false;
    (async () => {
      try {
        const r = await SUPA_CLIENT.from("kb_mutu_ok").select("data");
        if (!batal && r.data) { const d = r.data.map((x: any) => x.data).filter(Boolean) as MutuRow[]; listRef.current = d; setList(d); }
      } catch { /* offline — mulai dari kosong */ }
    })();
    const ch = SUPA_CLIENT.channel("kb-mutu-ok-rt")
      .on("postgres_changes", { event: "*", schema: "public", table: "kb_mutu_ok" }, (p: any) => {
        if (p.eventType === "DELETE") return;
        const d = p.new?.data as MutuRow | undefined;
        if (!d?.id) return;
        const lokal = listRef.current.find((r) => r.id === d.id);
        if (lokal && String(lokal.updated_at || "") >= String(d.updated_at || "")) return; /* abaikan gema simpanan sendiri */
        const next = [d, ...listRef.current.filter((r) => r.id !== d.id)];
        listRef.current = next; setList(next);
      }).subscribe();
    return () => {
      batal = true; SUPA_CLIENT.removeChannel(ch);
      pending.current.forEach((id) => { clearTimeout(timers.current[id]); const cur = listRef.current.find((r) => r.id === id); if (cur) upsertOneToSupa("kb_mutu_ok", cur); });
    };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const kirim = async (id: string): Promise<boolean> => {
    pending.current.delete(id);
    const cur = listRef.current.find((r) => r.id === id);
    if (!cur) return true;
    const res = await upsertOneToSupa("kb_mutu_ok", cur);
    if (!res?.ok) { setStatus("gagal"); showToast(`Gagal menyimpan ke server: ${res?.error || "error"}. Isian masih di layar ini.`, C.d); return false; }
    if (pending.current.size === 0) setStatus("tersimpan");
    return true;
  };

  const ubah = (op: any, patch: (r: MutuRow) => MutuRow) => {
    const id = idDariOp(op.id);
    const lama = listRef.current.find((r) => r.id === id) || baruRow(op);
    const baru: MutuRow = { ...patch(lama), tanggal: String(op.date || lama.tanggal), updated_at: new Date().toISOString() };
    const next = [baru, ...listRef.current.filter((r) => r.id !== id)];
    listRef.current = next; setList(next);
    setStatus("menyimpan");
    pending.current.add(id);
    clearTimeout(timers.current[id]);
    timers.current[id] = setTimeout(() => { kirim(id); }, 700);
  };
  const setField = (op: any, k: string, v: string) => ubah(op, (r) => ({ ...r, f: { ...r.f, [k]: v } }));
  const setRm = (op: any, v: string) => ubah(op, (r) => ({ ...r, rm: v }));

  const items = useMemo(() => gabung(ops || [], list), [ops, list]);
  const itemsBulan = useMemo(() => items.filter((it) => it.tanggal.startsWith(bulan)), [items, bulan]);
  const itemsHari = useMemo(() => itemsBulan.filter((it) => it.tanggal === tanggal), [itemsBulan, tanggal]);
  const noDi = (it: MutuItem): number => itemsBulan.indexOf(it) + 1;

  /* riwayat: bulan -> tanggal (berdasarkan pasien yang ada) */
  const riwayat = useMemo(() => {
    const m = new Map<string, Map<string, { total: number; isi: number }>>();
    items.forEach((it) => {
      const b = it.tanggal.slice(0, 7);
      if (!m.has(b)) m.set(b, new Map());
      const t = m.get(b)!;
      const x = t.get(it.tanggal) || { total: 0, isi: 0 };
      x.total += 1; if (adaIsiApapun(it.row)) x.isi += 1;
      t.set(it.tanggal, x);
    });
    return [...m.entries()].sort((a, b) => (a[0] < b[0] ? 1 : -1)).map(([b, t]) => ({ bulan: b, hari: [...t.entries()].sort((x, y) => (x[0] < y[0] ? 1 : -1)) }));
  }, [items]);

  const bangunBlob = async (): Promise<Blob> => workbookKeBlob(await buatWorkbookMutu(itemsBulan, bulan));
  const unduh = async () => {
    setBusy(true);
    try {
      const blob = await bangunBlob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a"); a.href = url; a.download = `PemantauanMutu_${bulan}.xlsx`;
      document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 2000);
    } catch { showToast("Gagal membuat Excel", C.d); }
    setBusy(false);
  };
  const simpan = async () => {
    setBusy(true);
    const tunggu = [...pending.current];
    tunggu.forEach((id) => clearTimeout(timers.current[id]));
    const hasil = await Promise.all(tunggu.map((id) => kirim(id)));
    if (hasil.some((x) => !x)) { setBusy(false); return; }
    try {
      const folder = (dbxCfg?.path || "/kamarbedah/").replace(/[^/]+$/, "");
      const buf = await (await buatWorkbookMutu(itemsBulan, bulan)).xlsx.writeBuffer();
      const d = await dropboxUploadBytes(`${folder}PemantauanMutu_${bulan}.xlsx`, buf as unknown as ArrayBuffer);
      if (d.ok) showToast("Pemantauan mutu tersimpan (Supabase dan Excel di Dropbox)", C.s);
      else showToast("Data tersimpan, tetapi Excel ke Dropbox gagal: " + d.msg, C.w);
    } catch { showToast("Data tersimpan, tetapi Excel ke Dropbox gagal", C.w); }
    AmbilLogMedis("MUTU_OK", role || "-", "SIMPAN_MUTU_OK", `Pemantauan mutu ${bulan} disimpan`, { bulan, jumlahPasien: itemsBulan.length });
    setBusy(false);
  };

  const namaPasien = (it: MutuItem): string => { const n = String(it.op?.patient || "-"); return maskName(n, !!privacyMode); };
  const tombol = (warna: string): any => ({ border: "none", borderRadius: 10, padding: "9px 14px", fontWeight: 700, fontSize: 13, cursor: busy ? "wait" : "pointer", background: warna, color: "#fff", fontFamily: "inherit" });
  const tombolTipis: any = { border: "1px solid #d1d5db", borderRadius: 10, padding: "8px 12px", fontWeight: 700, fontSize: 13, cursor: "pointer", background: "#fff", color: KEL, fontFamily: "inherit" };

  /* susun isian grup jadi baris rapi: pasangan YA/TDK satu baris, K/S/B/KH satu baris */
  const susunBaris = (g: GrupForm): { judul: string; fs: FieldDef[] }[] => {
    if (g.id === "kat") return [{ judul: "Ukuran operasi", fs: g.f }];
    const out: { judul: string; fs: FieldDef[] }[] = [];
    for (let i = 0; i < g.f.length; i++) {
      const a = g.f[i]; const b = g.f[i + 1];
      if (a.t === "pilih" && b && b.t === "pilih" && a.k.endsWith("_ya") && b.k === a.k.replace(/_ya$/, "_tdk")) {
        out.push({ judul: a.l === "YA" ? "" : a.l.replace(/ YA$/, ""), fs: [a, b] }); i++;
      } else out.push({ judul: a.t === "pilih" ? "" : a.l, fs: [a] });
    }
    return out;
  };
  const gayaPilih: any = { padding: "6px 4px", borderRadius: 8, border: "1px solid #d1d5db", fontSize: 14, fontFamily: "inherit", background: "#fff", minWidth: 62 };
  const gayaIsi: any = { padding: "6px 8px", borderRadius: 8, border: "1px solid #d1d5db", fontSize: 14, fontFamily: "inherit", background: "#fff", width: 150, boxSizing: "border-box" };

  const kotakIsian = (it: MutuItem) => {
    const op = it.op; const row = it.row;
    return (
      <div style={{ padding: "12px", borderTop: "1px solid #e5e7eb", background: "#fafafa", display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))", gap: 10 }}>
        {GRUP_FORM.map((g) => {
          const terisi = adaIsi(row, g);
          const dokter = g.auto === "anest" ? op.anesthesiologist : g.auto === "bedah" ? op.surgeon : "";
          return (
            <div key={g.id} style={{ border: terisi ? "1px solid #86efac" : "1px solid #e5e7eb", borderRadius: 10, background: "#fff", padding: "8px 10px" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 4 }}>
                <span style={{ flex: 1, fontSize: 12, fontWeight: 800, color: KEL, lineHeight: 1.3 }}>{g.l}</span>
                {terisi && <span style={{ fontSize: 11, fontWeight: 800, color: "#16a34a" }}>{"\u2713"}</span>}
              </div>
              {g.auto && (
                <div style={{ fontSize: 12, color: "#4b5563", background: "#f3f4f6", borderRadius: 6, padding: "4px 8px", marginBottom: 4 }}>
                  {g.auto === "anest" ? "Dokter anestesi" : "Dokter bedah"}: <b>{String(dokter || "").trim() || "-"}</b>
                </div>
              )}
              {susunBaris(g).map((b, i) => (
                <div key={i} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 8, padding: "6px 0", borderTop: i > 0 || g.auto ? "1px dashed #e5e7eb" : "none" }}>
                  <span style={{ flex: 1, minWidth: 96, fontSize: 13, color: "#374151" }}>{b.judul}</span>
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                    {b.fs.map((fd) => fd.t === "pilih" ? (
                      <label key={fd.k} style={{ display: "flex", flexDirection: "column", alignItems: "center", fontSize: 10, fontWeight: 700, color: "#6b7280", gap: 2 }}>
                        <span>{/(YA|TDK)$/.test(fd.l) ? (fd.l.match(/(YA|TDK)$/) as RegExpMatchArray)[0] : fd.l}</span>
                        <select value={row?.f?.[fd.k] ?? ""} onChange={(e) => setField(op, fd.k, e.target.value)} style={gayaPilih}>
                          <option value="">-</option><option value="1">1</option><option value="0">0</option>
                        </select>
                      </label>
                    ) : (
                      <input key={fd.k} type={fd.t === "jam" ? "time" : fd.t === "angka" ? "number" : "text"} value={row?.f?.[fd.k] ?? ""} onChange={(e) => setField(op, fd.k, e.target.value)} style={gayaIsi} />
                    ))}
                  </div>
                </div>
              ))}
            </div>
          );
        })}
      </div>
    );
  };

  const th: any = { border: "1px solid #9ca3af", padding: "4px 6px", background: "#f3f4f6", fontSize: 11, fontWeight: 800, textAlign: "center", whiteSpace: "nowrap" };
  const td: any = { border: "1px solid #d1d5db", padding: "3px 6px", fontSize: 12, textAlign: "center", whiteSpace: "nowrap" };

  return (
    <div style={{ maxWidth: 1040, margin: "0 auto", padding: "4px 2px 40px" }}>
      <div style={{ textAlign: "center", marginBottom: 14 }}>
        <div style={{ fontSize: 18, fontWeight: 800, letterSpacing: 0.5 }}>PEMANTAUAN MUTU KAMAR BEDAH</div>
        <div style={{ fontSize: 14, color: "#6b7280", fontWeight: 600 }}>{namaBulan(bulan)}</div>
      </div>

      <div style={{ display: "flex", flexWrap: "wrap", gap: 10, alignItems: "center", justifyContent: "space-between", padding: "10px 12px", border: "1px solid #e5e7eb", borderRadius: 12, background: "#f9fafb", marginBottom: 6 }}>
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          <input type="date" value={tanggal} onChange={(e) => { if (e.target.value) { setTanggal(e.target.value); setBuka(null); } }} style={{ padding: "8px 10px", borderRadius: 10, border: "1px solid #d1d5db", fontSize: 14, fontFamily: "inherit", background: "#fff" }} />
          <button style={tombolTipis} onClick={() => { setTanggal(hariIni); setBuka(null); }}>Hari ini</button>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          <button style={tombolTipis} onClick={() => setPreview(true)}>Preview</button>
          <button style={tombol("#4f46e5")} disabled={busy} onClick={simpan}>Simpan</button>
          <button style={tombol("#059669")} disabled={busy} onClick={unduh}>Unduh Excel</button>
        </div>
      </div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 8, justifyContent: "space-between", fontSize: 12, color: "#6b7280", margin: "0 4px 12px" }}>
        <span><b>{fmtIso(tanggal)}</b> · {itemsHari.length} pasien · hijau = sudah diisi, putih = belum</span>
        <span style={{ color: status === "gagal" ? "#dc2626" : status === "tersimpan" ? "#16a34a" : "#6b7280", fontWeight: 600 }}>
          {status === "menyimpan" ? "Menyimpan..." : status === "tersimpan" ? "Tersimpan otomatis" : status === "gagal" ? "Gagal menyimpan" : "Isian tersimpan otomatis"}
        </span>
      </div>

      {itemsHari.length === 0 && <div style={{ padding: 24, textAlign: "center", color: "#9ca3af", border: "1px dashed #d1d5db", borderRadius: 12 }}>Tidak ada pasien operasi pada tanggal ini.</div>}

      {itemsHari.map((it) => {
        const op = it.op; const id = String(op.id); const terbuka = buka === id;
        const sudah = adaIsiApapun(it.row); const jmlIsi = GRUP_FORM.filter((g) => adaIsi(it.row, g)).length;
        const rmOp = String(op.rm ?? "").trim();
        return (
          <div key={id} style={{ border: sudah ? "1px solid #86efac" : "1px solid #e5e7eb", borderLeft: sudah ? "6px solid #16a34a" : "6px solid #e5e7eb", borderRadius: 12, background: sudah ? "#f0fdf4" : "#fff", marginBottom: 8, overflow: "hidden" }}>
            <div onClick={() => setBuka(terbuka ? null : id)} style={{ padding: "10px 12px", cursor: "pointer" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <span style={{ width: 26, height: 26, borderRadius: 999, background: sudah ? "#16a34a" : "#e5e7eb", color: sudah ? "#fff" : "#374151", fontSize: 12, fontWeight: 800, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>{noDi(it)}</span>
                <span style={{ flex: 1, minWidth: 0, fontWeight: 700, fontSize: 15, color: sudah ? "#166534" : "inherit", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{namaPasien(it)}</span>
                {sudah && <span style={{ fontSize: 11, fontWeight: 800, padding: "2px 8px", borderRadius: 999, background: "#16a34a", color: "#fff", flexShrink: 0 }}>{jmlIsi}/{GRUP_FORM.length}</span>}
                <span style={{ fontSize: 13, fontWeight: 700, color: "#4f46e5", flexShrink: 0 }}>{terbuka ? "Tutup" : sudah ? "Edit" : "Isi"}</span>
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: "4px 14px", alignItems: "center", marginTop: 6, paddingLeft: 36, fontSize: 12, color: "#6b7280" }}>
                <span>{it.jam || "-"}</span>
                {rmOp ? (
                  <span>RM <b style={{ color: KEL }}>{rmOp}</b></span>
                ) : (
                  <span style={{ display: "inline-flex", alignItems: "center", gap: 4 }}>RM
                    <input placeholder="isi manual" value={it.row?.rm ?? ""} onClick={(e) => e.stopPropagation()} onChange={(e) => setRm(op, e.target.value)} style={{ width: 110, padding: "4px 8px", borderRadius: 8, border: "1px solid #d1d5db", fontSize: 13, fontFamily: "inherit", background: "#fff" }} />
                  </span>
                )}
                <span>Bedah: {String(op.surgeon || "-")}</span>
                <span>Anestesi: {String(op.anesthesiologist || "-")}</span>
              </div>
            </div>
            {terbuka && kotakIsian(it)}
          </div>
        );
      })}

      {/* Riwayat per bulan dan tanggal */}
      <div style={{ marginTop: 22 }}>
        <div style={{ fontSize: 13, fontWeight: 800, color: KEL, marginBottom: 8 }}>RIWAYAT</div>
        {riwayat.length === 0 && <div style={{ fontSize: 12, color: "#9ca3af" }}>Belum ada data.</div>}
        {(semuaBulan ? riwayat : riwayat.slice(0, 6)).map((rb) => (
          <div key={rb.bulan} style={{ border: "1px solid #e5e7eb", borderRadius: 12, padding: "10px 12px", marginBottom: 8, background: "#fff" }}>
            <div style={{ fontWeight: 800, fontSize: 13, marginBottom: 6 }}>{namaBulan(rb.bulan)}</div>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
              {rb.hari.map(([tgl, x]) => (
                <button key={tgl} onClick={() => { setTanggal(tgl); setBuka(null); window.scrollTo({ top: 0, behavior: "smooth" }); }} style={{ border: tgl === tanggal ? "2px solid #4f46e5" : "1px solid #d1d5db", background: x.isi >= x.total ? "#ecfdf5" : "#fff", borderRadius: 8, padding: "4px 8px", fontSize: 12, cursor: "pointer", fontFamily: "inherit" }}>
                  {tgl.slice(8, 10)}-{tgl.slice(5, 7)} · {x.isi}/{x.total}
                </button>
              ))}
            </div>
          </div>
        ))}
        {riwayat.length > 6 && <button style={tombolTipis} onClick={() => setSemuaBulan(!semuaBulan)}>{semuaBulan ? "Ringkas" : "Tampilkan semua bulan"}</button>}
      </div>

      {/* Preview bulan */}
      {preview && (
        <div onClick={() => setPreview(false)} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", zIndex: 9999, display: "flex", alignItems: "center", justifyContent: "center", padding: 12 }}>
          <div onClick={(e) => e.stopPropagation()} style={{ background: "#fff", borderRadius: 14, width: "100%", maxWidth: 1200, maxHeight: "92vh", display: "flex", flexDirection: "column" }}>
            <div style={{ padding: "12px 14px", display: "flex", flexWrap: "wrap", gap: 8, alignItems: "center", borderBottom: "1px solid #e5e7eb" }}>
              <div style={{ flex: 1, minWidth: 200 }}>
                <div style={{ fontWeight: 800, fontSize: 14 }}>PEMANTAUAN MUTU KAMAR BEDAH</div>
                <div style={{ fontSize: 12, color: "#6b7280" }}>{namaBulan(bulan)} · {itemsBulan.length} pasien (tampilan ringkas; Excel mengikuti format mutu.xlsx)</div>
              </div>
              <button style={tombol("#4f46e5")} disabled={busy} onClick={simpan}>Simpan</button>
              <button style={tombol("#059669")} disabled={busy} onClick={unduh}>Unduh Excel</button>
              <button style={tombolTipis} onClick={() => setPreview(false)}>Tutup</button>
            </div>
            <div style={{ overflow: "auto", padding: 10 }}>
              <table style={{ borderCollapse: "collapse", minWidth: "100%" }}>
                <thead>
                  <tr>{H1.map(([t, c, r], i) => <th key={i} colSpan={c} rowSpan={r} style={th}>{t}</th>)}</tr>
                  <tr>{H2.map(([t, c, r], i) => <th key={i} colSpan={c} rowSpan={r} style={th}>{t}</th>)}</tr>
                  <tr>{Array.from({ length: JML_PASANG }).flatMap((_, i) => [<th key={`a${i}`} style={th}>YA</th>, <th key={`b${i}`} style={th}>TDK</th>])}</tr>
                </thead>
                <tbody>
                  {itemsBulan.map((it, i) => (
                    <tr key={String(it.op.id)}>{nilaiBaris(it, i + 1).map((v, c) => <td key={c} style={{ ...td, textAlign: c === 9 || c === 24 || c === 16 ? "left" : "center" }}>{v === null ? "" : String(v)}</td>)}</tr>
                  ))}
                  {itemsBulan.length === 0 && <tr><td style={td} colSpan={35}>Tidak ada pasien operasi bulan ini.</td></tr>}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
