/* src/components/view-implan.tsx
   Tab grup MUTU: penggunaan implant (Tulang; Mata dan Hernia memakai komponen yang sama lewat prop kelompok).
   Pasien tampil otomatis dari daftar operasi yang di tab Daftar ditandai memakai implant + kelompoknya.
   Isian tersimpan otomatis ke Supabase (kb_implan_ok, realtime). Simpan = Excel bulanan ke Dropbox; Unduh = Excel yang sama. */
import { useState, useEffect, useMemo, useRef } from "react";
import { C } from "../constants";
import { maskName } from "../utils";
import { SUPA_CLIENT, AmbilLogMedis } from "../services/supabase";
import { dropboxUploadBytes } from "../services/dropbox";
import { isoLokal } from "../utils/laporan-harian";
import { workbookKeBlob } from "../utils/jadwal-ok";
import { NAMA_BULAN, rmItem } from "../utils/mutu-ok";
import type { KodeImplan } from "../constants/kode-implant";
import type { ImplanRow, ImplanItem, Kelompok, ItemImplan } from "../utils/implan-ok";
import { KONFIG_IMPLAN, adaIsiImplan, baruImplan, buatWorkbookImplan, fmtLahir, gabungImplan, idImplan, labelKatalog, susunBaris } from "../utils/implan-ok";

const KEL = "#374151";
const fmtIso = (s: string): string => (s && s.length >= 10 ? `${s.slice(8, 10)}-${s.slice(5, 7)}-${s.slice(0, 4)}` : s);
const namaBulan = (b: string): string => { const n = parseInt(b.slice(5, 7), 10); const m = NAMA_BULAN[n - 1] || ""; return m.charAt(0) + m.slice(1).toLowerCase() + " " + b.slice(0, 4); };
const norm = (s: string): string => s.toLowerCase().replace(/,/g, ".").replace(/\s+/g, " ").trim();

/* kotak ketik dengan tebakan: ketik sebagian nama/kode, pilih dari daftar, atau tambah tulisan sendiri */
function PilihImplan({ katalog, onTambah }: { katalog: KodeImplan[]; onTambah: (it: ItemImplan) => void }) {
  const [q, setQ] = useState("");
  const [fokus, setFokus] = useState(false);
  const daftar = useMemo(() => katalog.map((k) => ({ k, label: labelKatalog(k), cari: norm(labelKatalog(k) + " " + k.k) })), [katalog]);
  const hasil = useMemo(() => {
    const t = norm(q).split(" ").filter(Boolean);
    if (!t.length) return [];
    return daftar.filter((d) => t.every((x) => d.cari.includes(x))).slice(0, 8);
  }, [q, daftar]);
  const pilih = (d: { k: KodeImplan; label: string }) => { onTambah({ kode: d.k.k, nama: d.label, jumlah: 1 }); setQ(""); };
  const manual = () => { const v = q.trim(); if (v) { onTambah({ kode: "", nama: v, jumlah: 1 }); setQ(""); } };
  const adaPersis = hasil.some((d) => norm(d.label) === norm(q));
  return (
    <div style={{ position: "relative" }}>
      <input value={q} placeholder="Ketik nama atau kode implant, mis. cortical 3,5 14" autoComplete="off" spellCheck={false}
        onChange={(e) => setQ(e.target.value)} onFocus={() => setFokus(true)} onBlur={() => setTimeout(() => setFokus(false), 150)}
        onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); if (hasil[0]) pilih(hasil[0]); else manual(); } if (e.key === "Escape") setQ(""); }}
        style={{ width: "100%", boxSizing: "border-box", padding: "9px 10px", borderRadius: 10, border: "1px solid #d1d5db", fontSize: 14, fontFamily: "inherit", background: "#fff" }} />
      {fokus && q.trim() !== "" && (
        <div style={{ position: "absolute", left: 0, right: 0, top: "100%", zIndex: 20, marginTop: 4, background: "#fff", border: "1px solid #d1d5db", borderRadius: 10, boxShadow: "0 8px 24px rgba(0,0,0,0.12)", maxHeight: 280, overflowY: "auto" }}>
          {hasil.map((d, i) => (
            <div key={i} onMouseDown={(e) => { e.preventDefault(); pilih(d); }} style={{ padding: "8px 10px", cursor: "pointer", borderBottom: "1px solid #f3f4f6", display: "flex", justifyContent: "space-between", gap: 8, fontSize: 13 }}>
              <span style={{ fontWeight: 600, color: KEL }}>{d.label}</span>
              <span style={{ color: "#9ca3af", fontSize: 11, whiteSpace: "nowrap" }}>{d.k.k}</span>
            </div>
          ))}
          {!adaPersis && (
            <div onMouseDown={(e) => { e.preventDefault(); manual(); }} style={{ padding: "8px 10px", cursor: "pointer", fontSize: 13, color: "#4f46e5", fontWeight: 700 }}>
              + Tambah tulisan sendiri: {q.trim()}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

export function ViewImplan({ ops, showToast, role, privacyMode, upsertOneToSupa, dbxCfg, kelompok = "tulang" }: any) {
  const kel = kelompok as Kelompok;
  const cfg = KONFIG_IMPLAN[kel];
  const hariIni = isoLokal(new Date());
  const [tanggal, setTanggal] = useState<string>(hariIni);
  const [list, setList] = useState<ImplanRow[]>([]);
  const [buka, setBuka] = useState<string | null>(null);
  const [preview, setPreview] = useState(false);
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState<"" | "menyimpan" | "tersimpan" | "gagal">("");
  const [semuaBulan, setSemuaBulan] = useState(false);
  const listRef = useRef<ImplanRow[]>([]);
  const timers = useRef<Record<string, any>>({});
  const pending = useRef<Set<string>>(new Set());
  const bulan = tanggal.slice(0, 7);

  useEffect(() => {
    let batal = false;
    (async () => {
      try {
        const r = await SUPA_CLIENT.from("kb_implan_ok").select("data");
        if (!batal && r.data) { const d = r.data.map((x: any) => x.data).filter(Boolean) as ImplanRow[]; listRef.current = d; setList(d); }
      } catch { /* offline — mulai dari kosong */ }
    })();
    const ch = SUPA_CLIENT.channel("kb-implan-ok-rt-" + kel)
      .on("postgres_changes", { event: "*", schema: "public", table: "kb_implan_ok" }, (p: any) => {
        if (p.eventType === "DELETE") return;
        const d = p.new?.data as ImplanRow | undefined;
        if (!d?.id) return;
        const lokal = listRef.current.find((r) => r.id === d.id);
        if (lokal && String(lokal.updated_at || "") >= String(d.updated_at || "")) return;
        const next = [d, ...listRef.current.filter((r) => r.id !== d.id)];
        listRef.current = next; setList(next);
      }).subscribe();
    return () => {
      batal = true; SUPA_CLIENT.removeChannel(ch);
      pending.current.forEach((id) => { clearTimeout(timers.current[id]); const cur = listRef.current.find((r) => r.id === id); if (cur) upsertOneToSupa("kb_implan_ok", cur); });
    };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const kirim = async (id: string): Promise<boolean> => {
    pending.current.delete(id);
    const cur = listRef.current.find((r) => r.id === id);
    if (!cur) return true;
    const res = await upsertOneToSupa("kb_implan_ok", cur);
    if (!res?.ok) { setStatus("gagal"); showToast(`Gagal menyimpan ke server: ${res?.error || "error"}. Isian masih di layar ini.`, C.d); return false; }
    if (pending.current.size === 0) setStatus("tersimpan");
    return true;
  };
  const ubah = (op: any, patch: (r: ImplanRow) => ImplanRow) => {
    const id = idImplan(op.id);
    const lama = listRef.current.find((r) => r.id === id) || baruImplan(op);
    const baru: ImplanRow = { ...patch(lama), tanggal: String(op.date || lama.tanggal), updated_at: new Date().toISOString() };
    const next = [baru, ...listRef.current.filter((r) => r.id !== id)];
    listRef.current = next; setList(next);
    setStatus("menyimpan");
    pending.current.add(id);
    clearTimeout(timers.current[id]);
    timers.current[id] = setTimeout(() => { kirim(id); }, 700);
  };
  const setBidang = (op: any, k: "register" | "tglLahir" | "noTelp" | "alamat" | "ket", v: string) => ubah(op, (r) => ({ ...r, [k]: v }));
  const tambahItem = (op: any, it: ItemImplan) => ubah(op, (r) => {
    const ada = r.items.findIndex((x) => x.nama === it.nama);
    return ada >= 0 ? { ...r, items: r.items.map((x, i) => (i === ada ? { ...x, jumlah: x.jumlah + it.jumlah } : x)) } : { ...r, items: [...r.items, it] };
  });
  const ubahJumlah = (op: any, idx: number, v: string) => ubah(op, (r) => ({ ...r, items: r.items.map((x, i) => (i === idx ? { ...x, jumlah: Math.max(1, parseInt(v, 10) || 1) } : x)) }));
  const hapusItem = (op: any, idx: number) => ubah(op, (r) => ({ ...r, items: r.items.filter((_, i) => i !== idx) }));

  const items = useMemo(() => gabungImplan(ops || [], list, kel), [ops, list, kel]);
  const itemsBulan = useMemo(() => items.filter((it) => it.tanggal.startsWith(bulan)), [items, bulan]);
  const itemsHari = useMemo(() => itemsBulan.filter((it) => it.tanggal === tanggal), [itemsBulan, tanggal]);
  const noDi = (it: ImplanItem): number => itemsBulan.indexOf(it) + 1;
  const riwayat = useMemo(() => {
    const m = new Map<string, Map<string, { total: number; isi: number }>>();
    items.forEach((it) => {
      const b = it.tanggal.slice(0, 7);
      if (!m.has(b)) m.set(b, new Map());
      const t = m.get(b)!; const x = t.get(it.tanggal) || { total: 0, isi: 0 };
      x.total += 1; if (it.row && it.row.items.length > 0) x.isi += 1;
      t.set(it.tanggal, x);
    });
    return [...m.entries()].sort((a, b) => (a[0] < b[0] ? 1 : -1)).map(([b, t]) => ({ bulan: b, hari: [...t.entries()].sort((x, y) => (x[0] < y[0] ? 1 : -1)) }));
  }, [items]);

  const unduh = async () => {
    setBusy(true);
    try {
      const blob = await workbookKeBlob(await buatWorkbookImplan(itemsBulan, bulan, kel));
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a"); a.href = url; a.download = `${cfg?.berkas || "Implant"}_${bulan}.xlsx`;
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
      const buf = await (await buatWorkbookImplan(itemsBulan, bulan, kel)).xlsx.writeBuffer();
      const d = await dropboxUploadBytes(`${folder}${cfg?.berkas || "Implant"}_${bulan}.xlsx`, buf as unknown as ArrayBuffer);
      if (d.ok) showToast("Data implant tersimpan (Supabase dan Excel di Dropbox)", C.s);
      else showToast("Data tersimpan, tetapi Excel ke Dropbox gagal: " + d.msg, C.w);
    } catch { showToast("Data tersimpan, tetapi Excel ke Dropbox gagal", C.w); }
    AmbilLogMedis("IMPLANT_OK", role || "-", "SIMPAN_IMPLANT_OK", `Penggunaan implant ${kel} ${bulan} disimpan`, { bulan, kelompok: kel, jumlahPasien: itemsBulan.length });
    setBusy(false);
  };

  const namaPasien = (it: ImplanItem): string => { const n = String(it.op?.patient || "-"); return privacyMode ? maskName(n, true) : n; };
  const tombol = (warna: string): any => ({ border: "none", borderRadius: 10, padding: "9px 14px", fontWeight: 700, fontSize: 13, cursor: busy ? "wait" : "pointer", background: warna, color: "#fff", fontFamily: "inherit" });
  const tombolTipis: any = { border: "1px solid #d1d5db", borderRadius: 10, padding: "8px 12px", fontWeight: 700, fontSize: 13, cursor: "pointer", background: "#fff", color: KEL, fontFamily: "inherit" };
  const gayaIsi: any = { width: "100%", boxSizing: "border-box", padding: "8px 10px", borderRadius: 8, border: "1px solid #d1d5db", fontSize: 14, fontFamily: "inherit", background: "#fff" };
  const labelKecil: any = { display: "flex", flexDirection: "column", gap: 3, fontSize: 11, fontWeight: 700, color: "#6b7280" };

  if (!cfg) return <div style={{ padding: 24, textAlign: "center", color: "#9ca3af" }}>Template Excel untuk kelompok ini belum tersedia.</div>;

  const kotakIsian = (it: ImplanItem) => {
    const op = it.op; const row = it.row;
    const telpOp = String(op.noHp ?? "").trim(); const alamatOp = String(op.alamatPasien ?? "").trim();
    const lahirOp = String(op.tglLahir ?? "").trim();
    return (
      <div style={{ padding: 12, borderTop: "1px solid #e5e7eb", background: "#fafafa", display: "grid", gap: 10 }}>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 10 }}>
          {cfg.pakaiRegister && <label style={labelKecil}><span>NO REGISTER (RM: {rmItem(it as any) || "-"})</span>
            <input value={row?.register ?? ""} onChange={(e) => setBidang(op, "register", e.target.value)} placeholder="Nomor register" style={gayaIsi} /></label>}
          <label style={labelKecil}><span>TANGGAL LAHIR {lahirOp ? "(dari data pasien)" : ""}</span>
            {lahirOp ? <div style={{ ...gayaIsi, background: "#f3f4f6", color: KEL }}>{fmtLahir(lahirOp)}</div>
              : <input type="date" value={row?.tglLahir ?? ""} onChange={(e) => setBidang(op, "tglLahir", e.target.value)} style={gayaIsi} />}</label>
          <label style={labelKecil}><span>NO TELP {telpOp ? "(dari data pasien)" : ""}</span>
            {telpOp ? <div style={{ ...gayaIsi, background: "#f3f4f6", color: KEL }}>{telpOp}</div>
              : <input type="tel" inputMode="tel" value={row?.noTelp ?? ""} onChange={(e) => setBidang(op, "noTelp", e.target.value)} placeholder="08xxxxxxxxxx" style={gayaIsi} />}</label>
        </div>
        <label style={labelKecil}><span>ALAMAT {alamatOp ? "(dari data pasien)" : ""}</span>
          {alamatOp ? <div style={{ ...gayaIsi, background: "#f3f4f6", color: KEL }}>{alamatOp}</div>
            : <textarea value={row?.alamat ?? ""} onChange={(e) => setBidang(op, "alamat", e.target.value)} placeholder="Alamat pasien" rows={2} style={{ ...gayaIsi, resize: "vertical" }} />}</label>

        <div style={{ border: "1px solid #e5e7eb", borderRadius: 10, background: "#fff", padding: "10px 12px" }}>
          <div style={{ fontSize: 12, fontWeight: 800, color: KEL, marginBottom: 6 }}>IMPLANT YANG DIPAKAI</div>
          <PilihImplan katalog={cfg.katalog} onTambah={(x) => tambahItem(op, x)} />
          {(row?.items || []).length === 0 && <div style={{ fontSize: 12, color: "#9ca3af", marginTop: 8 }}>Belum ada implant. Ketik nama atau kode, lalu pilih dari daftar. Boleh lebih dari satu.</div>}
          {(row?.items || []).map((x, i) => (
            <div key={i} style={{ display: "flex", alignItems: "center", gap: 8, padding: "7px 0", borderTop: i === 0 ? "none" : "1px dashed #e5e7eb", marginTop: i === 0 ? 8 : 0 }}>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontSize: 13, fontWeight: 600, color: KEL }}>{x.nama}</div>
                {x.kode && <div style={{ fontSize: 11, color: "#9ca3af" }}>{x.kode}</div>}
              </div>
              <label style={{ ...labelKecil, alignItems: "center" }}><span>JUMLAH</span>
                <input type="number" min={1} value={x.jumlah} onChange={(e) => ubahJumlah(op, i, e.target.value)} style={{ ...gayaIsi, width: 68, textAlign: "center" }} /></label>
              <button type="button" title="Hapus" onClick={() => hapusItem(op, i)} style={{ border: "none", background: "#fee2e2", color: "#b91c1c", borderRadius: 8, width: 30, height: 30, fontWeight: 800, cursor: "pointer", fontFamily: "inherit" }}>×</button>
            </div>
          ))}
        </div>

        <label style={labelKecil}><span>{cfg.ketLabel}</span>
          <textarea value={row?.ket ?? ""} onChange={(e) => setBidang(op, "ket", e.target.value)} placeholder="Catatan (opsional)" rows={2} style={{ ...gayaIsi, resize: "vertical" }} /></label>
      </div>
    );
  };

  const th: any = { border: "1px solid #9ca3af", padding: "4px 6px", background: "#f3f4f6", fontSize: 11, fontWeight: 800, textAlign: "center", whiteSpace: "nowrap" };
  const td: any = { border: "1px solid #d1d5db", padding: "3px 6px", fontSize: 12, textAlign: "left" };
  const modelBaris = susunBaris(itemsBulan, cfg.pakaiRegister);

  return (
    <div style={{ maxWidth: 1040, margin: "0 auto", padding: "4px 2px 40px" }}>
      <div style={{ textAlign: "center", marginBottom: 14 }}>
        <div style={{ fontSize: 18, fontWeight: 800, letterSpacing: 0.5 }}>{cfg.judul}</div>
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
        <span><b>{fmtIso(tanggal)}</b> · {itemsHari.length} pasien · hijau = implant sudah diisi, putih = belum</span>
        <span style={{ color: status === "gagal" ? "#dc2626" : status === "tersimpan" ? "#16a34a" : "#6b7280", fontWeight: 600 }}>
          {status === "menyimpan" ? "Menyimpan..." : status === "tersimpan" ? "Tersimpan otomatis" : status === "gagal" ? "Gagal menyimpan" : "Isian tersimpan otomatis"}
        </span>
      </div>

      {itemsHari.length === 0 && <div style={{ padding: 24, textAlign: "center", color: "#9ca3af", border: "1px dashed #d1d5db", borderRadius: 12 }}>Tidak ada pasien pada tanggal ini. Pasien muncul otomatis setelah didaftarkan di tab Daftar dengan pilihan implant.</div>}

      {itemsHari.map((it) => {
        const op = it.op; const id = String(op.id); const terbuka = buka === id;
        const jml = it.row?.items.length || 0; const sudah = jml > 0; const adaData = adaIsiImplan(it.row);
        return (
          <div key={id} style={{ border: sudah ? "1px solid #86efac" : "1px solid #e5e7eb", borderLeft: sudah ? "6px solid #16a34a" : "6px solid #e5e7eb", borderRadius: 12, background: sudah ? "#f0fdf4" : "#fff", marginBottom: 8, overflow: "hidden" }}>
            <div onClick={() => setBuka(terbuka ? null : id)} style={{ padding: "10px 12px", cursor: "pointer" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <span style={{ width: 26, height: 26, borderRadius: 999, background: sudah ? "#16a34a" : "#e5e7eb", color: sudah ? "#fff" : "#374151", fontSize: 12, fontWeight: 800, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>{noDi(it)}</span>
                <span style={{ flex: 1, minWidth: 0, fontWeight: 700, fontSize: 15, color: sudah ? "#166534" : "inherit", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{namaPasien(it)}</span>
                {sudah && <span style={{ fontSize: 11, fontWeight: 800, padding: "2px 8px", borderRadius: 999, background: "#16a34a", color: "#fff", flexShrink: 0 }}>{jml} implant</span>}
                <span style={{ fontSize: 13, fontWeight: 700, color: "#4f46e5", flexShrink: 0 }}>{terbuka ? "Tutup" : adaData ? "Edit" : "Isi"}</span>
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: "4px 14px", marginTop: 6, paddingLeft: 36, fontSize: 12, color: "#6b7280" }}>
                <span>{it.jam || "-"}</span>
                <span>RM <b style={{ color: KEL }}>{rmItem(it as any) || "-"}</b></span>
                <span>Bedah: {String(op.surgeon || "-")}</span>
                <span>{String(op.procedure || "")}</span>
              </div>
            </div>
            {terbuka && kotakIsian(it)}
          </div>
        );
      })}

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

      {preview && (
        <div onClick={() => setPreview(false)} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", zIndex: 9999, display: "flex", alignItems: "center", justifyContent: "center", padding: 12 }}>
          <div onClick={(e) => e.stopPropagation()} style={{ background: "#fff", borderRadius: 14, width: "100%", maxWidth: 1200, maxHeight: "92vh", display: "flex", flexDirection: "column" }}>
            <div style={{ padding: "12px 14px", display: "flex", flexWrap: "wrap", gap: 8, alignItems: "center", borderBottom: "1px solid #e5e7eb" }}>
              <div style={{ flex: 1, minWidth: 200 }}>
                <div style={{ fontWeight: 800, fontSize: 14 }}>{cfg.judul}</div>
                <div style={{ fontSize: 12, color: "#6b7280" }}>{namaBulan(bulan)} · {itemsBulan.length} pasien (tampilan ringkas; Excel mengikuti form contoh)</div>
              </div>
              <button style={tombol("#4f46e5")} disabled={busy} onClick={simpan}>Simpan</button>
              <button style={tombol("#059669")} disabled={busy} onClick={unduh}>Unduh Excel</button>
              <button style={tombolTipis} onClick={() => setPreview(false)}>Tutup</button>
            </div>
            <div style={{ overflow: "auto", padding: 10 }}>
              <table style={{ borderCollapse: "collapse", minWidth: "100%" }}>
                <thead><tr>{["NO", "TGL", cfg.hdrRm, "NAMA", "TGL LAHIR", "NO TELP", "ALAMAT", cfg.hdrImplan, "JUMLAH", "KET"].map((h) => <th key={h} style={th}>{h}</th>)}</tr></thead>
                <tbody>
                  {modelBaris.filter((b) => b.jenis !== "kosong").map((b, i) => (
                    <tr key={i}>
                      <td style={{ ...td, textAlign: "center" }}>{b.no ?? ""}</td><td style={{ ...td, textAlign: "center" }}>{b.tgl ?? ""}</td>
                      <td style={td}>{b.rmReg ?? ""}</td><td style={td}>{b.nama ?? ""}</td><td style={{ ...td, textAlign: "center" }}>{b.lahir ? fmtLahir(b.lahir) : ""}</td>
                      <td style={td}>{b.telp ?? ""}</td><td style={td}>{b.alamat ?? ""}</td><td style={td}>{b.impl ?? ""}</td>
                      <td style={{ ...td, textAlign: "center" }}>{b.jml ?? ""}</td><td style={td}>{b.ket ?? ""}</td>
                    </tr>
                  ))}
                  {itemsBulan.length === 0 && <tr><td style={td} colSpan={10}>Tidak ada pasien implant bulan ini.</td></tr>}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
