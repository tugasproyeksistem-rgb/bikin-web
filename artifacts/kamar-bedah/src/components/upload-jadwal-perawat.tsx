/* src/components/upload-jadwal-perawat.tsx
   Kartu "Upload Jadwal Perawat" di tab Staf (admin). Membaca file Excel/CSV
   jadwal semua perawat 1 bulan, tampil ringkasan, lalu simpan ke Supabase
   (kb_jadwal_perawat, satu baris per bulan). Dipakai tab Jadwal & Peserta. */

import { useRef, useState, useEffect } from "react";
import { C, iS } from "../constants";
import { SUPA_CLIENT, AmbilLogMedis } from "../services/supabase";
import { Card, SH, Btn } from "./primitives";
import { bacaJadwalPerawat } from "../utils/jadwal-perawat";
import type { HasilBaca, JadwalPerawatBulan } from "../utils/jadwal-perawat";

const NAMA_BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
const labelBulan = (ym: string): string => { const [y, m] = ym.split("-").map(Number); return `${NAMA_BULAN[m - 1] || ""} ${y}`; };
const hitungKode = (h: HasilBaca): Record<string, number> => {
  const t: Record<string, number> = { P: 0, PG: 0, S: 0, SG: 0, T: 0 };
  h.perawat.forEach((p) => p.kode.forEach((k) => { if (Object.prototype.hasOwnProperty.call(t, k)) t[k]++; }));
  return t;
};

export function UploadJadwalPerawat({ showToast, upsertOneToSupa }: any) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [hasil, setHasil] = useState<HasilBaca | null>(null);
  const [bulan, setBulan] = useState<string>(() => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`; });
  const [busy, setBusy] = useState(false);
  const [tersimpan, setTersimpan] = useState<{ id: string; n: number }[]>([]);

  const muatDaftar = async () => {
    try {
      const r = await SUPA_CLIENT.from("kb_jadwal_perawat").select("data").order("updated_at", { ascending: false });
      setTersimpan((r.data || []).map((x: any) => x.data).filter(Boolean)
        .map((d: any) => ({ id: `${d.tahun}-${String(d.bulan).padStart(2, "0")}`, n: (d.perawat || []).length })));
    } catch { /* offline — daftar dikosongkan */ }
  };
  useEffect(() => { void muatDaftar(); }, []);

  const pilihFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    e.target.value = "";
    if (!f) return;
    setBusy(true);
    try {
      const h = await bacaJadwalPerawat(f);
      if (!h.perawat.length) {
        setHasil(null);
        showToast("⚠ Tidak ada baris perawat terbaca. Pastikan ada kolom NAMA, tanggal 1-31, dan kode P/PG/S/SG/T per sel.", C.d);
      } else {
        setHasil(h);
        if (h.tahun && h.bulan) setBulan(`${h.tahun}-${String(h.bulan).padStart(2, "0")}`);
        else showToast("⚠ Bulan tidak terbaca dari file, pilih bulannya di bawah", C.w);
      }
    } catch (err: any) { setHasil(null); showToast("⚠ Gagal membaca file: " + (err?.message || "error"), C.d); }
    setBusy(false);
  };

  const simpan = async () => {
    if (!hasil) return;
    const [y, m] = bulan.split("-").map(Number);
    if (!y || !m) { showToast("⚠ Pilih bulan jadwal dulu", C.w); return; }
    if (tersimpan.some((t) => t.id === bulan) && !window.confirm(`Jadwal perawat ${labelBulan(bulan)} sudah ada. Timpa dengan file ini?`)) return;
    setBusy(true);
    const data: JadwalPerawatBulan = { id: `jp_${y}_${String(m).padStart(2, "0")}`, tahun: y, bulan: m, perawat: hasil.perawat, updated_at: new Date().toISOString() };
    const res = await upsertOneToSupa("kb_jadwal_perawat", data);
    if (!res?.ok) { showToast(`⚠ Gagal menyimpan ke server: ${res?.error || "error"}`, C.d); setBusy(false); return; }
    AmbilLogMedis("JADWAL_PERAWAT", "Admin", "UPLOAD_JADWAL_PERAWAT", `Upload jadwal perawat ${bulan}`, { id: data.id, jumlahPerawat: data.perawat.length });
    showToast(`✓ Jadwal perawat ${labelBulan(bulan)} tersimpan (${data.perawat.length} perawat)`, C.s);
    setHasil(null);
    await muatDaftar();
    setBusy(false);
  };

  const t = hasil ? hitungKode(hasil) : null;
  return (
    <Card style={{ background: "#E0F2F1", border: "2px solid #80CBC4" }}>
      <SH label="🩺 Upload Jadwal Perawat (semua perawat)" color="#00695C" />
      <div style={{ background: "#B2DFDB55", borderRadius: 8, padding: "10px 12px", marginBottom: 12, fontSize: 12, color: "#004D40", lineHeight: 1.7 }}>
        <b>Fungsi:</b> mengisi perawat jaga (Pagi, PG, Sore, SG, Cuti) di tab Jadwal & Peserta.<br />
        Format: Excel/CSV seperti ekspor tab Generator. Kolom NAMA, lalu tanggal 1-31 berisi kode P, PG, S, SG, T.
      </div>
      <input ref={inputRef} type="file" accept=".csv,.txt,.xlsx,.xls" onChange={pilihFile} style={{ display: "none" }} />
      <Btn sm color="#00695C" onClick={() => inputRef.current && inputRef.current.click()} disabled={busy}>📂 Upload Jadwal Perawat (Excel/CSV)</Btn>

      {hasil && t && (
        <div style={{ marginTop: 12, background: "#fff", borderRadius: 10, padding: "10px 12px", border: "1px solid #B2DFDB" }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: "#004D40" }}>Terbaca {hasil.perawat.length} perawat{hasil.diabaikan > 0 ? ` (${hasil.diabaikan} baris diabaikan)` : ""}</div>
          <div style={{ fontSize: 12, color: C.tL, margin: "4px 0 8px" }}>
            Total kode: P {t.P} • PG {t.PG} • S {t.S} • SG {t.SG} • T {t.T}
          </div>
          <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 4 }}>Bulan jadwal:</div>
          <input style={iS} type="month" value={bulan} onChange={(e) => setBulan(e.target.value)} />
          <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
            <Btn onClick={simpan} disabled={busy}>💾 Simpan jadwal {bulan ? labelBulan(bulan) : ""}</Btn>
            <Btn outline color={C.g} onClick={() => setHasil(null)} disabled={busy}>Batal</Btn>
          </div>
        </div>
      )}

      {tersimpan.length > 0 && (
        <div style={{ marginTop: 10, fontSize: 12, color: "#004D40", fontWeight: 600 }}>
          ✓ Tersimpan: {tersimpan.slice(0, 6).map((x) => `${labelBulan(x.id)} (${x.n} perawat)`).join(" • ")}
        </div>
      )}
    </Card>
  );
}
