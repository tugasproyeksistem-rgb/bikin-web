/* src/components/riwayat-bulanan.tsx
   Riwayat laporan per bulan: kalender mini (pilih tanggal) + satu panel aksi untuk tanggal terpilih.
   Tanggal yang punya laporan ditandai; ketuk langsung membuka panelnya, tanpa scroll panjang. */

import { useState, useMemo } from "react";
import { C, iS } from "../constants";
import { Card, SH, Btn } from "./primitives";
import { fmtTgl, namaHari } from "../utils/laporan-harian";

const NAMA_BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
const HARI_SINGKAT = ["Sen", "Sel", "Rab", "Kam", "Jum", "Sab", "Min"];
const labelBulan = (b: string) => { const [y, m] = b.split("-").map(Number); return (NAMA_BULAN[(m || 1) - 1] || "") + " " + y; };

export interface ItemRiwayat { id: string; tanggal: string; label?: string; ket?: string; urut?: number; data: any; }

export function RiwayatBulanan({ judul = "Riwayat bulanan", items, role, kosongTeks, sibuk, onBuka, onUnduh, onHapus, onExcelBulan }: any) {
  const [pilih, setPilih] = useState("");
  const [tglPilih, setTglPilih] = useState("");
  const semua: ItemRiwayat[] = items || [];
  const bulanAda = useMemo(() => Array.from(new Set(semua.map((x) => String(x.tanggal).slice(0, 7)))).sort(), [semua]);
  const hariIni = (() => { const n = new Date(); return n.getFullYear() + "-" + String(n.getMonth() + 1).padStart(2, "0") + "-" + String(n.getDate()).padStart(2, "0"); })();
  const sekarang = hariIni.slice(0, 7);
  const bulan = pilih || bulanAda[bulanAda.length - 1] || sekarang;
  const ganti = (b: string) => { setPilih(b); setTglPilih(""); };
  const geser = (n: number) => {
    const [y, m] = bulan.split("-").map(Number);
    const d = new Date(y, m - 1 + n, 1);
    ganti(d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0"));
  };
  const daftarBulan = useMemo(() => semua
    .filter((x) => String(x.tanggal).startsWith(bulan))
    .sort((a, b) => String(a.tanggal).localeCompare(String(b.tanggal)) || (a.urut ?? 0) - (b.urut ?? 0)), [semua, bulan]);
  const grup = useMemo(() => {
    const m = new Map<string, ItemRiwayat[]>();
    daftarBulan.forEach((x) => { const k = String(x.tanggal); m.set(k, [...(m.get(k) || []), x]); });
    return Array.from(m.entries());
  }, [daftarBulan]);
  const peta = useMemo(() => new Map<string, ItemRiwayat[]>(grup), [grup]);

  const [thn, bln] = bulan.split("-").map(Number);
  const bulanSah = Number.isFinite(thn) && Number.isFinite(bln) && bln >= 1 && bln <= 12;
  const jmlHari = bulanSah ? new Date(thn, bln, 0).getDate() : 0;
  const awalMinggu = bulanSah ? (new Date(thn, bln - 1, 1).getDay() + 6) % 7 : 0;
  /* tanggal terpilih; kalau belum ada pilihan, pakai laporan terbaru di bulan ini */
  const tglAktif = tglPilih && peta.has(tglPilih) ? tglPilih : (grup.length ? grup[grup.length - 1][0] : "");
  const listAktif: ItemRiwayat[] = tglAktif ? (peta.get(tglAktif) || []) : [];

  return (
    <Card style={{ marginTop: 12 }}>
      <SH label={judul} />
      <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", marginBottom: 8 }}>
        <Btn sm outline onClick={() => geser(-1)}>◀</Btn>
        <input style={{ ...iS, width: "auto", flex: 1, minWidth: 150 }} type="month" value={bulan} onChange={(e) => ganti(e.target.value)} />
        <Btn sm outline onClick={() => geser(1)}>▶</Btn>
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8, flexWrap: "wrap", marginBottom: 8 }}>
        <div style={{ fontWeight: 800 }}>{labelBulan(bulan)} • {daftarBulan.length} laporan</div>
        {onExcelBulan && daftarBulan.length > 0 && (
          <Btn sm outline onClick={() => onExcelBulan(daftarBulan, bulan)} disabled={sibuk}>⬇ Excel satu bulan</Btn>
        )}
      </div>
      {!daftarBulan.length && <div style={{ color: C.tL, fontSize: 13 }}>{kosongTeks || "Belum ada laporan tersimpan di bulan ini."}</div>}

      {bulanSah && (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 6 }}>
          {HARI_SINGKAT.map((h) => (
            <div key={h} style={{ textAlign: "center", fontSize: 11, fontWeight: 700, color: "#64748B" }}>{h}</div>
          ))}
          {Array.from({ length: awalMinggu }, (_, i) => <div key={"k" + i} />)}
          {Array.from({ length: jmlHari }, (_, i) => {
            const d = i + 1;
            const tgl = bulan + "-" + String(d).padStart(2, "0");
            const ada = peta.get(tgl);
            const aktif = tgl === tglAktif;
            return (
              <button
                key={tgl}
                type="button"
                disabled={!ada}
                onClick={() => setTglPilih(tgl)}
                title={ada ? ada.length + " laporan" : "Tidak ada laporan"}
                style={{
                  position: "relative", padding: "10px 0", borderRadius: 10, fontSize: 14, fontFamily: "inherit",
                  fontWeight: ada ? 800 : 500, cursor: ada ? "pointer" : "default",
                  border: aktif ? "2px solid " + C.p : ada ? "1px solid #6EE7B7" : "1px solid #E2E8F0",
                  background: aktif ? "#E0F2FE" : ada ? "#ECFDF5" : "#F8FAFC",
                  color: ada ? "#166534" : "#94A3B8",
                  textDecoration: tgl === hariIni ? "underline" : "none",
                }}
              >
                {d}
                {ada && ada.length > 1 && <span style={{ position: "absolute", top: 1, right: 4, fontSize: 9, fontWeight: 800, color: "#C2410C" }}>{ada.length}</span>}
              </button>
            );
          })}
        </div>
      )}

      {tglAktif && (
        <div style={{ marginTop: 12, border: "2px solid " + C.p, borderRadius: 12, padding: "10px 12px" }}>
          <div style={{ fontWeight: 800, fontSize: 14, marginBottom: 2 }}>{namaHari(tglAktif)}, {fmtTgl(tglAktif)}</div>
          {listAktif.map((it, idx) => (
            <div key={it.id} style={{ padding: "8px 0", borderTop: idx ? "1px dashed #CBD5E1" : "none", marginTop: idx ? 6 : 0 }}>
              {(it.label || it.ket) && (
                <div style={{ fontSize: 13, marginBottom: 6 }}>{it.label && <b>{it.label}</b>}{it.label && it.ket ? " • " : ""}{it.ket}</div>
              )}
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                <Btn sm outline onClick={() => onBuka(it)} disabled={sibuk}>Buka</Btn>
                <Btn sm outline onClick={() => onUnduh(it, "excel")} disabled={sibuk}>⬇ Excel</Btn>
                <Btn sm outline onClick={() => onUnduh(it, "word")} disabled={sibuk}>⬇ Word</Btn>
                <Btn sm outline onClick={() => onUnduh(it, "pdf")} disabled={sibuk}>⬇ PDF</Btn>
                {role === "admin" && onHapus && <Btn sm outline color={C.d} onClick={() => onHapus(it)}>Hapus</Btn>}
              </div>
            </div>
          ))}
        </div>
      )}
    </Card>
  );
}
