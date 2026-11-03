/* src/components/riwayat-bulanan.tsx
   Riwayat laporan per bulan: tersusun per tanggal, tiap baris punya Buka, Excel, Word, PDF. */

import { useState, useMemo } from "react";
import { C, iS } from "../constants";
import { Card, SH, Btn } from "./primitives";
import { fmtTgl, namaHari } from "../utils/laporan-harian";

const NAMA_BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
const labelBulan = (b: string) => { const [y, m] = b.split("-").map(Number); return (NAMA_BULAN[(m || 1) - 1] || "") + " " + y; };

export interface ItemRiwayat { id: string; tanggal: string; label?: string; ket?: string; urut?: number; data: any; }

export function RiwayatBulanan({ judul = "Riwayat bulanan", items, role, kosongTeks, sibuk, onBuka, onUnduh, onHapus, onExcelBulan }: any) {
  const [pilih, setPilih] = useState("");
  const semua: ItemRiwayat[] = items || [];
  const bulanAda = useMemo(() => Array.from(new Set(semua.map((x) => String(x.tanggal).slice(0, 7)))).sort(), [semua]);
  const sekarang = (() => { const n = new Date(); return n.getFullYear() + "-" + String(n.getMonth() + 1).padStart(2, "0"); })();
  const bulan = pilih || bulanAda[bulanAda.length - 1] || sekarang;
  const geser = (n: number) => {
    const [y, m] = bulan.split("-").map(Number);
    const d = new Date(y, m - 1 + n, 1);
    setPilih(d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0"));
  };
  const daftarBulan = useMemo(() => semua
    .filter((x) => String(x.tanggal).startsWith(bulan))
    .sort((a, b) => String(a.tanggal).localeCompare(String(b.tanggal)) || (a.urut ?? 0) - (b.urut ?? 0)), [semua, bulan]);
  const grup = useMemo(() => {
    const m = new Map<string, ItemRiwayat[]>();
    daftarBulan.forEach((x) => { const k = String(x.tanggal); m.set(k, [...(m.get(k) || []), x]); });
    return Array.from(m.entries());
  }, [daftarBulan]);

  return (
    <Card style={{ marginTop: 12 }}>
      <SH label={judul} />
      <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", marginBottom: 8 }}>
        <Btn sm outline onClick={() => geser(-1)}>◀</Btn>
        <input style={{ ...iS, width: "auto", flex: 1, minWidth: 150 }} type="month" value={bulan} onChange={(e) => setPilih(e.target.value)} />
        <Btn sm outline onClick={() => geser(1)}>▶</Btn>
      </div>
      <div style={{ fontWeight: 800, marginBottom: 6 }}>{labelBulan(bulan)} • {daftarBulan.length} laporan</div>
      {onExcelBulan && daftarBulan.length > 0 && (
        <div style={{ marginBottom: 8 }}><Btn sm outline onClick={() => onExcelBulan(daftarBulan, bulan)} disabled={sibuk}>⬇ Excel satu bulan</Btn></div>
      )}
      {!daftarBulan.length && <div style={{ color: C.tL, fontSize: 13 }}>{kosongTeks || "Belum ada laporan tersimpan di bulan ini."}</div>}
      {grup.map(([tgl, list]) => (
        <div key={tgl} style={{ border: "1px solid " + C.b, borderRadius: 10, padding: "8px 10px", marginBottom: 8 }}>
          <div style={{ fontWeight: 800, fontSize: 14 }}>{namaHari(tgl)}, {fmtTgl(tgl)}</div>
          {list.map((it) => (
            <div key={it.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8, flexWrap: "wrap", padding: "6px 0", borderTop: "1px dashed " + C.b, marginTop: 6 }}>
              <div style={{ fontSize: 13, minWidth: 110 }}>{it.label && <b>{it.label}</b>}{it.label && it.ket ? " • " : ""}{it.ket}</div>
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
      ))}
    </Card>
  );
}
