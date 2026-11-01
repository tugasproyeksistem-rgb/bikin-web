/* src/components/statistik-bulanan.tsx
   Blok statistik per bulan: total operasi + Top dokter bedah + Top dokter
   anestesi. Dipakai di tab Statistik (admin) dan di Home (admin + perawat).
   Operasi berstatus "batal" tidak dihitung. */

import { useState } from "react";
import { C } from "../constants";
import { Card, SH, MiniBar } from "./primitives";

const NAMA_BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
const WARNA = ["#1F5A52", "#4338CA", "#E65100", "#7B1FA2", "#2E7D32", "#C62828", "#00838F", "#E91E63"];

function bulanIni(): string {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}
function geserBulan(ym: string, n: number): string {
  const [y, m] = ym.split("-").map(Number);
  const d = new Date(y, m - 1 + n, 1);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}
function labelBulan(ym: string): string {
  const [y, m] = ym.split("-").map(Number);
  return `${NAMA_BULAN[m - 1] || ""} ${y}`;
}
function topN(list: any[], field: string, n: number) {
  const peta: Record<string, number> = {};
  list.forEach((o: any) => {
    const nama = String(o?.[field] || "").trim();
    if (nama) peta[nama] = (peta[nama] || 0) + 1;
  });
  return Object.entries(peta).sort((a, b) => b[1] - a[1]).slice(0, n)
    .map(([label, value], i) => ({ label, value, color: WARNA[i % WARNA.length] }));
}

export function StatistikBulanan({ ops, archive }: { ops: any[]; archive?: any[] }) {
  const [bulan, setBulan] = useState<string>(bulanIni());
  const semua = [...(ops || []), ...((archive || []).flatMap((a: any) => a?.ops || []))];
  const bulanIniOps = semua.filter((o: any) => o && typeof o.date === "string" && o.date.startsWith(bulan));
  const aktif = bulanIniOps.filter((o: any) => o.status !== "batal");
  const batal = bulanIniOps.length - aktif.length;
  const selesai = aktif.filter((o: any) => o.status === "done").length;
  const dataBedah = topN(aktif, "surgeon", 8);
  const dataAnestesi = topN(aktif, "anesthesiologist", 6);

  const tombol: React.CSSProperties = { width: 40, height: 40, borderRadius: 10, border: `1px solid ${C.b}`, background: "#fff", color: C.p, fontSize: 16, fontWeight: 800, cursor: "pointer", fontFamily: "inherit" };

  return (
    <>
      <Card>
        <SH label="📅 Operasi per Bulan" />
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, marginBottom: 12 }}>
          <button type="button" style={tombol} aria-label="Bulan sebelumnya" onClick={() => setBulan(geserBulan(bulan, -1))}>◀</button>
          <div style={{ textAlign: "center" }}>
            <div style={{ fontSize: 16, fontWeight: 800, color: C.t }}>{labelBulan(bulan)}</div>
            {bulan !== bulanIni() && (
              <button type="button" onClick={() => setBulan(bulanIni())} style={{ background: "none", border: "none", color: C.p, fontSize: 11, cursor: "pointer", padding: 0, fontFamily: "inherit" }}>↺ bulan ini</button>
            )}
          </div>
          <button type="button" style={tombol} aria-label="Bulan berikutnya" onClick={() => setBulan(geserBulan(bulan, 1))}>▶</button>
        </div>
        <div style={{ background: C.pBg, borderRadius: 12, padding: "14px 16px", textAlign: "center", border: `1px solid ${C.p}33` }}>
          <div style={{ fontSize: 32, fontWeight: 800, color: C.p }}>{aktif.length}</div>
          <div style={{ fontSize: 11, color: C.p, fontWeight: 600, marginTop: 2 }}>Total operasi bulan ini</div>
        </div>
        <div style={{ fontSize: 11, color: C.tL, marginTop: 8, textAlign: "center" }}>
          Selesai {selesai}{batal > 0 ? ` · Batal ${batal} (tidak dihitung)` : ""}
        </div>
      </Card>

      {dataBedah.length > 0 && (
        <Card>
          <SH label="Operasi per Dokter Bedah (Top 8)" />
          <MiniBar data={dataBedah} />
        </Card>
      )}
      {dataAnestesi.length > 0 && (
        <Card>
          <SH label="Operasi per Dokter Anestesi (Top 6)" />
          <MiniBar data={dataAnestesi} />
        </Card>
      )}
      {aktif.length === 0 && (
        <div style={{ fontSize: 13, color: C.tL, textAlign: "center", padding: "4px 0 12px" }}>Belum ada operasi pada {labelBulan(bulan)}.</div>
      )}
    </>
  );
}
