/* src/components/monitoring-pengingat.tsx
   Pengingat HARI INI di Home untuk pemantauan suhu & kelembaban.
   Slot 07:00 (shift pagi) dan 14:00 (shift siang). Tampil mulai jam slot kalau ada ruangan yang belum terisi;
   hilang sendiri begitu semua ruangan terisi. Membaca data yang sama dengan tab Monitoring (monitoringEntries). */
import { useState, useEffect } from "react";
import { MON_ROOMS } from "../constants/monitoring";
import { isoLokal } from "../utils/laporan-harian";

const SLOT = [
  { jam: "07:00", menit: 7 * 60, label: "Shift pagi" },
  { jam: "14:00", menit: 14 * 60, label: "Shift siang" },
];

/* tanggal tersimpan bisa "dd/mm/yyyy", "dd / mm / yyyy", atau ISO "yyyy-mm-dd" -> jadikan ISO */
function tglIso(t: any): string {
  const s = String(t ?? "").replace(/\s+/g, "");
  const a = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (a) return a[1] + "-" + a[2] + "-" + a[3];
  const b = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})/);
  if (b) return b[3] + "-" + b[2].padStart(2, "0") + "-" + b[1].padStart(2, "0");
  return "";
}

/* "07:00" / "07.00" / "7:00" -> "07:00" */
function jamNorm(j: any): string {
  const m = String(j ?? "").trim().replace(".", ":").match(/^(\d{1,2}):(\d{2})/);
  return m ? m[1].padStart(2, "0") + ":" + m[2] : "";
}

export function MonitoringPengingat({ entries, setTab }: any) {
  const [, setTick] = useState<number>(0);
  const [tunggu, setTunggu] = useState<boolean>(true);

  useEffect(() => {
    const segarkan = () => setTick((x) => x + 1);
    const t = setInterval(segarkan, 30000);
    const w = setTimeout(() => setTunggu(false), 3000);
    window.addEventListener("focus", segarkan);
    return () => { clearInterval(t); clearTimeout(w); window.removeEventListener("focus", segarkan); };
  }, []);

  const list: any[] = Array.isArray(entries) ? entries : [];
  /* data masih dimuat dari server: tunggu sebentar supaya pengingat tidak berkedip palsu */
  if (tunggu && list.length === 0) return null;

  const now = new Date();
  const hari = isoLokal(now);
  const menitSekarang = now.getHours() * 60 + now.getMinutes();

  const kurang = SLOT
    .filter((s) => menitSekarang >= s.menit)
    .map((s) => {
      const ruang = MON_ROOMS.filter(
        (r: string) =>
          !list.some(
            (e: any) =>
              e && e.ruang === r && tglIso(e.tanggal) === hari && jamNorm(e.jam) === s.jam &&
              Number(e.suhu) > 0 && Number(e.kelembaban) > 0,
          ),
      );
      return { s, ruang };
    })
    .filter((z) => z.ruang.length > 0);

  if (kurang.length === 0) return null;

  return (
    <div role="alert" style={{ border: "3px solid #0369A1", borderRadius: 14, background: "#EFF6FF", padding: "12px 14px", marginBottom: 14 }}>
      <div style={{ fontSize: 16, fontWeight: 900, color: "#075985" }}>Pengingat: isi pemantauan suhu &amp; kelembaban</div>
      <div style={{ marginTop: 6, fontSize: 14, color: "#0C4A6E", lineHeight: 1.6 }}>
        {kurang.map((z) => (
          <div key={z.s.jam}>
            <b>{z.s.label} - jam {z.s.jam}</b> <span style={{ color: "#075985" }}>(belum diisi: {z.ruang.join(", ")})</span>
          </div>
        ))}
      </div>
      {setTab && (
        <div style={{ marginTop: 10 }}>
          <button type="button" onClick={() => setTab("monitoring")} style={{ border: "none", background: "#0369A1", color: "#fff", borderRadius: 8, padding: "7px 14px", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>
            Buka Monitoring Suhu &rarr;
          </button>
        </div>
      )}
    </div>
  );
}
