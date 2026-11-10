/* src/components/implan-pengingat.tsx
   Pengingat HARI H di Home: pasien implant yang dioperasi hari ini tetapi alamat dan/atau nomor HP-nya belum terisi
   (baik di data pasien tab Daftar maupun di isian manual tab IMPLANT/HERNIA MESH). Hilang otomatis setelah terisi.
   Tampil mulai 1 Januari 2027 (atau saat uji implant diaktifkan). */
import { useState, useEffect, useMemo } from "react";
import { SUPA_CLIENT } from "../services/supabase";
import { maskName } from "../utils";
import { implantAktif } from "../constants/fitur-tanggal";
import { isoLokal } from "../utils/laporan-harian";

const TUJUAN: Record<string, { tab: string; nama: string }> = {
  tulang: { tab: "implant_tulang", nama: "IMPLANT TULANG" },
  mata: { tab: "implant_mata", nama: "IMPLANT MATA" },
  hernia: { tab: "implant_hernia", nama: "HERNIA MESH" },
};

export function ImplanPengingat({ ops, setTab, privacyMode }: any) {
  const [rows, setRows] = useState<any[]>([]);
  const [siap, setSiap] = useState<boolean>(false);
  const [, setTick] = useState<number>(0);
  const aktif = implantAktif();
  const hari = isoLokal(new Date());

  const kandidat = useMemo(
    () => (ops || []).filter((o: any) => o && o.date === hari && o.status !== "batal" && o.implant === "ya" && TUJUAN[String(o.kelompokImplant)]),
    [ops, hari],
  );
  const perlu = aktif && kandidat.length > 0;

  useEffect(() => {
    if (!perlu) return;
    let batal = false;
    const muat = async () => {
      try {
        const r: any = await SUPA_CLIENT.from("kb_implan_ok").select("data");
        if (!batal && !r.error) setRows((r.data || []).map((x: any) => x.data).filter(Boolean));
      } catch { /* offline: pakai data pasien saja */ }
      if (!batal) setSiap(true);
    };
    muat();
    const segarkan = () => { muat(); setTick((x) => x + 1); };
    const t = setInterval(segarkan, 60000);
    window.addEventListener("focus", segarkan);
    return () => { batal = true; clearInterval(t); window.removeEventListener("focus", segarkan); };
  }, [perlu]);

  if (!perlu || !siap) return null;

  const peta = new Map<string, any>();
  rows.forEach((r: any) => { if (r && r.opId) peta.set(String(r.opId), r); });
  const kurang = kandidat
    .map((o: any) => {
      const r = peta.get(String(o.id));
      const alamat = String(o.alamatPasien ?? "").trim() || String(r?.alamat ?? "").trim();
      const hp = String(o.noHp ?? "").trim() || String(r?.noTelp ?? "").trim();
      const x: string[] = [];
      if (!alamat) x.push("alamat");
      if (!hp) x.push("nomor HP");
      return { o, x };
    })
    .filter((z: any) => z.x.length > 0);
  if (kurang.length === 0) return null;

  const tujuan: string[] = Array.from(new Set<string>(kurang.map((z: any) => String(z.o.kelompokImplant))));
  return (
    <div role="alert" style={{ border: "3px solid #EA580C", borderRadius: 14, background: "#FFF7ED", padding: "12px 14px", marginBottom: 14 }}>
      <div style={{ fontSize: 16, fontWeight: 900, color: "#9A3412" }}>Hari ini: lengkapi alamat dan nomor HP pasien implant</div>
      <div style={{ marginTop: 6, fontSize: 14, color: "#431407", lineHeight: 1.6 }}>
        {kurang.slice(0, 6).map((z: any, i: number) => (
          <div key={i}>{privacyMode ? maskName(String(z.o.patient || "-"), true) : String(z.o.patient || "-")} <span style={{ color: "#9A3412" }}>(belum ada {z.x.join(" dan ")})</span></div>
        ))}
        {kurang.length > 6 && <div>dan {kurang.length - 6} pasien lain</div>}
      </div>
      {setTab && (
        <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginTop: 10 }}>
          {tujuan.map((k) => (
            <button key={k} type="button" onClick={() => setTab(TUJUAN[k].tab)} style={{ border: "none", background: "#EA580C", color: "#fff", borderRadius: 8, padding: "7px 14px", fontSize: 13, fontWeight: 800, cursor: "pointer", fontFamily: "inherit" }}>Buka {TUJUAN[k].nama} →</button>
          ))}
        </div>
      )}
    </div>
  );
}
