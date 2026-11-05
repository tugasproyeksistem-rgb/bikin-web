/* src/components/home-ringkas.tsx
   Kartu ringkasan Home (kaca biru muda): Terjadwal, Berlangsung, Selesai, Cyto hari ini. */

import { ClipboardList, Play, CheckCircle2, AlertTriangle } from "lucide-react";
import { todayDate } from "../utils";

export function HomeRingkas({ ops }: any) {
  const hari = (ops || []).filter((o: any) => o && o.date === todayDate() && o.status !== "batal");
  const n = (f: (o: any) => boolean) => hari.filter(f).length;
  const kartu = [
    { l: "Terjadwal", v: n((o) => o.status === "scheduled"), Ik: ClipboardList, g: "linear-gradient(135deg,#38BDF8,#0284C7)", sh: "rgba(2,132,199,.35)" },
    { l: "Berlangsung", v: n((o) => o.status === "ongoing"), Ik: Play, g: "linear-gradient(135deg,#FBBF24,#D97706)", sh: "rgba(217,119,6,.35)" },
    { l: "Selesai", v: n((o) => o.status === "done"), Ik: CheckCircle2, g: "linear-gradient(135deg,#4ADE80,#16A34A)", sh: "rgba(22,163,74,.35)" },
    { l: "Cyto", v: n((o) => o.opType === "cyto"), Ik: AlertTriangle, g: "linear-gradient(135deg,#F87171,#DC2626)", sh: "rgba(220,38,38,.35)" },
  ];
  return (
    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(140px,1fr))", gap: 12, marginBottom: 16 }}>
      {kartu.map((k) => (
        <div key={k.l} className="kb-kaca" style={{ padding: "14px 14px", display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ width: 42, height: 42, borderRadius: 14, background: k.g, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, boxShadow: "0 6px 14px " + k.sh }}><k.Ik size={20} strokeWidth={2.4} /></div>
          <div>
            <div style={{ fontSize: 28, fontWeight: 900, lineHeight: 1, color: "#0B0F14" }}>{k.v}</div>
            <div style={{ fontSize: 12, fontWeight: 800, color: "#1F2937", marginTop: 3 }}>{k.l}</div>
          </div>
        </div>
      ))}
    </div>
  );
}
