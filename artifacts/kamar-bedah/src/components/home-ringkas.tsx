/* src/components/home-ringkas.tsx
   Kartu ringkasan Home (kaca, gradasi ungu-biru): Terjadwal, Berlangsung, Selesai, Cyto hari ini. */

import { todayDate } from "../utils";

export function HomeRingkas({ ops }: any) {
  const hari = (ops || []).filter((o: any) => o && o.date === todayDate() && o.status !== "batal");
  const n = (f: (o: any) => boolean) => hari.filter(f).length;
  const kartu = [
    { l: "Terjadwal", v: n((o) => o.status === "scheduled"), i: "📋", g: "linear-gradient(135deg,#6366F1,#8B5CF6)" },
    { l: "Berlangsung", v: n((o) => o.status === "ongoing"), i: "▶", g: "linear-gradient(135deg,#3B82F6,#6366F1)" },
    { l: "Selesai", v: n((o) => o.status === "done"), i: "✓", g: "linear-gradient(135deg,#4F46E5,#3B82F6)" },
    { l: "Cyto", v: n((o) => o.opType === "cyto"), i: "!", g: "linear-gradient(135deg,#7C3AED,#6366F1)" },
  ];
  return (
    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(140px,1fr))", gap: 12, marginBottom: 16 }}>
      {kartu.map((k) => (
        <div key={k.l} className="kb-kaca" style={{ padding: "14px 14px", display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ width: 42, height: 42, borderRadius: 14, background: k.g, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, fontWeight: 900, flexShrink: 0, boxShadow: "0 6px 14px rgba(99,102,241,.35)" }}>{k.i}</div>
          <div>
            <div style={{ fontSize: 28, fontWeight: 900, lineHeight: 1, background: k.g, WebkitBackgroundClip: "text", backgroundClip: "text", color: "transparent" }}>{k.v}</div>
            <div style={{ fontSize: 12, fontWeight: 700, color: "#5B5FA8", marginTop: 3 }}>{k.l}</div>
          </div>
        </div>
      ))}
    </div>
  );
}
