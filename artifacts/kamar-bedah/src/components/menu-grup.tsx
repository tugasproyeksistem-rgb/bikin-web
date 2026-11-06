/* Menu grup: tombol grup (Utama, Laporan, dst) yang membuka kotak berisi
   ikon tab di dalam grup. Pemilihan tab tetap lewat onKlik (handler lama
   handleTabClickEvt) sehingga log navigasi dan pemeriksaan akses tidak berubah. */
import { useState, useEffect } from "react";
import type { MouseEvent } from "react";
import { createPortal } from "react-dom";
import { TabIcon } from "./icons";

type Tab = { k: string; l: string };
type Grup = { id: string; l: string; s: string; ikon: string; isi: string[] };
type Props = {
  tabs: Tab[];
  aktif: string;
  onKlik: (e: MouseEvent<HTMLButtonElement>) => void;
  varian: "sidebar" | "bar";
  badge?: Record<string, number>;
};

const GRUP: Grup[] = [
  { id: "utama", l: "Utama", s: "Utama", ikon: "home", isi: ["home", "jadwal", "daftar"] },
  { id: "laporan", l: "Laporan", s: "Laporan", ikon: "laporan", isi: ["laporan", "laporan_harian", "jadwal_ok", "beset_jadwal"] },
  { id: "staf", l: "Staf & Komunikasi", s: "Staf", ikon: "staf", isi: ["lembur", "monitoring", "wa", "pengumuman", "staf"] },
  { id: "kelola", l: "Pengelolaan", s: "Kelola", ikon: "arsip", isi: ["roster_gen", "arsip", "statistik"] },
];

const svgP = { width: 18, height: 18, viewBox: "0 0 24 24", fill: "none", stroke: "currentColor", strokeWidth: 2, strokeLinecap: "round" as const, strokeLinejoin: "round" as const };

function Ikon({ k }: { k: string }) {
  switch (k) {
    case "laporan_harian":
      return <svg {...svgP}><rect x="5" y="4" width="14" height="17" rx="2"/><path d="M9 4V3h6v1"/><path d="M9 11h6M9 15h6"/></svg>;
    case "beset_jadwal":
      return <svg {...svgP}><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 10h18M9 4v16"/></svg>;
    case "jadwal_ok":
      return <svg {...svgP}><path d="M9 6h11M9 12h11M9 18h11"/><path d="m3.5 6 1 1 2-2M3.5 12l1 1 2-2M3.5 18l1 1 2-2"/></svg>;
    case "pengumuman":
      return <svg {...svgP}><path d="M3 11v2a1 1 0 0 0 1 1h2l5 4V6L6 10H4a1 1 0 0 0-1 1Z"/><path d="M15 9a4 4 0 0 1 0 6M18 6.5a8 8 0 0 1 0 11"/></svg>;
    default:
      return <TabIcon k={k} />;
  }
}

const Panah = ({ turun }: { turun: boolean }) => (
  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ flexShrink: 0 }}>
    {turun ? <path d="m6 9 6 6 6-6" /> : <path d="m9 6 6 6-6 6" />}
  </svg>
);

export function MenuGrup({ tabs, aktif, onKlik, varian, badge }: Props) {
  const [buka, setBuka] = useState<string | null>(null);
  const [pos, setPos] = useState<{ top: number; left: number }>({ top: 0, left: 0 });

  useEffect(() => {
    if (!buka) return;
    const tutup = () => setBuka(null);
    const esc = (e: KeyboardEvent) => { if (e.key === "Escape") tutup(); };
    window.addEventListener("resize", tutup);
    window.addEventListener("keydown", esc);
    return () => {
      window.removeEventListener("resize", tutup);
      window.removeEventListener("keydown", esc);
    };
  }, [buka]);

  const ada = new Set(tabs.map((t) => t.k));
  const dipakai = new Set<string>();
  const grup: Grup[] = GRUP.map((g) => {
    const isi = g.isi.filter((k) => ada.has(k));
    isi.forEach((k) => dipakai.add(k));
    return { ...g, isi };
  });
  const sisa = tabs.filter((t) => !dipakai.has(t.k)).map((t) => t.k);
  if (sisa.length > 0) grup.push({ id: "lainnya", l: "Lainnya", s: "Lainnya", ikon: sisa[0], isi: sisa });
  const tampil = grup.filter((g) => g.isi.length > 0);
  const label = (k: string): string => { const t = tabs.find((x) => x.k === k); return t ? t.l : k; };

  const bukaGrup = (id: string, e: MouseEvent<HTMLButtonElement>) => {
    if (buka === id) { setBuka(null); return; }
    const r = e.currentTarget.getBoundingClientRect();
    const lebar = 272, tinggi = 240;
    const vw = window.innerWidth, vh = window.innerHeight;
    if (varian === "bar") setPos({ top: r.bottom + 6, left: Math.max(8, Math.min(r.left, vw - lebar - 8)) });
    else setPos({ top: Math.max(8, Math.min(r.top, vh - tinggi - 8)), left: Math.max(8, Math.min(r.right + 8, vw - lebar - 8)) });
    setBuka(id);
  };

  const gBuka = tampil.find((g) => g.id === buka) || null;

  return (
    <>
      {tampil.map((g) => {
        const aktifG = g.isi.indexOf(aktif) >= 0;
        const tunggal = g.isi.length === 1;
        const jml = g.isi.reduce((n, k) => n + ((badge && badge[k]) || 0), 0);
        const dataTab = tunggal ? g.isi[0] : undefined;
        const klik = tunggal ? onKlik : (e: MouseEvent<HTMLButtonElement>) => bukaGrup(g.id, e);
        if (varian === "sidebar") {
          return (
            <button key={g.id} type="button" data-tab={dataTab} onClick={klik}
              className={"kb-sidebar-btn" + (aktifG ? " active" : "")}>
              <span style={{ flexShrink: 0, display: "flex" }}><Ikon k={g.ikon} /></span>
              <span>{g.l}</span>
              {jml > 0 && <span style={{ marginLeft: "auto", background: "#F44336", color: "#fff", fontSize: 9, fontWeight: 700, padding: "2px 6px", borderRadius: 10 }}>{jml}</span>}
              {!tunggal && <span style={{ marginLeft: jml > 0 ? 6 : "auto", display: "flex" }}><Panah turun={false} /></span>}
            </button>
          );
        }
        return (
          <button key={g.id} type="button" data-tab={dataTab} onClick={klik} className="kb-tab-btn"
            style={{ border: "none", background: aktifG ? "rgba(14,165,233,.18)" : "none", color: aktifG ? "#0B0F14" : "#1F2937", fontWeight: aktifG ? 900 : 600, cursor: "pointer", borderBottom: aktifG ? "3px solid #0284C7" : "3px solid transparent", borderRadius: "8px 8px 0 0", transition: "all .15s", fontFamily: "inherit", display: "inline-flex", alignItems: "center", gap: 6 }}>
            <Ikon k={g.ikon} />{g.s}
            {!tunggal && <Panah turun={true} />}
          </button>
        );
      })}
      {gBuka && createPortal(
        <>
          <div onClick={() => setBuka(null)} style={{ position: "fixed", top: 0, left: 0, right: 0, bottom: 0, zIndex: 9998, background: "rgba(15,23,42,.18)" }} />
          <div role="menu" style={{ position: "fixed", top: pos.top, left: pos.left, zIndex: 9999, background: "#fff", border: "1.5px solid #64748B", borderRadius: 16, padding: 12, boxShadow: "0 12px 32px rgba(15,23,42,.28)", width: "max-content", maxWidth: "calc(100vw - 16px)" }}>
            <div style={{ fontSize: 11, fontWeight: 800, color: "#374151", letterSpacing: 1, textTransform: "uppercase", marginBottom: 8 }}>{gBuka.l}</div>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 8, maxWidth: 264 }}>
              {gBuka.isi.map((k) => {
                const on = aktif === k;
                const b = (badge && badge[k]) || 0;
                return (
                  <button key={k} type="button" data-tab={k} onClick={(e) => { onKlik(e); setBuka(null); }}
                    style={{ position: "relative", width: 80, padding: "10px 4px", display: "flex", flexDirection: "column", alignItems: "center", gap: 6, borderRadius: 12, cursor: "pointer", fontFamily: "inherit", fontSize: 11, fontWeight: on ? 900 : 700, lineHeight: 1.2, textAlign: "center", color: "#0B0F14", background: on ? "rgba(14,165,233,.18)" : "#fff", border: on ? "1.5px solid #0284C7" : "1.5px solid #CBD5E1" }}>
                    <Ikon k={k} />
                    <span>{label(k)}</span>
                    {b > 0 && <span style={{ position: "absolute", top: 4, right: 4, background: "#F44336", color: "#fff", fontSize: 9, fontWeight: 700, padding: "1px 5px", borderRadius: 10 }}>{b}</span>}
                  </button>
                );
              })}
            </div>
          </div>
        </>,
        document.body
      )}
    </>
  );
}
