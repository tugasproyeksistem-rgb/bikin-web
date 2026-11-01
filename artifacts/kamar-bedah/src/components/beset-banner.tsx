/* src/components/beset-banner.tsx
   Pengumuman di tab Home (di atas "Operasi hari ini"): beset jadwal perawat dibuka / ditutup.
   Membaca kb_beset_jadwal bulan ini dan bulan depan, segar tiap 1 menit dan saat jendela fokus. */

import { useState, useEffect } from "react";
import { SUPA_CLIENT } from "../services/supabase";
import { isoLokal } from "../utils/laporan-harian";

const NAMA_BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
const lb = (b: string) => { const [y, m] = String(b).split("-").map(Number); return `${NAMA_BULAN[(m || 1) - 1]} ${y}`; };

export function BesetBanner({ setTab }: any) {
  const [rows, setRows] = useState<any[]>([]);

  useEffect(() => {
    let batal = false;
    const muat = async () => {
      try {
        const now = new Date();
        const nx = isoLokal(new Date(now.getFullYear(), now.getMonth() + 1, 1)).slice(0, 7);
        const r: any = await SUPA_CLIENT.from("kb_beset_jadwal").select("data").eq("id", "beset_" + nx);
        if (!batal && !r.error) setRows((r.data || []).map((x: any) => x.data).filter(Boolean));
      } catch { /* offline: banner tidak tampil */ }
    };
    muat();
    const t = setInterval(muat, 60000);
    const onFocus = () => { muat(); };
    window.addEventListener("focus", onFocus);
    return () => { batal = true; clearInterval(t); window.removeEventListener("focus", onFocus); };
  }, []);

  const dibuka = rows.filter((x) => x.buka).sort((a, b) => String(a.bulan).localeCompare(String(b.bulan)));
  const ditutup = rows
    .filter((x) => !x.buka && x.updated_at && Date.now() - new Date(x.updated_at).getTime() < 14 * 864e5)
    .sort((a, b) => String(b.bulan).localeCompare(String(a.bulan)));
  if (!dibuka.length && !ditutup.length) return null;

  const buka = dibuka.length > 0;
  const bulanTeks = (buka ? dibuka : [ditutup[0]]).map((x) => lb(x.bulan).toUpperCase()).join(" DAN ");
  const warna = buka ? "#DC2626" : "#6D28D9";

  return (
    <div>
      <style>{`
        @keyframes kbBlinkText{0%,100%{opacity:1}50%{opacity:.12}}
        @keyframes kbBlinkBoxMerah{0%,100%{border-color:#DC2626;box-shadow:0 0 0 0 rgba(220,38,38,.85)}50%{border-color:#FCA5A5;box-shadow:0 0 22px 8px rgba(220,38,38,.25)}}
        @keyframes kbBlinkBoxBiru{0%,100%{border-color:#6D28D9;box-shadow:0 0 0 0 rgba(37,99,235,.85)}50%{border-color:#C4B5FD;box-shadow:0 0 22px 8px rgba(37,99,235,.25)}}
        .kb-beset-teks{animation:kbBlinkText 1s ease-in-out infinite}
        .kb-beset-merah{animation:kbBlinkBoxMerah 1.1s ease-in-out infinite}
        .kb-beset-biru{animation:kbBlinkBoxBiru 1.1s ease-in-out infinite}
        @media (prefers-reduced-motion: reduce){.kb-beset-teks,.kb-beset-merah,.kb-beset-biru{animation:none}}
      `}</style>
      <div role="button" tabIndex={0}
        className={buka ? "kb-beset-merah" : "kb-beset-biru"}
        onClick={() => setTab && setTab("beset_jadwal")}
        onKeyDown={(e) => { if (e.key === "Enter" && setTab) setTab("beset_jadwal"); }}
        style={{ border: "4px solid " + warna, borderRadius: 14, padding: "14px 12px", marginBottom: 12, textAlign: "center", cursor: "pointer", background: buka ? "#FEF2F2" : "#F5F3FF" }}>
        <div className="kb-beset-teks" style={{ fontSize: 22, fontWeight: 900, lineHeight: 1.3, color: warna }}>
          {buka ? `📢 BESETAN BULAN ${bulanTeks} SUDAH DIBUKA!` : `🔒 BESETAN BULAN ${bulanTeks} SUDAH DITUTUP!`}
        </div>
        <div style={{ fontSize: 15, fontWeight: 700, marginTop: 6, color: "#111827", lineHeight: 1.4 }}>
          {buka
            ? "Yuk, segera isi besetan jadwal kamu sebelum ditutup. Klik di sini untuk mengisi 👉"
            : "Terima kasih sudah mengisi besetan tepat waktu 🙏 Jadwal sedang disusun, pantau terus ya!"}
        </div>
      </div>
    </div>
  );
}
